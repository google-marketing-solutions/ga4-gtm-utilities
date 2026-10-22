/**
 * @fileoverview UrlFetchApp mock with two transports:
 *  - a JS `fetchHandler(req) => {code, body, headers}` (fake backends), and
 *  - `realHttp: true`: a synchronous real request through `curl`.
 *
 * Security: in realHttp mode the whole curl configuration (URL, headers incl.
 * Authorization, payload) is passed on stdin (`curl -K -`), never on the
 * command line, so tokens don't show up in `ps` output or error messages.
 * Request logs (`runtime.fetchLog`) redact credentials.
 */
'use strict';

const crypto = require('node:crypto');
const {spawnSync} = require('node:child_process');
const {defineHidden} = require('./uiMock');
const {BlobMock, toBuffer, toSignedBytes} = require('./servicesMock');

const SENSITIVE_HEADER_RE = /^(authorization|proxy-authorization|cookie|set-cookie|x-goog-api-key|x-api-key)$/i;
const DEFAULT_ERROR_TRUNCATE = 300;
const DEFAULT_LOG_PAYLOAD_LIMIT = 2000;

/**
 * Case-insensitive header lookup.
 * @param {!Object<string, string>} headers
 * @param {string} name
 * @return {string|undefined}
 */
function getHeader(headers, name) {
  const lower = name.toLowerCase();
  for (const [k, v] of Object.entries(headers || {})) if (k.toLowerCase() === lower) return v;
  return undefined;
}

/**
 * Returns a copy of `headers` with credentials redacted.
 * @param {!Object<string, *>} headers
 * @return {!Object<string, string>}
 */
function redactHeaders(headers) {
  const out = {};
  for (const [k, v] of Object.entries(headers || {})) {
    if (SENSITIVE_HEADER_RE.test(k)) {
      const scheme = /^(\w+)\s+/.exec(String(v));
      out[k] = scheme ? `${scheme[1]} <redacted>` : '<redacted>';
    } else {
      out[k] = String(v);
    }
  }
  return out;
}

/**
 * Truncates a string for logs.
 * @param {string} s
 * @param {number} limit
 * @return {string}
 */
function truncate(s, limit) {
  if (s.length <= limit) return s;
  return `${s.slice(0, limit)}... [${s.length - limit} more chars]`;
}

/**
 * HTTPResponse mock.
 */
class HTTPResponseMock {
  /**
   * @param {!Object} realm
   * @param {{code: number, bytes: !Buffer, headers: !Object<string, (string|!Array<string>)>}} res
   */
  constructor(realm, res) {
    defineHidden(this, {_realm: realm, _code: res.code, _bytes: res.bytes, _headers: res.headers});
  }

  getResponseCode() {
    return this._code;
  }

  getContentText(charset) {
    return this._bytes.toString(charset && /ascii|latin|8859/i.test(charset) ? 'latin1' : 'utf8');
  }

  getContent() {
    return this._realm.toContext(toSignedBytes(this._bytes));
  }

  /** @return {!Object<string, string>} Single-valued headers (last value wins). */
  getHeaders() {
    const out = {};
    for (const [k, v] of Object.entries(this._headers)) out[k] = Array.isArray(v) ? v[v.length - 1] : v;
    return this._realm.toContext(out);
  }

  /** @return {!Object<string, (string|!Array<string>)>} Multi-valued headers as arrays. */
  getAllHeaders() {
    const out = {};
    for (const [k, v] of Object.entries(this._headers)) out[k] = Array.isArray(v) && v.length === 1 ? v[0] : v;
    return this._realm.toContext(out);
  }

  getBlob() {
    const type = getHeader(this.getHeaders(), 'content-type') || null;
    return new BlobMock(this._realm, this._bytes, type ? String(type).split(';')[0] : null, null);
  }

  getAs(contentType) {
    return new BlobMock(this._realm, this._bytes, contentType, null);
  }

  toString() {
    return 'HTTPResponse';
  }
}

/**
 * Creates the UrlFetchApp mock.
 *
 * @param {!Object} env Runtime environment: {realm, options, fetchLog, warn,
 *     redact, addSecret, now, harnessErrors}.
 * @return {{service: !Object, fetchInternal: function(string, !Object=): !HTTPResponseMock,
 *     setHandler: function(*): void}}
 */
function createUrlFetchApp(env) {
  const realm = env.realm;
  const opts = env.options;
  let handler = opts.fetchHandler || null;

  /** @return {?function(!Object): *} */
  const resolveHandler = () => {
    if (!handler) return null;
    if (typeof handler === 'function') return handler;
    if (typeof handler.handle === 'function') return (req) => handler.handle(req);
    if (typeof handler.handler === 'function') return handler.handler;
    throw new Error('fetchHandler must be a function or an object with handle(req)');
  };

  /**
   * Normalizes (url, params) like Apps Script into a host-realm request.
   * @param {*} url
   * @param {?Object} params
   * @return {!Object}
   */
  function normalizeRequest(url, params) {
    if (typeof url !== 'string' || !/^https?:\/\/[^\s/]+/i.test(url)) {
      throw realm.error(`Invalid argument: ${url}`);
    }
    const p = params || {};
    const method = String(p.method || 'get').toUpperCase();
    const headers = {};
    for (const [k, v] of Object.entries(p.headers || {})) {
      headers[k] = String(v);
      if (SENSITIVE_HEADER_RE.test(k)) {
        const m = /^\w+\s+(.+)$/.exec(String(v));
        env.addSecret(m ? m[1] : String(v));
      }
    }
    let contentType = p.contentType !== undefined ? String(p.contentType) : getHeader(headers, 'content-type');
    let payloadText = null;
    let payloadBytes = null;
    const raw = p.payload;
    if (raw !== undefined && raw !== null) {
      if (method === 'GET' || method === 'HEAD') {
        env.warn('payload-ignored', `A payload was passed to a ${method} request (${url}); ` +
            'Apps Script does not send it.', {url});
      } else if (typeof raw === 'string') {
        payloadText = raw;
      } else if (Array.isArray(raw) || raw instanceof BlobMock || (raw && typeof raw.getBytes === 'function')) {
        payloadBytes = toBuffer(raw);
        if (!contentType && raw.getContentType) contentType = raw.getContentType() || undefined;
      } else if (typeof raw === 'object') {
        if (contentType && /json/i.test(contentType)) {
          env.warn('object-payload-with-json-content-type',
              `payload for ${method} ${url} is a JavaScript object, so Apps Script form-encodes it even ` +
              'though contentType is JSON. Did you forget JSON.stringify()?', {url});
        }
        payloadText = Object.entries(raw)
            .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(String(v))}`).join('&');
      } else {
        payloadText = String(raw);
      }
      if (!contentType && (payloadText !== null || payloadBytes !== null)) {
        contentType = 'application/x-www-form-urlencoded';
      }
    }
    if (payloadText !== null) payloadBytes = Buffer.from(payloadText, 'utf8');
    const req = {
      url,
      method,
      headers,
      contentType: contentType || null,
      payload: payloadText,
      payloadBytes,
      muteHttpExceptions: Boolean(p.muteHttpExceptions),
      followRedirects: p.followRedirects !== false,
      validateHttpsCertificates: p.validateHttpsCertificates !== false,
      escaping: p.escaping !== false,
      timestamp: env.now(),
      /** @param {string} name @return {string|undefined} */
      header(name) {
        if (name.toLowerCase() === 'content-type' && req.contentType) return req.contentType;
        return getHeader(headers, name);
      },
      /** @return {*} Parsed JSON payload (throws on invalid JSON). */
      json() {
        return JSON.parse(payloadText === null ? '' : payloadText);
      },
    };
    return req;
  }

  /**
   * Normalizes a handler result into {code, bytes, headers}.
   * @param {*} res
   * @return {{code: number, bytes: !Buffer, headers: !Object}}
   */
  function normalizeResponse(res) {
    const code = Number(res.code ?? res.status ?? res.statusCode ?? 200);
    const headers = {...(res.headers || {})};
    let body = res.body ?? res.content ?? '';
    let bytes;
    if (Buffer.isBuffer(body) || body instanceof Uint8Array) {
      bytes = Buffer.from(body);
    } else if (typeof body === 'object') {
      bytes = Buffer.from(JSON.stringify(body, null, 2), 'utf8');
      if (!getHeader(headers, 'content-type')) headers['Content-Type'] = 'application/json; charset=UTF-8';
    } else {
      body = String(body);
      bytes = Buffer.from(body, 'utf8');
    }
    return {code, bytes, headers};
  }

  /**
   * Performs a synchronous HTTP request with curl (config passed on stdin).
   * @param {!Object} req
   * @return {{code: number, bytes: !Buffer, headers: !Object}}
   */
  function curlRequest(req) {
    const q = (s) => `"${String(s).replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\n/g, '\\n')
        .replace(/\r/g, '\\r').replace(/\t/g, '\\t').replace(/\v/g, '\\v')}"`;
    const marker = `__HARNESS_CURL_META_${crypto.randomUUID()}__`;
    const maxTime = Number(opts.httpTimeoutSeconds) || 120;
    const lines = [
      `url = ${q(req.url)}`,
      `request = ${q(req.method)}`,
      'silent',
      'show-error',
      `max-time = ${maxTime}`,
      `write-out = ${q(`${marker}%{http_code}\n%{header_json}`)}`,
    ];
    if (req.followRedirects) lines.push('location');
    if (!req.validateHttpsCertificates) lines.push('insecure');
    const sanitize = (s) => String(s).replace(/[\r\n]/g, ' ');
    // Transport-level headers for live runs (invisible to the script), e.g. a
    // quota project for gcloud user credentials.
    const extraHeaders = {...(opts.realHttpHeaders || {})};
    if (process.env.GTM_QUOTA_PROJECT && /\.googleapis\.com$/i.test(new URL(req.url).hostname)) {
      extraHeaders['X-Goog-User-Project'] = extraHeaders['X-Goog-User-Project'] || process.env.GTM_QUOTA_PROJECT;
    }
    let hasContentType = false;
    const allHeaders = {...extraHeaders, ...req.headers};
    for (const [k, v] of Object.entries(allHeaders)) {
      if (k.toLowerCase() === 'content-type') hasContentType = true;
      lines.push(`header = ${q(`${sanitize(k)}: ${sanitize(v)}`)}`);
    }
    if (req.payloadBytes) {
      if (!hasContentType && req.contentType) lines.push(`header = ${q(`Content-Type: ${sanitize(req.contentType)}`)}`);
      const text = req.payloadBytes.toString('utf8');
      if (!Buffer.from(text, 'utf8').equals(req.payloadBytes) || text.includes('\0')) {
        throw new Error('UrlFetchApp mock (realHttp): binary payloads are not supported.');
      }
      lines.push(`data-raw = ${q(text)}`);
    }
    const result = spawnSync(opts.curlPath || 'curl', ['-K', '-'], {
      input: `${lines.join('\n')}\n`,
      maxBuffer: 512 * 1024 * 1024,
      timeout: (maxTime + 15) * 1000,
      windowsHide: true,
    });
    if (result.error) {
      throw new Error(`UrlFetchApp mock (realHttp): could not run curl: ${env.redact(result.error.message)}`);
    }
    if (result.status !== 0) {
      const stderr = env.redact(String(result.stderr || '').trim());
      const messages = {6: 'DNS error', 7: 'Address unavailable', 28: 'Timeout', 35: 'SSL Error', 60: 'SSL Error'};
      const label = messages[result.status];
      const urlForMsg = env.redact(req.url);
      throw realm.error(label ? `${label}: ${urlForMsg}` : (stderr || `curl exited with ${result.status}`));
    }
    const stdout = result.stdout;
    const idx = stdout.lastIndexOf(Buffer.from(marker));
    if (idx < 0) throw new Error('UrlFetchApp mock (realHttp): could not parse curl output.');
    const meta = stdout.subarray(idx + Buffer.byteLength(marker)).toString('utf8');
    const nl = meta.indexOf('\n');
    const code = parseInt(nl >= 0 ? meta.slice(0, nl) : meta, 10);
    let headers = {};
    try {
      const parsed = JSON.parse(meta.slice(nl + 1));
      for (const [k, v] of Object.entries(parsed)) headers[k] = Array.isArray(v) && v.length === 1 ? v[0] : v;
    } catch (e) {
      headers = {};
    }
    return {code, bytes: stdout.subarray(0, idx), headers};
  }

  /**
   * Executes one normalized request through the configured transport.
   * @param {!Object} req
   * @return {{code: number, bytes: !Buffer, headers: !Object, mode: string}}
   */
  function transport(req) {
    const fn = resolveHandler();
    if (fn) {
      let res;
      try {
        res = fn(req);
      } catch (e) {
        env.harnessErrors.push({where: 'fetchHandler', url: env.redact(req.url), error: e});
        throw new Error(`HarnessError: fetchHandler threw for ${req.method} ${env.redact(req.url)}: ` +
            `${e && e.stack ? e.stack : e}`);
      }
      if (res && typeof res.then === 'function') {
        throw new Error('HarnessError: fetchHandler must be synchronous (it returned a Promise).');
      }
      if (res && res.networkError) {
        // Simulated transport failure, e.g. {networkError: 'Address unavailable'}.
        throw realm.error(`${res.networkError}: ${env.redact(req.url)}`);
      }
      if (res !== undefined && res !== null) return {...normalizeResponse(res), mode: 'handler'};
    }
    if (opts.realHttp) return {...curlRequest(req), mode: 'realHttp'};
    const message = `UrlFetchApp mock: no fetchHandler response and realHttp is off for ${req.method} ` +
        `${env.redact(req.url)}`;
    env.harnessErrors.push({where: 'UrlFetchApp', url: env.redact(req.url), error: new Error(message)});
    throw new Error(`HarnessError: ${message}`);
  }

  /**
   * Fetches and logs, returning the raw response record.
   * @param {!Object} req
   * @return {{code: number, bytes: !Buffer, headers: !Object}}
   */
  function execute(req) {
    const started = Date.now();
    const entry = {
      index: env.fetchLog.length,
      method: req.method,
      url: env.redact(req.url),
      status: null,
      durationMs: null,
      mode: null,
      requestHeaders: redactHeaders(req.headers),
      contentType: req.contentType,
      payload: req.payload === null ? null :
        truncate(env.redact(req.payload), opts.fetchLogPayloadLimit || DEFAULT_LOG_PAYLOAD_LIMIT),
      payloadBytes: req.payloadBytes ? req.payloadBytes.length : 0,
      responsePreview: null,
      responseBytes: 0,
      error: null,
    };
    env.fetchLog.push(entry);
    try {
      const res = transport(req);
      entry.status = res.code;
      entry.mode = res.mode;
      entry.responseBytes = res.bytes.length;
      entry.responsePreview = truncate(env.redact(res.bytes.toString('utf8')), 500);
      return res;
    } catch (e) {
      entry.error = env.redact(String(e && e.message ? e.message : e)).split('\n')[0];
      throw e;
    } finally {
      entry.durationMs = Date.now() - started;
    }
  }

  /**
   * Throws the Apps Script-style exception for HTTP errors.
   * @param {!Object} req
   * @param {{code: number, bytes: !Buffer}} res
   */
  function throwIfHttpError(req, res) {
    if (req.muteHttpExceptions || res.code < 400) return;
    const origin = new URL(req.url).origin;
    const limit = opts.fetchErrorTruncate || DEFAULT_ERROR_TRUNCATE;
    const text = res.bytes.toString('utf8').replace(/\s+/g, ' ').trim();
    const shown = text.length > limit ? `${text.slice(0, limit)}...` : text;
    throw realm.error(`Request failed for ${origin} returned code ${res.code}. Truncated server ` +
        `response: ${shown} (use muteHttpExceptions option to examine full response)`);
  }

  /**
   * Simple triggers (AuthMode.LIMITED) may not use UrlFetchApp.
   * @param {string} method
   */
  function checkAuthMode(method) {
    if (env.authMode === 'LIMITED') {
      throw realm.error(`You do not have permission to call UrlFetchApp.${method}. Required permissions: ` +
          'https://www.googleapis.com/auth/script.external_request');
    }
  }

  /**
   * @param {string} url
   * @param {!Object=} params
   * @return {!HTTPResponseMock}
   */
  function fetch(url, params) {
    checkAuthMode('fetch');
    const req = normalizeRequest(url, params);
    const res = execute(req);
    throwIfHttpError(req, res);
    return new HTTPResponseMock(realm, res);
  }

  const service = {
    fetch,

    /**
     * Requests are executed sequentially (Apps Script runs them in parallel);
     * results keep request order.
     * @param {!Array<(string|!Object)>} requests
     * @return {!Array<!HTTPResponseMock>}
     */
    fetchAll(requests) {
      checkAuthMode('fetchAll');
      if (!Array.isArray(requests)) {
        throw realm.error(`The parameters (${typeof requests}) don't match the method signature for ` +
            'UrlFetchApp.fetchAll.');
      }
      const normalized = requests.map((r) => (typeof r === 'string' ? normalizeRequest(r, {}) :
        normalizeRequest(r && r.url, r)));
      const results = normalized.map((req) => [req, execute(req)]);
      for (const [req, res] of results) throwIfHttpError(req, res);
      return realm.toContext(results.map(([, res]) => new HTTPResponseMock(realm, res)));
    },

    /**
     * @param {string} url
     * @param {!Object=} params
     * @return {!Object} The request Apps Script would send.
     */
    getRequest(url, params) {
      const req = normalizeRequest(url, params);
      return realm.toContext({
        url: req.url,
        method: req.method.toLowerCase(),
        headers: {...req.headers},
        payload: req.payload === null ? '' : req.payload,
        contentType: req.contentType || '',
        followRedirects: req.followRedirects,
        validateHttpsCertificates: req.validateHttpsCertificates,
        muteHttpExceptions: req.muteHttpExceptions,
        useIntranet: false,
      });
    },

    toString() {
      return 'UrlFetchApp';
    },
  };

  return {
    service,
    fetchInternal: fetch,
    setHandler(h) {
      handler = h;
    },
    getHandler() {
      return handler;
    },
  };
}

module.exports = {createUrlFetchApp, HTTPResponseMock, redactHeaders, getHeader};
