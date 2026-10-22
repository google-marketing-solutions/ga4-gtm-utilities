/**
 * @fileoverview Mocks of the non-spreadsheet Apps Script services:
 * Utilities, Session, PropertiesService, CacheService, LockService, Logger,
 * console, ScriptApp, HtmlService, MimeType and Browser, plus the Blob class
 * shared with UrlFetchApp.
 */
'use strict';

const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const {defineHidden} = require('./uiMock');
const {isDate, isError} = require('./realm');
const {formatDateJava, parseDateJava} = require('./dateFormat');
const {enumOf} = require('./spreadsheetBuilders');

const MAX_SLEEP_MS = 300000;
const PROPERTY_VALUE_LIMIT = 9 * 1024;
const PROPERTY_STORE_LIMIT = 500 * 1024;
const CACHE_VALUE_LIMIT = 100 * 1024;
const CACHE_KEY_LIMIT = 250;
const CACHE_MAX_TTL_SECONDS = 21600;

// ---------------------------------------------------------------------------
// Bytes & blobs
// ---------------------------------------------------------------------------

/**
 * Converts Apps Script byte input (string | Byte[] | Blob) into a Buffer.
 * @param {*} data
 * @param {string=} charset 'UTF_8' (default) or 'US_ASCII'.
 * @return {!Buffer}
 */
function toBuffer(data, charset) {
  if (data === null || data === undefined) return Buffer.alloc(0);
  if (Buffer.isBuffer(data)) return data;
  if (data instanceof Uint8Array) return Buffer.from(data);
  if (Array.isArray(data)) return Buffer.from(data.map((b) => Number(b) & 0xff));
  if (data instanceof BlobMock) return Buffer.from(data._bytes);
  if (typeof data.getBytes === 'function') return toBuffer(data.getBytes());
  return Buffer.from(String(data), charset === 'US_ASCII' ? 'latin1' : 'utf8');
}

/**
 * @param {!Buffer} buf
 * @return {!Array<number>} Java-style signed bytes (-128..127).
 */
function toSignedBytes(buf) {
  return Array.from(buf, (b) => (b > 127 ? b - 256 : b));
}

/** Blob mock (Utilities.newBlob, HTTPResponse.getBlob). */
class BlobMock {
  /**
   * @param {!Object} realm
   * @param {!Buffer} bytes
   * @param {?string=} contentType
   * @param {?string=} name
   */
  constructor(realm, bytes, contentType = null, name = null) {
    defineHidden(this, {_realm: realm, _bytes: Buffer.from(bytes), _contentType: contentType, _name: name});
  }

  getBytes() {
    return this._realm.toContext(toSignedBytes(this._bytes));
  }

  setBytes(data) {
    this._bytes = toBuffer(data);
    return this;
  }

  getDataAsString(charset) {
    return this._bytes.toString(charset && /ascii|latin|8859/i.test(charset) ? 'latin1' : 'utf8');
  }

  setDataFromString(string, charset) {
    this._bytes = toBuffer(String(string), charset);
    return this;
  }

  getContentType() {
    return this._contentType;
  }

  setContentType(contentType) {
    this._contentType = contentType;
    return this;
  }

  getName() {
    return this._name;
  }

  setName(name) {
    this._name = name;
    return this;
  }

  copyBlob() {
    return new BlobMock(this._realm, this._bytes, this._contentType, this._name);
  }

  getAs(contentType) {
    return new BlobMock(this._realm, this._bytes, contentType, this._name);
  }

  getBlob() {
    return this;
  }

  isGoogleType() {
    return false;
  }

  toString() {
    return 'Blob';
  }
}

// ---------------------------------------------------------------------------
// Formatting helpers
// ---------------------------------------------------------------------------

/**
 * Formats one log argument roughly like the Apps Script execution log.
 * @param {*} value
 * @return {string}
 */
function formatLogArg(value) {
  if (typeof value === 'string') return value;
  if (value === undefined) return 'undefined';
  if (value === null) return 'null';
  if (typeof value === 'function') return `function ${value.name || ''}() {...}`;
  if (typeof value !== 'object') return String(value);
  if (isError(value)) return value.stack ? String(value.stack) : `${value.name}: ${value.message}`;
  if (isDate(value)) return Number.isNaN(value.getTime()) ? 'Invalid Date' : value.toISOString();
  if (!Array.isArray(value)) {
    const proto = Object.getPrototypeOf(value);
    // Object.prototype of *any* realm has a null prototype; class instances don't.
    const isPlain = proto === null || Object.getPrototypeOf(proto) === null;
    const ownToString = Object.prototype.hasOwnProperty.call(value, 'toString') &&
        typeof value.toString === 'function';
    if ((!isPlain || ownToString) && typeof value.toString === 'function') return String(value);
  }
  try {
    const seen = new WeakSet();
    return JSON.stringify(value, (k, v) => {
      if (v && typeof v === 'object') {
        if (seen.has(v)) return '[Circular]';
        seen.add(v);
      }
      return v;
    }) || String(value);
  } catch (e) {
    return String(value);
  }
}

/**
 * Minimal Java String.format / Utilities.formatString implementation:
 * %s %d %i %f %.Nf %x %o %e %b %% with optional width/zero-padding.
 * @param {string} template
 * @param {!Array<*>} args
 * @return {string}
 */
function formatString(template, args) {
  let i = 0;
  return String(template).replace(/%(-)?(0)?(\d+)?(?:\.(\d+))?([sdifxXoeEbB%])/g,
      (match, left, zero, width, precision, conv) => {
        if (conv === '%') return '%';
        const arg = args[i++];
        let out;
        switch (conv) {
          case 'd':
          case 'i': out = String(Math.trunc(Number(arg))); break;
          case 'f': out = Number(arg).toFixed(precision === undefined ? 6 : Number(precision)); break;
          case 'e':
          case 'E': out = Number(arg).toExponential(precision === undefined ? 6 : Number(precision)); break;
          case 'x': out = Math.trunc(Number(arg)).toString(16); break;
          case 'X': out = Math.trunc(Number(arg)).toString(16).toUpperCase(); break;
          case 'o': out = Math.trunc(Number(arg)).toString(8); break;
          case 'b':
          case 'B': out = String(Boolean(arg)); break;
          default:
            out = formatLogArg(arg);
            if (precision !== undefined) out = out.slice(0, Number(precision));
        }
        if (conv === 'E') out = out.toUpperCase();
        if (width && out.length < Number(width)) {
          out = left ? out.padEnd(Number(width)) :
            (zero && /[dfixXoeE]/.test(conv) ? out.padStart(Number(width), '0') : out.padStart(Number(width)));
        }
        return out;
      });
}

/**
 * Parses CSV text like Utilities.parseCsv.
 * @param {string} csv
 * @param {string=} delimiter
 * @return {!Array<!Array<string>>}
 */
function parseCsv(csv, delimiter = ',') {
  const rows = [];
  let row = [];
  let field = '';
  let quoted = false;
  const text = String(csv);
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"' && text[i + 1] === '"') {
        field += '"';
        i++;
      } else if (ch === '"') {
        quoted = false;
      } else {
        field += ch;
      }
    } else if (ch === '"') {
      quoted = true;
    } else if (ch === delimiter) {
      row.push(field);
      field = '';
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
    } else {
      field += ch;
    }
  }
  if (field !== '' || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

// ---------------------------------------------------------------------------
// Properties
// ---------------------------------------------------------------------------

/**
 * Host-side storage for one property store (exposed as runtime.properties.*).
 */
class PropertiesStore {
  /**
   * @param {string} scope 'document' | 'user' | 'script'
   * @param {!Object=} initial
   */
  constructor(scope, initial = {}) {
    this.scope = scope;
    /** @type {!Map<string, string>} */
    this.data = new Map(Object.entries(initial).map(([k, v]) => [k, String(v)]));
  }

  /** @return {?string} */
  get(key) {
    return this.data.has(key) ? this.data.get(key) : null;
  }

  /** Sets a property directly (no quota checks). */
  set(key, value) {
    this.data.set(String(key), String(value));
    return this;
  }

  /** @return {!Object<string, string>} Plain copy of all properties. */
  all() {
    return Object.fromEntries(this.data);
  }

  clear() {
    this.data.clear();
  }
}

/** The `Properties` object returned by PropertiesService.get*Properties(). */
class PropertiesMock {
  /**
   * @param {!Object} env
   * @param {!PropertiesStore} store
   */
  constructor(env, store) {
    defineHidden(this, {_env: env, _store: store});
  }

  /**
   * @param {string} key
   * @param {*} value
   */
  _put(key, value) {
    const k = String(key);
    if (value !== null && typeof value === 'object') {
      this._env.warn('non-string-property', `PropertiesService value for "${k}" is an object; ` +
          'Apps Script stores its string form. Use JSON.stringify.', {key: k});
    }
    const v = String(value);
    if (this._env.options.enforceQuotas !== false) {
      if (Buffer.byteLength(v, 'utf8') > PROPERTY_VALUE_LIMIT) {
        throw this._env.realm.error('Argument too large: value');
      }
      let total = Buffer.byteLength(k) + Buffer.byteLength(v);
      for (const [ek, ev] of this._store.data) if (ek !== k) total += Buffer.byteLength(ek) + Buffer.byteLength(ev);
      if (total > PROPERTY_STORE_LIMIT) {
        throw this._env.realm.error('You have exceeded the property storage quota. Please remove ' +
            'some properties and try again.');
      }
    }
    this._store.data.set(k, v);
  }

  getProperty(key) {
    return this._store.get(String(key));
  }

  getProperties() {
    return this._env.realm.toContext(this._store.all());
  }

  getKeys() {
    return this._env.realm.toContext([...this._store.data.keys()]);
  }

  setProperty(key, value) {
    this._put(key, value);
    return this;
  }

  setProperties(properties, deleteAllOthers = false) {
    if (deleteAllOthers) this._store.data.clear();
    for (const [k, v] of Object.entries(properties || {})) this._put(k, v);
    return this;
  }

  deleteProperty(key) {
    this._store.data.delete(String(key));
    return this;
  }

  deleteAllProperties() {
    this._store.data.clear();
    return this;
  }

  toString() {
    return 'Properties';
  }
}

// ---------------------------------------------------------------------------
// Cache
// ---------------------------------------------------------------------------

/** Host-side storage for one cache scope (runtime.cache.*). */
class CacheStore {
  /**
   * @param {string} scope
   * @param {function(): number} now Virtual clock (ms).
   */
  constructor(scope, now) {
    this.scope = scope;
    this.now = now;
    /** @type {!Map<string, {value: string, expiresAt: number}>} */
    this.entries = new Map();
  }

  /** @return {?string} The live value or null. */
  get(key) {
    const e = this.entries.get(key);
    if (!e) return null;
    if (e.expiresAt <= this.now()) {
      this.entries.delete(key);
      return null;
    }
    return e.value;
  }

  /** @return {!Object<string, string>} Live entries. */
  all() {
    const out = {};
    for (const key of [...this.entries.keys()]) {
      const v = this.get(key);
      if (v !== null) out[key] = v;
    }
    return out;
  }

  clear() {
    this.entries.clear();
  }
}

/** The `Cache` object returned by CacheService.get*Cache(). */
class CacheMock {
  /**
   * @param {!Object} env
   * @param {!CacheStore} store
   */
  constructor(env, store) {
    defineHidden(this, {_env: env, _store: store});
  }

  _check(key, value) {
    const realm = this._env.realm;
    if (String(key).length > CACHE_KEY_LIMIT) throw realm.error('Argument too large: key');
    if (value !== undefined && Buffer.byteLength(String(value), 'utf8') > CACHE_VALUE_LIMIT) {
      throw realm.error('Argument too large: value');
    }
  }

  get(key) {
    return this._store.get(String(key));
  }

  getAll(keys) {
    const out = {};
    for (const key of keys) {
      const v = this._store.get(String(key));
      if (v !== null) out[key] = v;
    }
    return this._env.realm.toContext(out);
  }

  put(key, value, expirationInSeconds = 600) {
    this._check(key, value);
    const ttl = Math.min(Math.max(1, Number(expirationInSeconds) || 600), CACHE_MAX_TTL_SECONDS);
    this._store.entries.set(String(key), {value: String(value), expiresAt: this._store.now() + ttl * 1000});
  }

  putAll(values, expirationInSeconds = 600) {
    for (const [k, v] of Object.entries(values || {})) this.put(k, v, expirationInSeconds);
  }

  remove(key) {
    this._store.entries.delete(String(key));
  }

  removeAll(keys) {
    for (const key of keys) this._store.entries.delete(String(key));
  }

  toString() {
    return 'Cache';
  }
}

// ---------------------------------------------------------------------------
// Locks
// ---------------------------------------------------------------------------

/**
 * Host-side lock controller (runtime.locks). Execution is single-threaded, so
 * locks always succeed unless contention is simulated.
 */
class LockController {
  /** @param {function(number): void} advanceClock */
  constructor(advanceClock) {
    this.advanceClock = advanceClock;
    /** @type {!Object<string, ?Object>} Current holder per scope. */
    this.holders = {document: null, user: null, script: null};
    /** @type {!Object<string, boolean>} Simulated foreign holders per scope. */
    this.contended = {document: false, user: false, script: false};
    /** @type {!Array<!Object>} Every lock operation, for assertions. */
    this.events = [];
  }

  /**
   * Makes every lock in `scope` (or all scopes) appear held by another execution.
   * @param {boolean} value
   * @param {string=} scope
   */
  simulateContention(value = true, scope) {
    for (const s of scope ? [scope] : Object.keys(this.contended)) this.contended[s] = Boolean(value);
  }
}

/** Lock mock. */
class LockMock {
  /**
   * @param {!Object} env
   * @param {!LockController} controller
   * @param {string} scope
   */
  constructor(env, controller, scope) {
    defineHidden(this, {_env: env, _c: controller, _scope: scope, _has: false});
  }

  _available() {
    const holder = this._c.holders[this._scope];
    return !this._c.contended[this._scope] && (!holder || holder === this);
  }

  tryLock(timeoutInMillis) {
    const ok = this._available();
    if (ok) {
      this._c.holders[this._scope] = this;
      this._has = true;
    } else {
      this._c.advanceClock(Number(timeoutInMillis) || 0);
    }
    this._c.events.push({op: 'tryLock', scope: this._scope, timeoutInMillis, acquired: ok});
    return ok;
  }

  waitLock(timeoutInMillis) {
    const ok = this.tryLock(timeoutInMillis);
    this._c.events[this._c.events.length - 1].op = 'waitLock';
    if (!ok) {
      throw this._env.realm.error('Lock timeout: another process was holding the lock for too long.');
    }
  }

  releaseLock() {
    if (this._c.holders[this._scope] === this) this._c.holders[this._scope] = null;
    this._has = false;
    this._c.events.push({op: 'releaseLock', scope: this._scope});
  }

  hasLock() {
    return this._has;
  }

  toString() {
    return 'Lock';
  }
}

// ---------------------------------------------------------------------------
// Logging
// ---------------------------------------------------------------------------

/**
 * Formats console.* arguments: Apps Script's console methods accept
 * `(formatOrObject, ...values)` with %s %d %i %f %o %O %j substitutions;
 * leftover values are appended separated by spaces.
 * @param {!Array<*>} args
 * @return {string}
 */
function formatConsoleArgs(args) {
  if (args.length > 1 && typeof args[0] === 'string' && /%[sdifoOj%]/.test(args[0])) {
    let i = 1;
    const head = args[0].replace(/%([sdifoOj%])/g, (match, conv) => {
      if (conv === '%') return '%';
      if (i >= args.length) return match;
      const arg = args[i++];
      if (conv === 'd' || conv === 'i') return String(parseInt(arg, 10));
      if (conv === 'f') return String(parseFloat(arg));
      return formatLogArg(arg);
    });
    return [head, ...args.slice(i).map(formatLogArg)].join(' ');
  }
  return args.map(formatLogArg).join(' ');
}

/**
 * Creates Logger and console mocks writing into env.logs.
 * @param {!Object} env
 * @return {{Logger: !Object, console: !Object}}
 */
function createLogging(env) {
  const record = (source, level, args) => {
    const message = source === 'console' ? formatConsoleArgs(args) : args.map(formatLogArg).join(' ');
    const entry = {source, level, message, time: env.now()};
    env.logs.push(entry);
    if (env.options.echoLogs) {
      const out = level === 'error' || level === 'warn' ? process.stderr : process.stdout;
      out.write(`[apps-script ${source}.${level}] ${env.redact(message)}\n`);
    }
    return entry;
  };
  const timers = new Map();
  const loggerLines = [];
  const Logger = {
    log(data, ...values) {
      const message = values.length && typeof data === 'string' ? formatString(data, values) : formatLogArg(data);
      record('Logger', 'log', [message]);
      loggerLines.push(message);
      return Logger;
    },
    getLog() {
      return loggerLines.join('\n');
    },
    clear() {
      loggerLines.length = 0;
    },
    toString() {
      return 'Logger';
    },
  };
  const console = {
    log: (...args) => void record('console', 'log', args),
    info: (...args) => void record('console', 'info', args),
    warn: (...args) => void record('console', 'warn', args),
    error: (...args) => void record('console', 'error', args),
    debug: (...args) => void record('console', 'debug', args),
    time: (label = 'default') => void timers.set(label, env.now()),
    timeEnd: (label = 'default') => {
      const start = timers.get(label);
      timers.delete(label);
      record('console', 'info', [`${label}: ${start === undefined ? 0 : env.now() - start}ms`]);
    },
  };
  return {Logger, console};
}

// ---------------------------------------------------------------------------
// Service factory
// ---------------------------------------------------------------------------

/**
 * Creates all non-spreadsheet service mocks.
 *
 * @param {!Object} env Runtime environment: {realm, options, logs, warn,
 *     redact, now, advanceClock, sleeps, userEmail, scriptTimeZone,
 *     getOAuthToken, rootDir, ui}.
 * @return {{globals: !Object<string, *>, properties: !Object, cache: !Object,
 *     locks: !LockController, triggers: !Array<!Object>}}
 */
function createServices(env) {
  const realm = env.realm;
  const opts = env.options;

  // --- Utilities -------------------------------------------------------------
  const DigestAlgorithm = enumOf('MD2', 'MD5', 'SHA_1', 'SHA_256', 'SHA_384', 'SHA_512');
  const MacAlgorithm = enumOf('HMAC_MD5', 'HMAC_SHA_1', 'HMAC_SHA_256', 'HMAC_SHA_384', 'HMAC_SHA_512');
  const Charset = enumOf('US_ASCII', 'UTF_8');
  const nodeHash = {
    MD2: 'md2', MD5: 'md5', SHA_1: 'sha1', SHA_256: 'sha256', SHA_384: 'sha384', SHA_512: 'sha512',
    HMAC_MD5: 'md5', HMAC_SHA_1: 'sha1', HMAC_SHA_256: 'sha256', HMAC_SHA_384: 'sha384', HMAC_SHA_512: 'sha512',
  };
  const hashName = (algorithm, method) => {
    const name = nodeHash[String(algorithm)];
    if (!name) {
      throw realm.error(`The parameters (${typeof algorithm}) don't match the method signature for ` +
          `Utilities.${method}.`);
    }
    return name;
  };

  const Utilities = {
    DigestAlgorithm,
    MacAlgorithm,
    Charset,
    RsaAlgorithm: enumOf('RSA_SHA_1', 'RSA_SHA_256'),

    /**
     * Virtual by default (advances the runtime clock); real when options.realSleep.
     * @param {number} milliseconds
     */
    sleep(milliseconds) {
      const ms = Math.max(0, Math.trunc(Number(milliseconds) || 0));
      if (ms > MAX_SLEEP_MS) throw realm.error(`Invalid argument: milliseconds (max ${MAX_SLEEP_MS})`);
      env.sleeps.calls.push(ms);
      env.sleeps.totalMs += ms;
      if (opts.realSleep) {
        Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);
      } else {
        env.advanceClock(ms);
      }
    },

    computeDigest(algorithm, value, charset) {
      const hash = crypto.createHash(hashName(algorithm, 'computeDigest'));
      hash.update(toBuffer(value, charset));
      return realm.toContext(toSignedBytes(hash.digest()));
    },

    computeHmacSignature(algorithm, value, key, charset) {
      const hmac = crypto.createHmac(hashName(algorithm, 'computeHmacSignature'), toBuffer(key, charset));
      hmac.update(toBuffer(value, charset));
      return realm.toContext(toSignedBytes(hmac.digest()));
    },

    computeHmacSha256Signature(value, key, charset) {
      return Utilities.computeHmacSignature('HMAC_SHA_256', value, key, charset);
    },

    computeRsaSha256Signature() {
      throw realm.error('Utilities.computeRsaSha256Signature is not implemented by the test harness.');
    },

    base64Encode(data, charset) {
      return toBuffer(data, charset).toString('base64');
    },

    base64EncodeWebSafe(data, charset) {
      return toBuffer(data, charset).toString('base64').replace(/\+/g, '-').replace(/\//g, '_');
    },

    base64Decode(encoded) {
      return realm.toContext(toSignedBytes(Buffer.from(String(encoded), 'base64')));
    },

    base64DecodeWebSafe(encoded) {
      const std = String(encoded).replace(/-/g, '+').replace(/_/g, '/');
      return realm.toContext(toSignedBytes(Buffer.from(std, 'base64')));
    },

    newBlob(data, contentType, name) {
      return new BlobMock(realm, toBuffer(data), contentType || null, name || null);
    },

    formatDate(date, timeZone, format) {
      if (!isDate(date)) {
        throw realm.error(`The parameters (${typeof date},${typeof timeZone},${typeof format}) don't ` +
            'match the method signature for Utilities.formatDate.');
      }
      try {
        return formatDateJava(date, String(timeZone), String(format));
      } catch (e) {
        throw realm.error(`Invalid argument: ${e.message}`);
      }
    },

    parseDate(dateString, timeZone, format) {
      try {
        return realm.date(parseDateJava(String(dateString), String(timeZone), String(format)));
      } catch (e) {
        throw realm.error(`Invalid argument: ${e.message}`);
      }
    },

    formatString(template, ...args) {
      return formatString(template, args);
    },

    getUuid() {
      return crypto.randomUUID();
    },

    jsonStringify(obj) {
      return JSON.stringify(obj);
    },

    jsonParse(jsonString) {
      return realm.toContext(JSON.parse(jsonString));
    },

    parseCsv(csv, delimiter) {
      return realm.toContext(parseCsv(csv, delimiter === undefined ? ',' : String(delimiter)));
    },

    zip() {
      throw realm.error('Utilities.zip is not implemented by the test harness.');
    },

    unzip() {
      throw realm.error('Utilities.unzip is not implemented by the test harness.');
    },

    toString() {
      return 'Utilities';
    },
  };

  // --- Session ---------------------------------------------------------------
  const makeUser = (email) => ({
    getEmail: () => email,
    getUsername: () => email.split('@')[0],
    toString: () => 'User',
  });
  const Session = {
    getActiveUser: () => makeUser(env.userEmail),
    getEffectiveUser: () => makeUser(opts.effectiveUserEmail || env.userEmail),
    getScriptTimeZone: () => env.scriptTimeZone,
    getTimeZone: () => env.scriptTimeZone,
    getActiveUserLocale: () => opts.userLocale || 'en',
    getTemporaryActiveUserKey: () => crypto.createHash('sha256').update(env.userEmail).digest('base64').slice(0, 24),
    toString: () => 'Session',
  };

  // --- PropertiesService -----------------------------------------------------
  const initialProps = opts.properties || {};
  const properties = {
    document: new PropertiesStore('document', initialProps.document),
    user: new PropertiesStore('user', initialProps.user),
    script: new PropertiesStore('script', initialProps.script),
  };
  const PropertiesService = {
    getDocumentProperties: () => new PropertiesMock(env, properties.document),
    getUserProperties: () => new PropertiesMock(env, properties.user),
    getScriptProperties: () => new PropertiesMock(env, properties.script),
    toString: () => 'PropertiesService',
  };

  // --- CacheService ----------------------------------------------------------
  const cache = {
    document: new CacheStore('document', env.now),
    user: new CacheStore('user', env.now),
    script: new CacheStore('script', env.now),
  };
  const CacheService = {
    getDocumentCache: () => new CacheMock(env, cache.document),
    getUserCache: () => new CacheMock(env, cache.user),
    getScriptCache: () => new CacheMock(env, cache.script),
    toString: () => 'CacheService',
  };

  // --- LockService -----------------------------------------------------------
  const locks = new LockController(env.advanceClock);
  const LockService = {
    getDocumentLock: () => new LockMock(env, locks, 'document'),
    getUserLock: () => new LockMock(env, locks, 'user'),
    getScriptLock: () => new LockMock(env, locks, 'script'),
    toString: () => 'LockService',
  };

  // --- Logger / console --------------------------------------------------------
  const {Logger, console} = createLogging(env);

  // --- ScriptApp -------------------------------------------------------------
  const triggers = [];
  const makeTrigger = (spec) => {
    const id = String(1000 + triggers.length);
    const trigger = {
      getHandlerFunction: () => spec.functionName,
      getUniqueId: () => id,
      getEventType: () => spec.eventType,
      getTriggerSource: () => spec.source,
      getTriggerSourceId: () => spec.sourceId || null,
      toString: () => 'Trigger',
    };
    defineHidden(trigger, {_spec: spec});
    return trigger;
  };
  const ScriptApp = {
    AuthMode: enumOf('NONE', 'CUSTOM_FUNCTION', 'LIMITED', 'FULL'),
    AuthorizationStatus: enumOf('REQUIRED', 'NOT_REQUIRED'),
    EventType: enumOf('CLOCK', 'ON_OPEN', 'ON_EDIT', 'ON_FORM_SUBMIT', 'ON_CHANGE', 'ON_EVENT_UPDATED'),
    TriggerSource: enumOf('SPREADSHEETS', 'CLOCK', 'FORMS', 'DOCUMENTS', 'CALENDAR'),
    WeekDay: enumOf('MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY', 'SUNDAY'),

    getOAuthToken: () => env.getOAuthToken(),
    getIdentityToken: () => 'header.payload.signature',
    getScriptId: () => opts.scriptId || 'test-script-id',
    getProjectTriggers: () => realm.toContext(triggers.map((t) => t.trigger)),
    getUserTriggers: () => realm.toContext(triggers.map((t) => t.trigger)),
    deleteTrigger(trigger) {
      const idx = triggers.findIndex((t) => t.trigger === trigger ||
          t.trigger.getUniqueId() === (trigger && trigger.getUniqueId && trigger.getUniqueId()));
      if (idx >= 0) triggers.splice(idx, 1);
    },
    newTrigger(functionName) {
      const spec = {functionName: String(functionName), options: {}};
      const finish = () => {
        const trigger = makeTrigger(spec);
        triggers.push({spec, trigger});
        return trigger;
      };
      const clock = {};
      for (const m of ['everyMinutes', 'everyHours', 'everyDays', 'everyWeeks', 'atHour', 'nearMinute',
        'onWeekDay', 'onMonthDay', 'at', 'atDate', 'after', 'inTimezone']) {
        clock[m] = (...a) => {
          spec.options[m] = a.length > 1 ? a : a[0];
          return clock;
        };
      }
      clock.create = () => finish();
      const source = {};
      for (const m of ['onOpen', 'onEdit', 'onChange', 'onFormSubmit']) {
        source[m] = () => {
          spec.eventType = {onOpen: 'ON_OPEN', onEdit: 'ON_EDIT', onChange: 'ON_CHANGE',
            onFormSubmit: 'ON_FORM_SUBMIT'}[m];
          return source;
        };
      }
      source.create = () => finish();
      return {
        timeBased: () => {
          spec.eventType = 'CLOCK';
          spec.source = 'CLOCK';
          return clock;
        },
        forSpreadsheet: (ss) => {
          spec.source = 'SPREADSHEETS';
          spec.sourceId = ss && ss.getId ? ss.getId() : String(ss);
          return source;
        },
        toString: () => 'TriggerBuilder',
      };
    },
    getService: () => ({getUrl: () => 'https://script.google.com/macros/s/test-deployment/exec',
      isEnabled: () => false, toString: () => 'Service'}),
    getAuthorizationInfo: () => ({getAuthorizationStatus: () => 'NOT_REQUIRED',
      getAuthorizationUrl: () => null, getAuthorizedScopes: () => realm.toContext([]),
      toString: () => 'AuthorizationInfo'}),
    invalidateAuth: () => {},
    requireScopes: () => {},
    requireAllScopes: () => {},
    toString: () => 'ScriptApp',
  };

  // --- HtmlService -------------------------------------------------------------
  const makeHtmlOutput = (content = '') => {
    const state = {content: String(content), title: '', width: null, height: null};
    const out = {
      getContent: () => state.content,
      setContent: (c) => ((state.content = String(c)), out),
      append: (c) => ((state.content += String(c)), out),
      appendUntrusted: (c) => ((state.content += String(c).replace(/[<>&]/g,
          (ch) => ({'<': '&lt;', '>': '&gt;', '&': '&amp;'}[ch]))), out),
      clear: () => ((state.content = ''), out),
      setTitle: (t) => ((state.title = String(t)), out),
      getTitle: () => state.title,
      setWidth: (w) => ((state.width = w), out),
      getWidth: () => state.width,
      setHeight: (h) => ((state.height = h), out),
      getHeight: () => state.height,
      setSandboxMode: () => out,
      setXFrameOptionsMode: () => out,
      addMetaTag: () => out,
      setFaviconUrl: () => out,
      getBlob: () => new BlobMock(realm, Buffer.from(state.content), 'text/html', null),
      getAs: (ct) => new BlobMock(realm, Buffer.from(state.content), ct, null),
      toString: () => 'HtmlOutput',
    };
    return out;
  };
  const readHtmlFile = (filename) => {
    const name = String(filename);
    if (opts.htmlFiles && Object.prototype.hasOwnProperty.call(opts.htmlFiles, name)) return opts.htmlFiles[name];
    for (const candidate of [name, `${name}.html`]) {
      const full = path.resolve(env.rootDir, candidate);
      if (fs.existsSync(full) && fs.statSync(full).isFile()) return fs.readFileSync(full, 'utf8');
    }
    throw realm.error(`No HTML file named ${name} was found.`);
  };
  const makeTemplate = (source) => {
    const tpl = {
      evaluate: () => makeHtmlOutput(source.replace(/<\?!?=?[\s\S]*?\?>/g, '')),
      getRawContent: () => source,
      getCode: () => source,
      toString: () => 'HtmlTemplate',
    };
    return tpl;
  };
  const HtmlService = {
    SandboxMode: enumOf('EMULATED', 'IFRAME', 'NATIVE'),
    XFrameOptionsMode: enumOf('ALLOWALL', 'DEFAULT'),
    createHtmlOutput: (html) => makeHtmlOutput(html === undefined ? '' :
      (html && typeof html.getDataAsString === 'function' ? html.getDataAsString() : html)),
    createHtmlOutputFromFile: (filename) => makeHtmlOutput(readHtmlFile(filename)),
    createTemplate: (html) => makeTemplate(String(html && html.getDataAsString ? html.getDataAsString() : html)),
    createTemplateFromFile: (filename) => makeTemplate(readHtmlFile(filename)),
    getUserAgent: () => 'Mozilla/5.0 (apps-script-test-harness)',
    toString: () => 'HtmlService',
  };

  // --- MimeType ----------------------------------------------------------------
  const MimeType = Object.freeze({
    JSON: 'application/json', PLAIN_TEXT: 'text/plain', CSV: 'text/csv', HTML: 'text/html',
    JAVASCRIPT: 'application/javascript', PDF: 'application/pdf', ZIP: 'application/zip',
    PNG: 'image/png', JPEG: 'image/jpeg', GIF: 'image/gif', SVG: 'image/svg+xml',
    MICROSOFT_EXCEL: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    GOOGLE_SHEETS: 'application/vnd.google-apps.spreadsheet',
    GOOGLE_DOCS: 'application/vnd.google-apps.document',
    GOOGLE_APPS_SCRIPT: 'application/vnd.google-apps.script',
    FOLDER: 'application/vnd.google-apps.folder',
  });

  // --- Browser (legacy dialogs, routed through the Ui recorder) -----------------
  const Browser = {
    Buttons: enumOf('OK', 'OK_CANCEL', 'YES_NO', 'YES_NO_CANCEL'),
    msgBox(...args) {
      const ui = env.ui.instance;
      const button = ui.alert(...args);
      return String(button).toLowerCase();
    },
    inputBox(...args) {
      const response = env.ui.instance.prompt(...args);
      const button = response.getSelectedButton();
      return button === 'CANCEL' || button === 'CLOSE' ? 'cancel' : response.getResponseText();
    },
    toString: () => 'Browser',
  };

  return {
    globals: {
      Utilities, Session, PropertiesService, CacheService, LockService, Logger, console,
      ScriptApp, HtmlService, MimeType, Browser,
    },
    properties,
    cache,
    locks,
    triggers,
  };
}

module.exports = {
  createServices,
  BlobMock,
  PropertiesStore,
  CacheStore,
  LockController,
  toBuffer,
  toSignedBytes,
  formatLogArg,
  formatString,
  parseCsv,
};
