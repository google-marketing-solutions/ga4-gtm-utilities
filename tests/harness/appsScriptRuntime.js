/**
 * @fileoverview Zero-dependency Apps Script runtime emulator for Node.js tests.
 *
 * Usage:
 *   const {createRuntime} = require('./harness/appsScriptRuntime');
 *   const {createFakeGtmApi} = require('./harness/fakeGtmApi');
 *   const fake = createFakeGtmApi();
 *   const rt = createRuntime({fetchHandler: fake});
 *   rt.spreadsheet.addSheet('GTM Workspace', [['Name', 'Path', 'Select']]);
 *   rt.call('onOpen');
 *
 * The Apps Script sources (features/, shared/, utils/ by default) are loaded
 * into one node:vm context, in path order, sharing a single global scope like
 * the V8 Apps Script runtime: function declarations become properties of the
 * global object and top-level const/let are visible to later files (but are not
 * properties of `runtime.context` - use runtime.get(name)).
 */
'use strict';

const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const {execSync} = require('node:child_process');

const {createRealm} = require('./realm');
const {UiController} = require('./uiMock');
const {SpreadsheetAppMock, SpreadsheetMock} = require('./spreadsheetMock');
const {SpreadsheetHarness} = require('./spreadsheetHarness');
const {createServices} = require('./servicesMock');
const {createUrlFetchApp} = require('./urlFetchMock');
const {createTagManagerService} = require('./tagManagerService');

/** Repository root (tests/harness/../..). */
const DEFAULT_ROOT = path.resolve(__dirname, '..', '..');
const DEFAULT_DIRS = ['features', 'shared', 'utils'];
const IDENTIFIER_PATH_RE = /^[A-Za-z_$][\w$]*(\.[A-Za-z_$][\w$]*)*$/;
const DEFAULT_FAKE_TOKEN = 'test-oauth-token';

/**
 * Replaces the context's Date with a wrapper whose "now" includes the virtual
 * clock offset. Instances are still real Dates of the context realm
 * (same prototype), so `instanceof Date` keeps working everywhere.
 */
const VIRTUAL_DATE_SHIM = `(function (offset, RealDate) {
  'use strict';
  function VirtualDate(...args) {
    if (!new.target) return new RealDate(RealDate.now() + offset()).toString();
    if (args.length === 0) return new RealDate(RealDate.now() + offset());
    return new RealDate(...args);
  }
  Object.defineProperty(VirtualDate, 'name', {value: 'Date'});
  Object.defineProperty(VirtualDate, 'length', {value: 7});
  VirtualDate.prototype = RealDate.prototype;
  VirtualDate.now = function now() { return RealDate.now() + offset(); };
  VirtualDate.parse = RealDate.parse;
  VirtualDate.UTC = RealDate.UTC;
  Object.defineProperty(RealDate.prototype, 'constructor',
      {value: VirtualDate, writable: true, configurable: true, enumerable: false});
  Object.defineProperty(globalThis, 'Date',
      {value: VirtualDate, writable: true, configurable: true, enumerable: false});
})(__harnessClockOffset, Date);`;

let cachedCommandToken = null;

/**
 * Resolves an OAuth token for live (realHttp) runs from the environment:
 * GTM_OAUTH_TOKEN, or the output of GTM_OAUTH_TOKEN_CMD (e.g.
 * `gcloud auth print-access-token`), cached for 45 minutes.
 * @return {?string} null when neither variable is set.
 */
function resolveLiveToken() {
  if (process.env.GTM_OAUTH_TOKEN && process.env.GTM_OAUTH_TOKEN.trim()) {
    return process.env.GTM_OAUTH_TOKEN.trim();
  }
  const cmd = process.env.GTM_OAUTH_TOKEN_CMD;
  if (!cmd) return null;
  if (cachedCommandToken && cachedCommandToken.cmd === cmd && cachedCommandToken.expires > Date.now()) {
    return cachedCommandToken.token;
  }
  let token;
  try {
    token = execSync(cmd, {encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'], timeout: 60000}).trim();
  } catch (e) {
    throw new Error(`GTM_OAUTH_TOKEN_CMD failed (exit ${e.status}); stderr: ` +
        `${String(e.stderr || '').trim().slice(0, 300)}`);
  }
  if (!token) throw new Error('GTM_OAUTH_TOKEN_CMD printed an empty token');
  cachedCommandToken = {cmd, token, expires: Date.now() + 45 * 60 * 1000};
  return token;
}

/**
 * Lists Apps Script source files (.js/.gs) under `dirs`, sorted by relative
 * path (code-point order) for a deterministic load order.
 * @param {string} rootDir
 * @param {!Array<string>} dirs
 * @return {!Array<string>} Absolute paths.
 */
function collectSourceFiles(rootDir, dirs) {
  const out = [];
  const walk = (dir) => {
    if (!fs.existsSync(dir)) return;
    for (const entry of fs.readdirSync(dir, {withFileTypes: true})) {
      const full = path.join(dir, entry.name);
      if (entry.isDirectory()) walk(full);
      else if (entry.isFile() && /\.(js|gs)$/.test(entry.name)) out.push(full);
    }
  };
  for (const d of dirs) walk(path.resolve(rootDir, d));
  const rel = (f) => path.relative(rootDir, f).split(path.sep).join('/');
  return out.sort((a, b) => (rel(a) < rel(b) ? -1 : rel(a) > rel(b) ? 1 : 0));
}

/**
 * Reads appsscript.json from rootDir (null when absent/invalid).
 * @param {string} rootDir
 * @return {?Object}
 */
function readManifest(rootDir) {
  const file = path.join(rootDir, 'appsscript.json');
  try {
    return JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch (e) {
    return null;
  }
}

/**
 * @typedef {{
 *   rootDir: (string|undefined),
 *   dirs: (!Array<string>|undefined),
 *   files: (!Array<string>|undefined),
 *   sources: (!Array<{filename: string, code: string}>|undefined),
 *   load: (boolean|undefined),
 *   loadErrors: (string|undefined),
 *   manifest: (?Object|undefined),
 *   advancedServices: (string|boolean|!Object|undefined),
 *   fetchHandler: (*|undefined),
 *   realHttp: (boolean|undefined),
 *   curlPath: (string|undefined),
 *   httpTimeoutSeconds: (number|undefined),
 *   oauthToken: (string|function(): string|undefined),
 *   userEmail: (string|undefined),
 *   effectiveUserEmail: (string|undefined),
 *   timeZone: (string|undefined),
 *   spreadsheet: (!Object|undefined),
 *   uiResponses: (!Array<*>|undefined),
 *   uiHandler: (function(!Object): *|undefined),
 *   properties: (!Object|undefined),
 *   realSleep: (boolean|undefined),
 *   echoLogs: (boolean|undefined),
 *   strictBounds: (boolean|undefined),
 *   enforceDataValidation: (boolean|undefined),
 *   enforceQuotas: (boolean|undefined),
 *   formulaEvaluator: (function(string, !Object): *|undefined),
 *   timeoutMs: (number|undefined),
 *   globals: (!Object|undefined),
 *   fetchErrorTruncate: (number|undefined),
 *   fetchLogPayloadLimit: (number|undefined),
 *   htmlFiles: (!Object<string, string>|undefined),
 *   scriptId: (string|undefined),
 * }}
 */
let RuntimeOptions;  // eslint-disable-line no-unused-vars

/**
 * Creates an emulated Apps Script runtime and loads the project sources.
 *
 * @param {!RuntimeOptions=} options See tests/README.md for every option.
 * @return {!Object} The runtime (see README "createRuntime API").
 */
function createRuntime(options = {}) {
  const opts = {...options};
  if (opts.echoLogs === undefined) opts.echoLogs = process.env.HARNESS_ECHO_LOGS === '1';
  const rootDir = path.resolve(opts.rootDir || DEFAULT_ROOT);
  const manifest = opts.manifest !== undefined ? opts.manifest : readManifest(rootDir);
  const scriptTimeZone = opts.timeZone || (manifest && manifest.timeZone) || 'America/New_York';

  const context = vm.createContext({}, {name: 'apps-script'});
  const realm = createRealm(context);

  // Virtual clock: real time plus everything "slept" via Utilities.sleep / lock waits.
  let virtualOffsetMs = 0;
  const now = () => Date.now() + virtualOffsetMs;
  const advanceClock = (ms) => {
    virtualOffsetMs += Math.max(0, Number(ms) || 0);
  };
  if (opts.virtualTime !== false) {
    // Make `Date.now()` / `new Date()` inside the script follow the virtual
    // clock, so elapsed-time logic (execution budgets, rate-limit windows)
    // sees time pass when Utilities.sleep() is virtual.
    Object.defineProperty(context, '__harnessClockOffset',
        {value: () => virtualOffsetMs, configurable: true, writable: true});
    vm.runInContext(VIRTUAL_DATE_SHIM, context, {filename: 'harness-virtual-date.js'});
    delete context.__harnessClockOffset;
  }

  const secrets = new Set();
  const addSecret = (s) => {
    if (typeof s === 'string' && s.length >= 8) secrets.add(s);
  };
  const redact = (text) => {
    let out = String(text);
    for (const s of secrets) out = out.split(s).join('<redacted>');
    return out;
  };

  const warnings = [];
  const ui = new UiController(realm, {responses: opts.uiResponses, handler: opts.uiHandler});

  const getOAuthToken = () => {
    let token;
    if (typeof opts.oauthToken === 'function') token = opts.oauthToken();
    else if (typeof opts.oauthToken === 'string') token = opts.oauthToken;
    else if (opts.realHttp) token = resolveLiveToken();
    else token = DEFAULT_FAKE_TOKEN;
    if (!token) {
      throw new Error('HarnessError: ScriptApp.getOAuthToken() needs a token in realHttp mode. Set ' +
          'options.oauthToken, GTM_OAUTH_TOKEN or GTM_OAUTH_TOKEN_CMD.');
    }
    addSecret(String(token));
    return String(token);
  };

  /** Shared environment handed to every mock. */
  const env = {
    realm,
    options: opts,
    rootDir,
    ui,
    logs: [],
    fetchLog: [],
    harnessErrors: [],
    sleeps: {totalMs: 0, calls: []},
    spreadsheets: new Map(),
    userEmail: opts.userEmail || 'test.user@example.com',
    scriptTimeZone,
    /** 'FULL' normally; 'LIMITED' while runtime.callTrigger() runs a simple trigger. */
    authMode: 'FULL',
    now,
    advanceClock,
    redact,
    addSecret,
    getOAuthToken,
    warn(code, message, details = {}) {
      warnings.push({code, message, ...details});
    },
  };

  // --- Spreadsheet ---------------------------------------------------------------
  const ssOpts = opts.spreadsheet || {};
  const spreadsheet = new SpreadsheetMock(env, {
    id: ssOpts.id || 'test-spreadsheet-id',
    name: ssOpts.name || 'Test Spreadsheet',
    timeZone: ssOpts.timeZone || scriptTimeZone,
    locale: ssOpts.locale,
  });
  env.spreadsheets.set(spreadsheet.getId(), spreadsheet);
  const SpreadsheetApp = new SpreadsheetAppMock(env, spreadsheet);
  const harness = new SpreadsheetHarness(env, spreadsheet, SpreadsheetApp);
  if (ssOpts.fixture) {
    harness.loadFixture(ssOpts.fixture);
  } else if (ssOpts.sheets) {
    const entries = Array.isArray(ssOpts.sheets) ? ssOpts.sheets.map((s) => [s.name, s.values || [], s]) :
      Object.entries(ssOpts.sheets).map(([name, values]) => [name, values, {}]);
    for (const [name, values, spec] of entries) harness.addSheet(name, values, spec);
  }
  if (spreadsheet._sheets.length === 0) spreadsheet._createSheet('Sheet1', 0, true);
  spreadsheet._activeSheet = spreadsheet._sheets[0];

  // --- Services ------------------------------------------------------------------
  const services = createServices(env);
  const urlFetch = createUrlFetchApp(env);
  const globals = {SpreadsheetApp, UrlFetchApp: urlFetch.service, ...services.globals};

  const advanced = opts.advancedServices === undefined ? 'auto' : opts.advancedServices;
  const enabledInManifest = ((manifest && manifest.dependencies &&
      manifest.dependencies.enabledAdvancedServices) || [])
      .filter((s) => String(s.serviceId).toLowerCase() === 'tagmanager');
  let tagManagerSymbols = [];
  if (advanced === 'auto') tagManagerSymbols = enabledInManifest.map((s) => s.userSymbol || 'TagManager');
  else if (advanced === true) tagManagerSymbols = ['TagManager'];
  else if (advanced && typeof advanced === 'object' && advanced.TagManager) {
    tagManagerSymbols = [typeof advanced.TagManager === 'string' ? advanced.TagManager : 'TagManager'];
  }
  for (const symbol of tagManagerSymbols) {
    globals[symbol] = createTagManagerService(env, urlFetch.fetchInternal);
  }
  Object.assign(globals, opts.globals || {});
  for (const [name, value] of Object.entries(globals)) {
    Object.defineProperty(context, name, {value, writable: true, configurable: true, enumerable: false});
  }
  /** Globals present before any project code ran (built-ins, mocks, shims). */
  const baselineGlobals = new Set(Object.getOwnPropertyNames(context));

  // --- Runtime object --------------------------------------------------------------
  const loadedFiles = [];
  const loadErrors = [];

  /**
   * Runs code in the Apps Script context.
   * @param {string} code
   * @param {string=} filename
   * @return {*} Raw (context-realm) completion value.
   */
  function run(code, filename = 'harness-eval.js') {
    const runOpts = {filename};
    if (opts.timeoutMs) runOpts.timeout = opts.timeoutMs;
    return vm.runInContext(code, context, runOpts);
  }

  /**
   * Loads one source file into the context.
   * @param {string} file Absolute or rootDir-relative path.
   */
  function loadFile(file) {
    const full = path.resolve(rootDir, file);
    const code = fs.readFileSync(full, 'utf8');
    const rel = path.relative(rootDir, full).split(path.sep).join('/');
    try {
      vm.runInContext(code, context, {filename: full});
      loadedFiles.push(rel);
    } catch (e) {
      const wrapped = new Error(`Failed to load ${rel}: ${e && e.message ? e.message : e}`, {cause: e});
      if (opts.loadErrors === 'collect') {
        loadErrors.push({file: rel, error: e});
      } else {
        throw wrapped;
      }
    }
  }

  /**
   * Reads a global (function declaration, var, or top-level const/let).
   * @param {string} name Identifier or dotted path ('sheetsMeta.eventTags').
   * @return {*} undefined when not defined.
   */
  function getRaw(name) {
    if (!IDENTIFIER_PATH_RE.test(name)) throw new Error(`Not an identifier path: ${name}`);
    const root = name.split('.')[0];
    if (vm.runInContext(`typeof ${root}`, context) === 'undefined') return undefined;
    return vm.runInContext(name, context);
  }

  /**
   * Invokes a global function with raw (unconverted) arguments.
   * @param {string} fnName
   * @param {...*} args
   * @return {*} Raw result.
   */
  function callRaw(fnName, ...args) {
    const fn = getRaw(fnName);
    if (typeof fn !== 'function') {
      throw new Error(`"${fnName}" is not a function in the Apps Script context (got ${typeof fn}).`);
    }
    if (!opts.timeoutMs) return fn.apply(undefined, args);
    Object.defineProperty(context, '__harnessArgs', {value: args, configurable: true, writable: true});
    try {
      return vm.runInContext(`${fnName}.apply(undefined, __harnessArgs)`, context,
          {timeout: opts.timeoutMs, filename: `call:${fnName}`});
    } finally {
      delete context.__harnessArgs;
    }
  }

  const runtime = {
    /** The vm context (the Apps Script global object). */
    context,
    /** Realm helpers: toContext/toHost/error. */
    realm,
    /** Options as passed (plus defaults). */
    options: opts,
    rootDir,
    manifest,
    run,
    loadFile,
    /**
     * Calls a global function; args are converted into the context realm and
     * the result is converted back to host plain data (arrays/objects/Dates).
     * @param {string} fnName
     * @param {...*} args
     * @return {*}
     */
    call(fnName, ...args) {
      return realm.toHost(callRaw(fnName, ...args.map((a) => realm.toContext(a))));
    },
    callRaw,
    /** @return {*} Host-converted value of a global (incl. top-level const). */
    get(name) {
      return realm.toHost(getRaw(name));
    },
    getRaw,
    /** @return {boolean} Whether a global with this name exists. */
    has(name) {
      return getRaw(name) !== undefined;
    },
    /** @return {boolean} Whether `name` is a function in the context. */
    isFunction(name) {
      return typeof getRaw(name) === 'function';
    },
    /** Deep-copies context-realm plain data (arrays/objects/Dates) into host objects. */
    toHost: (value) => realm.toHost(value),
    /** Deep-copies host plain data into the script's realm. */
    toContext: (value) => realm.toContext(value),
    /** Fixture helpers + the Spreadsheet mock (runtime.spreadsheet.active). */
    spreadsheet: harness,
    /** UI recorder/controller: alerts, prompts, toasts, menus, dialogs, queueResponse(). */
    ui,
    /** Every UrlFetchApp request (credentials redacted). */
    fetchLog: env.fetchLog,
    /** Logger/console output: [{source, level, message, time}]. */
    logs: env.logs,
    /** Suspicious-but-legal usages detected by the mocks. */
    warnings,
    /** Errors inside the harness itself (e.g. a throwing fetchHandler). */
    harnessErrors: env.harnessErrors,
    /** {totalMs, calls} for Utilities.sleep. */
    sleeps: env.sleeps,
    /** PropertiesStore per scope: {document, user, script}. */
    properties: services.properties,
    /** CacheStore per scope: {document, user, script}. */
    cache: services.cache,
    /** LockController (simulateContention(), events). */
    locks: services.locks,
    /** Triggers created through ScriptApp.newTrigger(). */
    triggers: services.triggers,
    /** The mock service objects installed as globals. */
    globals,
    loadedFiles,
    loadErrors,
    /** @return {number} Virtual "now" (real time + slept time). */
    now,
    advanceClock,
    /** Replaces the fetch handler (function or object with handle()). */
    setFetchHandler: urlFetch.setHandler,
    /** @return {string} All log messages joined by newlines. */
    logText() {
      return env.logs.map((l) => l.message).join('\n');
    },
    /** @return {!Array<string>} Function names referenced by menus added via addToUi(). */
    menuFunctionNames() {
      return ui.menuFunctionNames();
    },
    /** @return {!Array<string>} Menu function names that are not defined functions. */
    missingMenuFunctions() {
      return ui.menuFunctionNames().filter((n) => !IDENTIFIER_PATH_RE.test(n) || typeof getRaw(n) !== 'function');
    },
    /** @return {!Array<string>} Names of global functions declared by the loaded code. */
    globalFunctionNames() {
      return Object.getOwnPropertyNames(context)
          .filter((n) => !baselineGlobals.has(n) && typeof context[n] === 'function').sort();
    },
    /** Clears fetchLog, logs, warnings, UI recordings and sleep stats. */
    resetRecordings() {
      env.fetchLog.length = 0;
      env.logs.length = 0;
      warnings.length = 0;
      env.harnessErrors.length = 0;
      env.sleeps.totalMs = 0;
      env.sleeps.calls.length = 0;
      ui.reset();
    },
    /**
     * Calls a function the way Apps Script invokes a simple trigger: with an
     * event object and AuthMode.LIMITED, in which services that need
     * authorization (UrlFetchApp, and therefore the TagManager advanced
     * service) throw "You do not have permission to call ...".
     * @param {string} fnName e.g. 'onOpen'
     * @param {!Object=} event Merged into the default event object.
     * @param {{authMode: (string|undefined)}=} callOptions authMode 'LIMITED' (default) or 'FULL'.
     * @return {*} Host-converted result.
     */
    callTrigger(fnName, event = {}, callOptions = {}) {
      const authMode = callOptions.authMode || 'LIMITED';
      const evt = {authMode, source: spreadsheet, user: {getEmail: () => env.userEmail,
        toString: () => 'User'}, triggerUid: '0', ...event};
      const previous = env.authMode;
      env.authMode = authMode;
      try {
        return realm.toHost(callRaw(fnName, realm.toContext(evt)));
      } finally {
        env.authMode = previous;
      }
    },
    /**
     * Emulates the NEXT Apps Script execution (e.g. the user's next menu
     * click): a brand-new runtime with a fresh global scope (all top-level
     * code runs again), carrying over persistent state: the spreadsheet (via
     * toJSON()/loadFixture; filters, protections, conditional formats and
     * banding are not carried), properties, cache entries, the virtual clock,
     * pending UI responses and the same fetchHandler (e.g. FakeGtmApi state).
     * @param {!Object=} overrides createRuntime options to override.
     * @return {!Object} The new runtime.
     */
    nextExecution(overrides = {}) {
      const next = createRuntime({
        ...opts,
        uiResponses: [...ui.responses],
        fetchHandler: urlFetch.getHandler(),
        spreadsheet: {...ssOpts, sheets: undefined, fixture: harness.toJSON()},
        properties: {
          document: services.properties.document.all(),
          user: services.properties.user.all(),
          script: services.properties.script.all(),
        },
        ...overrides,
      });
      next.advanceClock(virtualOffsetMs);
      for (const scope of ['document', 'user', 'script']) {
        for (const [key, entry] of services.cache[scope].entries) next.cache[scope].entries.set(key, {...entry});
      }
      return next;
    },
  };

  // --- Load sources ------------------------------------------------------------------
  // `load: false` skips the project files on disk; inline `sources` always load.
  if (opts.load !== false) {
    const files = opts.files ? opts.files.map((f) => path.resolve(rootDir, f)) :
      collectSourceFiles(rootDir, opts.dirs || DEFAULT_DIRS);
    for (const file of files) loadFile(file);
  }
  for (const src of opts.sources || []) {
    const filename = src.filename || 'inline-source.js';
    try {
      run(src.code, filename);
      loadedFiles.push(filename);
    } catch (e) {
      if (opts.loadErrors === 'collect') loadErrors.push({file: filename, error: e});
      else throw new Error(`Failed to load ${filename}: ${e && e.message ? e.message : e}`, {cause: e});
    }
  }

  return runtime;
}

module.exports = {
  createRuntime,
  collectSourceFiles,
  resolveLiveToken,
  readManifest,
  DEFAULT_ROOT,
  DEFAULT_DIRS,
};
