/**
 * @fileoverview In-memory fake of the Google Tag Manager REST API v2, usable
 * as a `fetchHandler` for the Apps Script runtime (or directly via
 * `fake.handle(req)`).
 *
 *   const fake = createFakeGtmApi({pageSize: 200});
 *   const {workspacePath} = fake.seedWorkspace({tags: [{name: 'T', type: 'html'}]});
 *   const rt = createRuntime({fetchHandler: fake});
 *
 * Behaviour modelled on the real API:
 *  - List responses use the API's keys ({tag: [...], nextPageToken}) and omit
 *    the key entirely when empty (the real API returns `{}`), paginating
 *    after `pageSize` items (the API ignores client page sizes).
 *  - Creates assign server fields (ids from a per-container counter shared by
 *    all entity kinds, path, fingerprint, tagManagerUrl, account/container/
 *    workspace ids) and strip top-level false/''/[] fields like the API does.
 *  - PUT replaces the entity; `?fingerprint=` must match (409 ABORTED).
 *  - DELETE returns 204 with an empty body.
 *  - name (and type, where applicable) required; names unique per kind and
 *    workspace; unknown ids -> 404 "Not found or permission denied.";
 *    trigger/folder references are validated.
 *  - Errors use the Google error envelope:
 *    {"error": {"code", "message", "errors": [{"message", "domain", "reason"}], "status"}}
 *  - bulk_update is atomic, supports "new_N" ids and cross references.
 *  - Optional auth / content-type checks, rate limiting (429) and error
 *    injection for resilience tests.
 */
'use strict';

const {GTM_SCHEMAS} = require('./gtmSchema');

const GTM_API_ORIGIN = 'https://tagmanager.googleapis.com';
const GTM_API_PREFIX = '/tagmanager/v2/';
const GTM_API_BASE = `${GTM_API_ORIGIN}${GTM_API_PREFIX}`;
const UI_BASE = 'https://tagmanager.google.com/#';

/** Entity kinds stored inside workspaces. */
const ENTITY_KINDS = Object.freeze({
  tag: {collection: 'tags', idField: 'tagId', listKey: 'tag', entityKey: 'tag', requiresType: true},
  trigger: {collection: 'triggers', idField: 'triggerId', listKey: 'trigger', entityKey: 'trigger',
    requiresType: true},
  variable: {collection: 'variables', idField: 'variableId', listKey: 'variable', entityKey: 'variable',
    requiresType: true},
  folder: {collection: 'folders', idField: 'folderId', listKey: 'folder', entityKey: 'folder'},
  client: {collection: 'clients', idField: 'clientId', listKey: 'client', entityKey: 'client',
    requiresType: true},
  transformation: {collection: 'transformations', idField: 'transformationId', listKey: 'transformation',
    entityKey: 'transformation', requiresType: true},
  template: {collection: 'templates', idField: 'templateId', listKey: 'template', entityKey: 'customTemplate'},
  zone: {collection: 'zones', idField: 'zoneId', listKey: 'zone', entityKey: 'zone'},
  gtagConfig: {collection: 'gtag_config', idField: 'gtagConfigId', listKey: 'gtagConfig',
    entityKey: 'gtagConfig', requiresType: true, nameless: true},
});

const COLLECTION_TO_KIND = Object.freeze(Object.fromEntries(
    Object.entries(ENTITY_KINDS).map(([kind, info]) => [info.collection, kind])));

/** Keys accepted in bulk_update / getStatus Entity objects. */
const ENTITY_KEY_TO_KIND = Object.freeze({
  tag: 'tag', trigger: 'trigger', variable: 'variable', folder: 'folder', client: 'client',
  transformation: 'transformation', customTemplate: 'template', template: 'template', zone: 'zone',
  gtagConfig: 'gtagConfig', builtInVariable: 'builtInVariable',
});

/** Built-in trigger ids that tags may reference without them existing. */
const BUILT_IN_TRIGGER_IDS = Object.freeze(['2147479553', '2147479572', '2147479573',
  '2147479574', '2147479576']);

/** Fields that reference other entities by id: field -> referenced kind. */
const REFERENCE_FIELDS = Object.freeze({
  parentFolderId: 'folder',
  firingTriggerId: 'trigger',
  blockingTriggerId: 'trigger',
  enablingTriggerId: 'trigger',
  disablingTriggerId: 'trigger',
});

const BUILT_IN_VARIABLE_TYPES = Object.freeze(('pageUrl pageHostname pagePath referrer event clickElement ' +
  'clickClasses clickId clickTarget clickUrl clickText firstPartyServingUrl formElement formClasses formId ' +
  'formTarget formUrl formText errorMessage errorUrl errorLine newHistoryUrl oldHistoryUrl ' +
  'newHistoryFragment oldHistoryFragment newHistoryState oldHistoryState historySource containerVersion ' +
  'debugMode randomNumber containerId appId appName appVersionCode appVersionName language osVersion ' +
  'platform sdkVersion deviceName resolution advertiserId advertisingTrackingEnabled htmlId ' +
  'environmentName ampBrowserLanguage ampCanonicalPath ampCanonicalUrl ampCanonicalHost ampReferrer ' +
  'ampTitle ampClientId ampClientTimezone ampClientTimestamp ampClientScreenWidth ampClientScreenHeight ' +
  'ampClientScrollX ampClientScrollY ampClientMaxScrollX ampClientMaxScrollY ampTotalEngagedTime ' +
  'ampPageViewId ampPageLoadTime ampPageDownloadTime ampGtmEvent eventName ' +
  'firebaseEventParameterCampaign firebaseEventParameterCampaignAclid firebaseEventParameterCampaignAnid ' +
  'firebaseEventParameterCampaignClickTimestamp firebaseEventParameterCampaignContent ' +
  'firebaseEventParameterCampaignCp1 firebaseEventParameterCampaignGclid ' +
  'firebaseEventParameterCampaignSource firebaseEventParameterCampaignTerm firebaseEventParameterCurrency ' +
  'firebaseEventParameterDynamicLinkAcceptTime firebaseEventParameterDynamicLinkLinkid ' +
  'firebaseEventParameterNotificationMessageDeviceTime firebaseEventParameterNotificationMessageId ' +
  'firebaseEventParameterNotificationMessageName firebaseEventParameterNotificationMessageTime ' +
  'firebaseEventParameterNotificationTopic firebaseEventParameterPreviousAppVersion ' +
  'firebaseEventParameterPreviousOsVersion firebaseEventParameterPrice firebaseEventParameterProductId ' +
  'firebaseEventParameterQuantity firebaseEventParameterValue videoProvider videoUrl videoTitle ' +
  'videoDuration videoPercent videoVisible videoStatus videoCurrentTime scrollDepthThreshold ' +
  'scrollDepthUnits scrollDepthDirection elementVisibilityRatio elementVisibilityTime ' +
  'elementVisibilityFirstTime elementVisibilityRecentTime requestPath requestMethod clientName ' +
  'queryString serverPageLocationUrl serverPageLocationPath serverPageLocationHostname visitorRegion ' +
  'analyticsClientId analyticsSessionId analyticsSessionNumber').split(' '));

const BUILT_IN_VARIABLE_NAMES = Object.freeze({
  pageUrl: 'Page URL', pageHostname: 'Page Hostname', pagePath: 'Page Path', referrer: 'Referrer',
  event: 'Event', clickElement: 'Click Element', clickClasses: 'Click Classes', clickId: 'Click ID',
  clickTarget: 'Click Target', clickUrl: 'Click URL', clickText: 'Click Text', formElement: 'Form Element',
  formClasses: 'Form Classes', formId: 'Form ID', formTarget: 'Form Target', formUrl: 'Form URL',
  formText: 'Form Text', containerId: 'Container ID', containerVersion: 'Container Version',
  debugMode: 'Debug Mode', randomNumber: 'Random Number', htmlId: 'HTML ID',
  environmentName: 'Environment Name', eventName: 'Event Name', clientName: 'Client Name',
  queryString: 'Query String', requestPath: 'Request Path', requestMethod: 'Request Method',
});

/** Canonical status/reason per HTTP code for error envelopes. */
const ERROR_DEFAULTS = Object.freeze({
  400: ['INVALID_ARGUMENT', 'badRequest'],
  401: ['UNAUTHENTICATED', 'required'],
  403: ['PERMISSION_DENIED', 'forbidden'],
  404: ['NOT_FOUND', 'notFound'],
  409: ['ABORTED', 'conflict'],
  412: ['FAILED_PRECONDITION', 'conditionNotMet'],
  429: ['RESOURCE_EXHAUSTED', 'rateLimitExceeded'],
  500: ['INTERNAL', 'backendError'],
  501: ['NOT_IMPLEMENTED', 'notImplemented'],
  503: ['UNAVAILABLE', 'backendError'],
});

const NOT_FOUND_MESSAGE = 'Not found or permission denied.';

/** Error thrown inside the fake; converted to a Google error response. */
class ApiError extends Error {
  /**
   * @param {number} code
   * @param {string} message
   * @param {{status: (string|undefined), reason: (string|undefined), details: (!Array|undefined),
   *     extraErrorFields: (!Object|undefined)}=} opts
   */
  constructor(code, message, opts = {}) {
    super(message);
    this.code = code;
    const [status, reason] = ERROR_DEFAULTS[code] || ['UNKNOWN', 'unknown'];
    this.status = opts.status || status;
    this.reason = opts.reason || reason;
    this.details = opts.details;
    this.extraErrorFields = opts.extraErrorFields;
  }
}

/** @return {*} Deep copy of JSON data. */
function clone(v) {
  return v === undefined ? undefined : JSON.parse(JSON.stringify(v));
}

/** Discovery schema per entity kind, and the request field name used in error paths. */
const KIND_SCHEMA = Object.freeze({
  tag: ['Tag', 'tag'], trigger: ['Trigger', 'trigger'], variable: ['Variable', 'variable'],
  folder: ['Folder', 'folder'], client: ['Client', 'client'], transformation: ['Transformation', 'transformation'],
  template: ['CustomTemplate', 'template'], zone: ['Zone', 'zone'], gtagConfig: ['GtagConfig', 'gtag_config'],
});

/** @return {string} camelCase -> snake_case (proto field names used in ESF error paths). */
function snake(name) {
  return String(name).replace(/[A-Z]/g, (c) => `_${c.toLowerCase()}`);
}

/**
 * Validates a request object against a discovery schema the way the API's
 * JSON parser does: unknown field names, invalid enum values and structural
 * mismatches (object vs. list vs. scalar) are rejected with 400. Scalar
 * coercions the parser tolerates (numbers for strings, "true" for booleans)
 * are accepted.
 * @param {*} value
 * @param {string} schemaName
 * @param {string} where Error path, e.g. 'tag' or 'changes[0].tag'.
 */
function checkSchema(value, schemaName, where) {
  const schema = GTM_SCHEMAS[schemaName];
  if (!schema) return;
  if (value === null || value === undefined) return;
  if (typeof value !== 'object' || Array.isArray(value)) {
    throw new ApiError(400, `Invalid JSON payload received. Expected an object at '${where}', got ` +
        `${JSON.stringify(value).slice(0, 60)}.`, {reason: 'invalid'});
  }
  for (const [key, v] of Object.entries(value)) {
    const spec = schema[key];
    if (!spec) {
      throw new ApiError(400, `Invalid JSON payload received. Unknown name "${key}" at '${where}': ` +
          'Cannot find field.', {reason: 'invalid'});
    }
    if (v === null || v === undefined) continue;
    const at = `${where}.${snake(key)}`;
    const bad = (kind) => new ApiError(400, `Invalid value at '${at}' (${kind}), ` +
        `${JSON.stringify(v).slice(0, 80)}`, {reason: 'invalid'});
    if (spec.ref) {
      checkSchema(v, spec.ref, at);
    } else if (spec.arrayOf || spec.array) {
      if (!Array.isArray(v)) throw bad('repeated field: expected a JSON array');
      v.forEach((item, i) => {
        if (spec.arrayOf) {
          checkSchema(item, spec.arrayOf, `${at}[${i}]`);
        } else if (item !== null && typeof item === 'object') {
          throw new ApiError(400, `Invalid value at '${at}[${i}]' (TYPE_STRING), ` +
              `${JSON.stringify(item).slice(0, 80)}`, {reason: 'invalid'});
        } else if (spec.enum && !spec.enum.includes(item)) {
          throw new ApiError(400, `Invalid value at '${at}[${i}]' (TYPE_ENUM), ${JSON.stringify(item)}`,
              {reason: 'invalid'});
        }
      });
    } else if (typeof v === 'object') {
      throw bad(spec.enum ? 'TYPE_ENUM' : `TYPE_${String(spec.type).toUpperCase()}`);
    } else if (spec.enum && !spec.enum.includes(v)) {
      throw bad('TYPE_ENUM');
    } else if (spec.type === 'integer' && !/^-?\d+$/.test(String(v))) {
      throw bad('TYPE_INT32');
    } else if (spec.type === 'boolean' && !/^(true|false)$/i.test(String(v))) {
      throw bad('TYPE_BOOL');
    }
  }
}

/** @return {string} Kind name for user-facing messages ('gtagConfig' -> 'Google tag config'). */
function kindLabel(kind) {
  return {gtagConfig: 'Google tag config', builtInVariable: 'built-in variable'}[kind] || kind;
}

/**
 * Normalizes user-provided kind names: 'tags' | 'Tag' | 'gtag_config' | 'customTemplate' ...
 * @param {string} kind
 * @return {string}
 */
function normalizeKind(kind) {
  const k = String(kind);
  if (ENTITY_KINDS[k] || k === 'builtInVariable') return k;
  if (COLLECTION_TO_KIND[k]) return COLLECTION_TO_KIND[k];
  if (ENTITY_KEY_TO_KIND[k]) return ENTITY_KEY_TO_KIND[k];
  const lower = k.toLowerCase().replace(/[_\s-]/g, '');
  const aliases = {
    tag: 'tag', tags: 'tag', trigger: 'trigger', triggers: 'trigger', variable: 'variable',
    variables: 'variable', folder: 'folder', folders: 'folder', client: 'client', clients: 'client',
    transformation: 'transformation', transformations: 'transformation', template: 'template',
    templates: 'template', customtemplate: 'template', customtemplates: 'template', zone: 'zone',
    zones: 'zone', gtagconfig: 'gtagConfig', gtagconfigs: 'gtagConfig', builtinvariable: 'builtInVariable',
    builtinvariables: 'builtInVariable',
  };
  if (aliases[lower]) return aliases[lower];
  throw new Error(`FakeGtmApi: unknown entity kind "${kind}"`);
}

/**
 * @param {string} type Built-in variable type.
 * @return {string} Display name.
 */
function builtInName(type) {
  if (BUILT_IN_VARIABLE_NAMES[type]) return BUILT_IN_VARIABLE_NAMES[type];
  const words = type.replace(/([a-z0-9])([A-Z])/g, '$1 $2');
  return words.charAt(0).toUpperCase() + words.slice(1);
}

/**
 * The fake API. Create with createFakeGtmApi(options).
 */
class FakeGtmApi {
  /**
   * @param {{
   *   pageSize: (number|undefined),
   *   rateLimit: ({max: number, windowMs: number, code: (number|undefined)}|undefined),
   *   requireAuth: (boolean|undefined),
   *   acceptedTokens: (!Array<string>|undefined),
   *   strictContentType: (boolean|undefined),
   *   validateReferences: (boolean|undefined),
   *   omitEmptyFields: (boolean|undefined),
   *   maxWorkspacesPerContainer: (number|undefined),
   *   fingerprintMismatchCode: (number|undefined),
   *   builtInTriggerIds: (!Array<string>|undefined),
   *   now: (function(): number|undefined),
   * }=} options
   */
  constructor(options = {}) {
    this.options = {
      pageSize: 200,
      requireAuth: true,
      strictContentType: true,
      validateReferences: true,
      validateSchema: true,
      omitEmptyFields: true,
      maxWorkspacesPerContainer: 3,
      fingerprintMismatchCode: 409,
      builtInTriggerIds: BUILT_IN_TRIGGER_IDS,
      ...options,
    };
    /** @type {!Array<!Object>} Every request handled (see README). */
    this.requests = [];
    /** Bound handler, usable directly as a fetchHandler. */
    this.handler = (req) => this.handle(req);
    this.reset();
  }

  /** Clears all state, request logs, injected errors and rate-limit history. */
  reset() {
    /** @type {!Map<string, !Object>} accountId -> account */
    this._accounts = new Map();
    /** @type {!Map<string, !Object>} 'a/c' -> {data, counters, baseEntities} */
    this._containers = new Map();
    /** @type {!Map<string, !Object>} 'a/c/w' -> workspace record */
    this._workspaces = new Map();
    /** @type {!Map<string, !Map<string, !Object>>} 'a/c' -> linkId -> destination */
    this._destinations = new Map();
    this._accountCounter = 1000;
    this._lastFingerprint = 0;
    this._errorInjections = [];
    this._rateTimes = [];
    this.requests.length = 0;
  }

  // -------------------------------------------------------------------------
  // Paths, ids, fingerprints
  // -------------------------------------------------------------------------

  /** @return {string} A new, strictly increasing Date.now()-based fingerprint. */
  _fingerprint() {
    const now = this.options.now ? this.options.now() : Date.now();
    this._lastFingerprint = Math.max(now, this._lastFingerprint + 1);
    return String(this._lastFingerprint);
  }

  /**
   * Builds a workspace path.
   * @param {string|number} accountId
   * @param {string|number} containerId
   * @param {string|number} workspaceId
   * @return {string}
   */
  workspacePath(accountId, containerId, workspaceId) {
    return `accounts/${accountId}/containers/${containerId}/workspaces/${workspaceId}`;
  }

  /**
   * Parses 'accounts/1/containers/2/workspaces/3[/...]' into ids.
   * @param {string} p
   * @return {{accountId: string, containerId: (string|undefined), workspaceId: (string|undefined)}}
   */
  static parsePath(p) {
    const m = /^accounts\/([^/]+)(?:\/containers\/([^/]+)(?:\/workspaces\/([^/]+))?)?/.exec(String(p));
    if (!m) throw new Error(`FakeGtmApi: not a GTM path: ${p}`);
    return {accountId: m[1], containerId: m[2], workspaceId: m[3]};
  }

  /** @return {!Object} Container record (creates counters lazily). */
  _containerRec(accountId, containerId) {
    const rec = this._containers.get(`${accountId}/${containerId}`);
    if (!rec) throw new ApiError(404, NOT_FOUND_MESSAGE);
    return rec;
  }

  /** @return {!Object} Workspace record or throws 404. */
  _workspaceRec(accountId, containerId, workspaceId) {
    this._containerRec(accountId, containerId);
    const rec = this._workspaces.get(`${accountId}/${containerId}/${workspaceId}`);
    if (!rec) throw new ApiError(404, NOT_FOUND_MESSAGE);
    return rec;
  }

  /** @return {!Object} Empty entity maps for a workspace state. */
  static _emptyEntities() {
    const out = {};
    for (const kind of Object.keys(ENTITY_KINDS)) out[kind] = new Map();
    return out;
  }

  /**
   * Deep-copies workspace entity state (for atomic operations / snapshots).
   * @param {!Object} state {entities, builtIns}
   * @return {!Object}
   */
  static _cloneState(state) {
    const entities = {};
    for (const [kind, map] of Object.entries(state.entities)) {
      entities[kind] = new Map([...map].map(([id, e]) => [id, clone(e)]));
    }
    return {entities, builtIns: new Map([...state.builtIns].map(([t, b]) => [t, clone(b)]))};
  }

  // -------------------------------------------------------------------------
  // Seeding
  // -------------------------------------------------------------------------

  /**
   * Seeds accounts, containers, workspaces, entities and destinations.
   * Missing parents are created automatically. Seeded entities form the
   * workspace's base version (changeStatus "none" in getStatus).
   *
   * @param {{
   *   accounts: (!Array<!Object>|undefined),
   *   containers: (!Array<!Object>|undefined),
   *   workspaces: (!Array<!Object>|undefined),
   *   entities: (!Object<string, !Object<string, !Array>>|undefined),
   *   destinations: (!Array<!Object>|undefined),
   * }} spec entities: {workspacePath: {tag: [...], triggers: [...],
   *     builtInVariable: ['pageUrl' | {type}], ...}} (kind names are flexible).
   * @return {!FakeGtmApi}
   */
  seed(spec = {}) {
    for (const a of spec.accounts || []) this._seedAccount(a);
    for (const c of spec.containers || []) this._seedContainer(c);
    for (const w of spec.workspaces || []) this._seedWorkspace(w);
    for (const [wsPath, kinds] of Object.entries(spec.entities || {})) {
      const {accountId, containerId, workspaceId} = FakeGtmApi.parsePath(wsPath);
      if (!workspaceId) throw new Error(`FakeGtmApi.seed: entities key must be a workspace path: ${wsPath}`);
      const rec = this._seedWorkspace({accountId, containerId, workspaceId});
      for (const [kindName, items] of Object.entries(kinds || {})) {
        const kind = normalizeKind(kindName);
        for (const item of items || []) {
          if (kind === 'builtInVariable') this._seedBuiltIn(rec, typeof item === 'string' ? item : item.type);
          else this._seedEntity(rec, kind, item);
        }
      }
      // Seeded references must resolve too (catches test-data mistakes).
      for (const kind of Object.keys(ENTITY_KINDS)) {
        for (const entity of rec.entities[kind].values()) {
          try {
            this._validateRefs(rec, kind, entity);
          } catch (e) {
            throw new Error(`FakeGtmApi.seed: ${kind} ${JSON.stringify(entity.name || entity[ENTITY_KINDS[kind].idField])} ` +
                `in ${wsPath}: ${e.message} (seed the referenced entity, or pass validateReferences: false)`);
          }
        }
      }
      rec.base = FakeGtmApi._cloneState(rec);
      const cRec = this._containers.get(`${accountId}/${containerId}`);
      if (!cRec.baseFrom || cRec.baseFrom === workspaceId) {
        cRec.baseFrom = workspaceId;
        cRec.baseEntities = FakeGtmApi._cloneState(rec);
      }
    }
    for (const d of spec.destinations || []) this._seedDestination(d);
    return this;
  }

  /**
   * Convenience: ensures account/container/workspace exist and seeds entities.
   * @param {{accountId: (string|undefined), containerId: (string|undefined),
   *     workspaceId: (string|undefined), accountName: (string|undefined),
   *     containerName: (string|undefined), workspaceName: (string|undefined),
   *     publicId: (string|undefined), usageContext: (!Array<string>|undefined),
   *     tags: (!Array|undefined), triggers: (!Array|undefined), variables: (!Array|undefined),
   *     folders: (!Array|undefined), clients: (!Array|undefined),
   *     transformations: (!Array|undefined), templates: (!Array|undefined),
   *     zones: (!Array|undefined), gtagConfigs: (!Array|undefined),
   *     builtInVariables: (!Array|undefined)}=} spec
   * @return {{accountPath: string, containerPath: string, workspacePath: string,
   *     accountId: string, containerId: string, workspaceId: string, publicId: string}}
   */
  seedWorkspace(spec = {}) {
    const accountId = String(spec.accountId || '6000001');
    const containerId = String(spec.containerId || '9000001');
    const workspaceId = String(spec.workspaceId || '1');
    this._seedAccount({accountId, name: spec.accountName});
    this._seedContainer({accountId, containerId, name: spec.containerName, publicId: spec.publicId,
      usageContext: spec.usageContext});
    this._seedWorkspace({accountId, containerId, workspaceId, name: spec.workspaceName});
    const wsPath = this.workspacePath(accountId, containerId, workspaceId);
    const entities = {};
    // Dependency order, so implicit ids are intuitive: folders get the lowest
    // ids, then triggers and variables, then tags and the rest.
    const map = {folders: 'folder', triggers: 'trigger', variables: 'variable', builtInVariables: 'builtInVariable',
      tags: 'tag', clients: 'client', transformations: 'transformation', templates: 'template', zones: 'zone',
      gtagConfigs: 'gtagConfig'};
    for (const [key, kind] of Object.entries(map)) if (spec[key]) entities[kind] = spec[key];
    this.seed({entities: {[wsPath]: entities}});
    const container = this._containers.get(`${accountId}/${containerId}`).data;
    return {
      accountId, containerId, workspaceId,
      accountPath: `accounts/${accountId}`,
      containerPath: `accounts/${accountId}/containers/${containerId}`,
      workspacePath: wsPath,
      publicId: container.publicId,
    };
  }

  _seedAccount(a) {
    const accountId = String(a.accountId || ++this._accountCounter);
    const existing = this._accounts.get(accountId);
    if (existing) {
      if (a.name) existing.name = a.name;
      return existing;
    }
    const account = {
      ...clone(a),
      accountId,
      name: a.name || `Account ${accountId}`,
      path: `accounts/${accountId}`,
      fingerprint: this._fingerprint(),
      tagManagerUrl: `${UI_BASE}/admin/accounts/${accountId}/?apiLink=account`,
    };
    this._accounts.set(accountId, account);
    return account;
  }

  _seedContainer(c) {
    const accountId = String(c.accountId || '6000001');
    this._seedAccount({accountId});
    const containerId = String(c.containerId || this._nextContainerId());
    const key = `${accountId}/${containerId}`;
    const existing = this._containers.get(key);
    if (existing) {
      if (c.name) existing.data.name = c.name;
      if (c.usageContext) existing.data.usageContext = [...c.usageContext];
      if (c.publicId) existing.data.publicId = c.publicId;
      return existing;
    }
    const publicId = c.publicId || `GTM-${Number(containerId).toString(36).toUpperCase().padStart(7, 'X')}`;
    const data = {
      ...clone(c),
      accountId,
      containerId,
      name: c.name || `Container ${containerId}`,
      publicId,
      usageContext: c.usageContext ? [...c.usageContext] : ['web'],
      tagIds: c.tagIds ? [...c.tagIds] : [publicId],
      path: `accounts/${accountId}/containers/${containerId}`,
      fingerprint: this._fingerprint(),
      tagManagerUrl: `${UI_BASE}/container/accounts/${accountId}/containers/${containerId}/workspaces?apiLink=container`,
    };
    const rec = {data, counters: {entity: 0, workspace: 0, destinationLink: 0}, baseEntities: null, baseFrom: null};
    this._containers.set(key, rec);
    return rec;
  }

  _nextContainerId() {
    let max = 9000000;
    for (const rec of this._containers.values()) max = Math.max(max, Number(rec.data.containerId) || 0);
    return String(max + 1);
  }

  _seedWorkspace(w) {
    const accountId = String(w.accountId || '6000001');
    const containerId = String(w.containerId || '9000001');
    const cRec = this._seedContainer({accountId, containerId});
    const workspaceId = String(w.workspaceId || ++cRec.counters.workspace);
    cRec.counters.workspace = Math.max(cRec.counters.workspace, Number(workspaceId) || 0);
    const key = `${accountId}/${containerId}/${workspaceId}`;
    const existing = this._workspaces.get(key);
    if (existing) {
      if (w.name) existing.data.name = w.name;
      return existing;
    }
    const wsPath = this.workspacePath(accountId, containerId, workspaceId);
    const rec = {
      data: {
        ...clone(w),
        accountId, containerId, workspaceId,
        name: w.name || (workspaceId === '1' ? 'Default Workspace' : `Workspace ${workspaceId}`),
        path: wsPath,
        fingerprint: this._fingerprint(),
        tagManagerUrl: `${UI_BASE}/container/${wsPath}?apiLink=workspace`,
      },
      entities: FakeGtmApi._emptyEntities(),
      builtIns: new Map(),
      base: null,
    };
    rec.base = FakeGtmApi._cloneState(rec);
    this._workspaces.set(key, rec);
    return rec;
  }

  _seedEntity(rec, kind, item) {
    const info = ENTITY_KINDS[kind];
    if (this.options.validateSchema) {
      try {
        checkSchema(item, KIND_SCHEMA[kind][0], KIND_SCHEMA[kind][1]);
      } catch (e) {
        throw new Error(`FakeGtmApi.seed: invalid ${kind} ${JSON.stringify(item).slice(0, 120)}: ${e.message}`);
      }
    }
    const cRec = this._containers.get(`${rec.data.accountId}/${rec.data.containerId}`);
    let id = item[info.idField] !== undefined ? String(item[info.idField]) : null;
    if (id === null) id = String(++cRec.counters.entity);
    else if (/^\d+$/.test(id)) cRec.counters.entity = Math.max(cRec.counters.entity, Number(id));
    const entity = this._finalize(rec, kind, this._normalizeBody(clone(item)), id);
    rec.entities[kind].set(id, entity);
  }

  _seedBuiltIn(rec, type) {
    rec.builtIns.set(type, this._builtInEntity(rec, type));
  }

  _seedDestination(d) {
    const accountId = String(d.accountId || '6000001');
    const containerId = String(d.containerId || '9000001');
    const cRec = this._seedContainer({accountId, containerId});
    this._link(cRec, String(d.destinationId), d);
  }

  // -------------------------------------------------------------------------
  // Entity mechanics (operate on a workspace "state": {entities, builtIns})
  // -------------------------------------------------------------------------

  /**
   * Coerces id references to strings and strips server-managed fields.
   * @param {!Object} body
   * @return {!Object}
   */
  _normalizeBody(body) {
    const out = body && typeof body === 'object' && !Array.isArray(body) ? body : {};
    for (const field of Object.keys(REFERENCE_FIELDS)) {
      if (out[field] === undefined || out[field] === null) continue;
      out[field] = Array.isArray(out[field]) ? out[field].map(String) : String(out[field]);
    }
    for (const field of ['scheduleStartMs', 'scheduleEndMs']) {
      if (typeof out[field] === 'number') out[field] = String(out[field]);
    }
    if (this.options.omitEmptyFields) {
      for (const [k, v] of Object.entries(out)) {
        if (v === false || v === '' || v === null || (Array.isArray(v) && v.length === 0)) delete out[k];
      }
    }
    return out;
  }

  /**
   * Adds server-managed fields.
   * @param {!Object} rec Workspace record (for ids/paths).
   * @param {string} kind
   * @param {!Object} entity
   * @param {string} id
   * @return {!Object}
   */
  _finalize(rec, kind, entity, id) {
    const info = ENTITY_KINDS[kind];
    const {accountId, containerId, workspaceId} = rec.data;
    const wsPath = this.workspacePath(accountId, containerId, workspaceId);
    for (const f of ['accountId', 'containerId', 'workspaceId', 'path', 'fingerprint', 'tagManagerUrl']) {
      delete entity[f];
    }
    delete entity[info.idField];
    return {
      accountId, containerId, workspaceId,
      [info.idField]: id,
      ...entity,
      path: `${wsPath}/${info.collection}/${id}`,
      fingerprint: this._fingerprint(),
      tagManagerUrl: `${UI_BASE}/container/${wsPath}/${info.collection}/${id}?apiLink=${kind}`,
    };
  }

  /**
   * Validates name/type/uniqueness (and references unless skipRefs) for a
   * create or update.
   * @param {!Object} state
   * @param {string} kind
   * @param {!Object} entity Normalized body.
   * @param {?string} selfId Id being updated (excluded from uniqueness).
   * @param {boolean=} skipRefs Defer reference checks (bulk_update).
   */
  _validate(state, kind, entity, selfId, skipRefs = false) {
    const info = ENTITY_KINDS[kind];
    if (!info.nameless) {
      if (typeof entity.name !== 'string' || !entity.name.trim()) {
        throw new ApiError(400, `${kindLabel(kind)}.name: The name field is required.`);
      }
      for (const [id, other] of state.entities[kind]) {
        if (id !== selfId && other.name === entity.name) {
          throw new ApiError(400, `Found entity with duplicate name: a ${kindLabel(kind)} named ` +
              `"${entity.name}" already exists in this workspace.`);
        }
      }
    }
    if (info.requiresType && (typeof entity.type !== 'string' || !entity.type)) {
      throw new ApiError(400, `${kindLabel(kind)}.type: The type field is required.`);
    }
    if (!skipRefs) this._validateRefs(state, kind, entity);
  }

  /**
   * Checks that trigger/folder references point at existing entities.
   * @param {!Object} state
   * @param {string} kind
   * @param {!Object} entity
   */
  _validateRefs(state, kind, entity) {
    if (!this.options.validateReferences) return;
    for (const [field, refKind] of Object.entries(REFERENCE_FIELDS)) {
      if (entity[field] === undefined) continue;
      const ids = Array.isArray(entity[field]) ? entity[field] : [entity[field]];
      for (const id of ids) {
        const ok = state.entities[refKind].has(id) ||
            (refKind === 'trigger' && this.options.builtInTriggerIds.includes(id));
        if (!ok) {
          throw new ApiError(400, `${kindLabel(kind)}.${field}: Unknown ${refKind} ID "${id}".`);
        }
      }
    }
  }

  _create(rec, state, kind, body, forcedId, skipRefs = false, where = undefined) {
    this._checkSchema(body, kind, where);
    const entity = this._normalizeBody(clone(body));
    this._validate(state, kind, entity, null, skipRefs);
    const cRec = this._containers.get(`${rec.data.accountId}/${rec.data.containerId}`);
    const id = forcedId || String(++cRec.counters.entity);
    const finalized = this._finalize(rec, kind, entity, id);
    state.entities[kind].set(id, finalized);
    return finalized;
  }

  _update(rec, state, kind, id, body, fingerprint, skipRefs = false, where = undefined) {
    const existing = state.entities[kind].get(id);
    if (!existing) throw new ApiError(404, NOT_FOUND_MESSAGE);
    this._checkFingerprint(existing, fingerprint);
    this._checkSchema(body, kind, where);
    const entity = this._normalizeBody(clone(body));
    this._validate(state, kind, entity, id, skipRefs);
    const finalized = this._finalize(rec, kind, entity, id);
    state.entities[kind].set(id, finalized);
    return finalized;
  }

  /**
   * Schema check for an entity body (when options.validateSchema).
   * @param {*} body
   * @param {string} kind
   * @param {string=} where Error path; defaults to the request field name.
   */
  _checkSchema(body, kind, where) {
    if (!this.options.validateSchema) return;
    const [schemaName, field] = KIND_SCHEMA[kind];
    checkSchema(body, schemaName, where || field);
  }

  _delete(state, kind, id) {
    if (!state.entities[kind].has(id)) throw new ApiError(404, NOT_FOUND_MESSAGE);
    state.entities[kind].delete(id);
    if (kind === 'folder') {
      for (const map of Object.values(state.entities)) {
        for (const e of map.values()) if (e.parentFolderId === id) delete e.parentFolderId;
      }
    }
  }

  _checkFingerprint(existing, fingerprint) {
    if (fingerprint === undefined || fingerprint === null || fingerprint === '') return;
    if (String(fingerprint) !== String(existing.fingerprint)) {
      throw new ApiError(this.options.fingerprintMismatchCode, 'Fingerprint does not match the ' +
          `stored entity (expected ${existing.fingerprint}, got ${fingerprint}). The entity was modified; ` +
          'fetch it again before updating.');
    }
  }

  _builtInEntity(rec, type) {
    const {accountId, containerId, workspaceId} = rec.data;
    const wsPath = this.workspacePath(accountId, containerId, workspaceId);
    return {
      accountId, containerId, workspaceId, type,
      name: builtInName(type),
      path: `${wsPath}/built_in_variables?type=${type}`,
    };
  }

  // -------------------------------------------------------------------------
  // Public helpers for assertions
  // -------------------------------------------------------------------------

  /**
   * @param {string} workspacePath
   * @param {string} kind 'tag' | 'tags' | 'builtInVariable' | ...
   * @return {!Array<!Object>} Copies of the entities, sorted by id.
   */
  list(workspacePath, kind) {
    const {accountId, containerId, workspaceId} = FakeGtmApi.parsePath(workspacePath);
    const rec = this._workspaces.get(`${accountId}/${containerId}/${workspaceId}`);
    if (!rec) throw new Error(`FakeGtmApi.list: unknown workspace ${workspacePath}`);
    const k = normalizeKind(kind);
    if (k === 'builtInVariable') return [...rec.builtIns.values()].map(clone);
    return FakeGtmApi._sorted(rec.entities[k]).map(clone);
  }

  /**
   * Finds an entity by name (or predicate).
   * @param {string} workspacePath
   * @param {string} kind
   * @param {string|function(!Object): boolean} nameOrPredicate
   * @return {?Object}
   */
  find(workspacePath, kind, nameOrPredicate) {
    const pred = typeof nameOrPredicate === 'function' ? nameOrPredicate : (e) => e.name === nameOrPredicate;
    return this.list(workspacePath, kind).find(pred) || null;
  }

  /**
   * Returns a copy of any resource by API path (account, container, workspace,
   * entity, destination), or null.
   * @param {string} resourcePath
   * @return {?Object}
   */
  get(resourcePath) {
    const p = String(resourcePath).replace(/^\/+|\/+$/g, '');
    let m = /^accounts\/([^/]+)$/.exec(p);
    if (m) return clone(this._accounts.get(m[1]) || null);
    m = /^accounts\/([^/]+)\/containers\/([^/]+)$/.exec(p);
    if (m) return clone((this._containers.get(`${m[1]}/${m[2]}`) || {}).data || null);
    m = /^accounts\/([^/]+)\/containers\/([^/]+)\/destinations\/([^/]+)$/.exec(p);
    if (m) return clone(((this._destinations.get(`${m[1]}/${m[2]}`) || new Map()).get(m[3])) || null);
    m = /^accounts\/([^/]+)\/containers\/([^/]+)\/workspaces\/([^/]+)$/.exec(p);
    if (m) return clone((this._workspaces.get(`${m[1]}/${m[2]}/${m[3]}`) || {}).data || null);
    m = /^accounts\/([^/]+)\/containers\/([^/]+)\/workspaces\/([^/]+)\/([a-z_]+)\/([^/]+)$/.exec(p);
    if (m && COLLECTION_TO_KIND[m[4]]) {
      const rec = this._workspaces.get(`${m[1]}/${m[2]}/${m[3]}`);
      return rec ? clone(rec.entities[COLLECTION_TO_KIND[m[4]]].get(m[5]) || null) : null;
    }
    return null;
  }

  /**
   * Changes relative to the seeded base (like workspaces.getStatus).
   * @param {string} workspacePath
   * @return {!Array<{kind: string, id: string, name: (string|undefined), changeStatus: string, entity: !Object}>}
   */
  workspaceChanges(workspacePath) {
    const {accountId, containerId, workspaceId} = FakeGtmApi.parsePath(workspacePath);
    const rec = this._workspaces.get(`${accountId}/${containerId}/${workspaceId}`);
    if (!rec) throw new Error(`FakeGtmApi.workspaceChanges: unknown workspace ${workspacePath}`);
    return this._diff(rec).map((c) => ({...c, entity: clone(c.entity)}));
  }

  /**
   * Live plain-data snapshot of everything (deep copy).
   * @return {!Object}
   */
  get state() {
    const entities = {};
    for (const rec of this._workspaces.values()) {
      const out = {};
      for (const kind of Object.keys(ENTITY_KINDS)) out[kind] = FakeGtmApi._sorted(rec.entities[kind]).map(clone);
      out.builtInVariable = [...rec.builtIns.values()].map(clone);
      entities[rec.data.path] = out;
    }
    const destinations = {};
    for (const [key, map] of this._destinations) {
      destinations[`accounts/${key.split('/')[0]}/containers/${key.split('/')[1]}`] = [...map.values()].map(clone);
    }
    return {
      accounts: [...this._accounts.values()].map(clone),
      containers: [...this._containers.values()].map((r) => clone(r.data)),
      workspaces: [...this._workspaces.values()].map((r) => clone(r.data)),
      entities,
      destinations,
    };
  }

  /** @return {!Object} An opaque deep snapshot for restore(). */
  snapshot() {
    return {
      accounts: clone([...this._accounts]),
      containers: [...this._containers].map(([k, r]) => [k, {...clone({data: r.data, counters: r.counters,
        baseFrom: r.baseFrom}), baseEntities: r.baseEntities ? FakeGtmApi._cloneState(r.baseEntities) : null}]),
      workspaces: [...this._workspaces].map(([k, r]) => [k, {data: clone(r.data),
        ...FakeGtmApi._cloneState(r), base: r.base ? FakeGtmApi._cloneState(r.base) : null}]),
      destinations: [...this._destinations].map(([k, m]) => [k, clone([...m])]),
      accountCounter: this._accountCounter,
    };
  }

  /** @param {!Object} snap From snapshot(). */
  restore(snap) {
    this._accounts = new Map(clone(snap.accounts));
    this._containers = new Map(snap.containers.map(([k, r]) => [k, {...clone({data: r.data,
      counters: r.counters, baseFrom: r.baseFrom}), baseEntities: r.baseEntities ?
      FakeGtmApi._cloneState(r.baseEntities) : null}]));
    this._workspaces = new Map(snap.workspaces.map(([k, r]) => [k, {data: clone(r.data),
      ...FakeGtmApi._cloneState(r), base: r.base ? FakeGtmApi._cloneState(r.base) : null}]));
    this._destinations = new Map(snap.destinations.map(([k, entries]) => [k, new Map(clone(entries))]));
    this._accountCounter = snap.accountCounter;
  }

  /**
   * @param {string=} method
   * @param {(string|!RegExp)=} pathPattern API-relative path (exact string or regex).
   * @return {!Array<!Object>} Matching entries of fake.requests.
   */
  requestsTo(method, pathPattern) {
    return this.requests.filter((r) => (!method || r.method === method.toUpperCase()) &&
        (!pathPattern || (pathPattern instanceof RegExp ? pathPattern.test(r.path) : r.path === pathPattern)));
  }

  /**
   * Makes matching requests fail. Matching requests are consumed `times`
   * times (default 1; Infinity for always).
   * @param {{method: (string|undefined), path: (string|!RegExp|undefined),
   *     match: (function(!Object): boolean|undefined), code: (number|undefined),
   *     message: (string|undefined), status: (string|undefined), reason: (string|undefined),
   *     times: (number|undefined), networkError: (string|undefined)}} spec
   * @return {!FakeGtmApi}
   */
  injectError(spec) {
    this._errorInjections.push({times: 1, code: 500, ...spec});
    return this;
  }

  clearInjectedErrors() {
    this._errorInjections = [];
  }

  // -------------------------------------------------------------------------
  // HTTP handling
  // -------------------------------------------------------------------------

  /**
   * Handles a request in the runtime's fetchHandler format.
   * @param {{url: string, method: string, headers: (!Object|undefined), payload: (?string|undefined),
   *     contentType: (?string|undefined), timestamp: (number|undefined)}} req
   * @return {({code: number, body: string, headers: !Object}|{networkError: string}|undefined)}
   *     undefined for non-GTM URLs (so other handlers / real HTTP can take over).
   */
  handle(req) {
    let url;
    try {
      url = new URL(req.url);
    } catch (e) {
      return undefined;
    }
    if (url.origin !== GTM_API_ORIGIN || !url.pathname.startsWith(GTM_API_PREFIX)) return undefined;
    const method = String(req.method || 'GET').toUpperCase();
    const apiPath = url.pathname.slice(GTM_API_PREFIX.length).replace(/\/+$/, '')
        .split('/').map((s) => decodeURIComponent(s)).join('/');
    const query = {};
    for (const [k, v] of url.searchParams) (query[k] = query[k] || []).push(v);
    const q1 = (k) => (query[k] ? query[k][0] : undefined);
    const header = (name) => {
      if (typeof req.header === 'function') return req.header(name);
      const lower = name.toLowerCase();
      for (const [k, v] of Object.entries(req.headers || {})) if (k.toLowerCase() === lower) return v;
      if (lower === 'content-type') return req.contentType || undefined;
      return undefined;
    };
    const entry = {
      method, url: req.url, path: apiPath, query, body: undefined, rawBody: req.payload ?? null,
      status: null, response: undefined, timestamp: req.timestamp ?? Date.now(),
    };
    this.requests.push(entry);

    const respond = (code, obj) => {
      entry.status = code;
      entry.response = obj;
      if (code === 204 || obj === undefined) return {code, body: '', headers: {}};
      return {
        code,
        body: JSON.stringify(obj, null, 2),
        headers: {'Content-Type': 'application/json; charset=UTF-8', 'Vary': 'Origin, X-Origin, Referer'},
      };
    };
    const fail = (err) => {
      const error = {code: err.code, message: err.message,
        errors: [{message: err.message, domain: 'global', reason: err.reason, ...(err.extraErrorFields || {})}],
        status: err.status};
      if (err.details) error.details = err.details;
      return respond(err.code, {error});
    };

    try {
      // Injected failures.
      const inj = this._errorInjections.find((i) => i.times > 0 &&
          (!i.method || i.method.toUpperCase() === method) &&
          (!i.path || (i.path instanceof RegExp ? i.path.test(apiPath) : i.path === apiPath)) &&
          (!i.match || i.match({...entry, req})));
      if (inj) {
        inj.times--;
        if (inj.networkError) {
          entry.status = 'networkError';
          return {networkError: inj.networkError};
        }
        throw new ApiError(inj.code, inj.message || 'Injected error', {status: inj.status, reason: inj.reason});
      }
      // Rate limiting.
      if (this.options.rateLimit) {
        const {max, windowMs} = this.options.rateLimit;
        const t = entry.timestamp;
        this._rateTimes = this._rateTimes.filter((x) => t - x < windowMs);
        if (this._rateTimes.length >= max) {
          throw new ApiError(this.options.rateLimit.code || 429, 'Quota exceeded for quota metric ' +
              '\'Queries\' and limit \'Queries per minute per user\' of service \'tagmanager.googleapis.com\' ' +
              'for consumer \'project_number:000000000000\'.', {
            status: 'RESOURCE_EXHAUSTED', reason: 'rateLimitExceeded',
            details: [{'@type': 'type.googleapis.com/google.rpc.ErrorInfo', reason: 'RATE_LIMIT_EXCEEDED',
              domain: 'googleapis.com', metadata: {service: 'tagmanager.googleapis.com',
                quota_metric: 'tagmanager.googleapis.com/queries'}}],
          });
        }
        this._rateTimes.push(t);
      }
      // Authentication.
      if (this.options.requireAuth) {
        const auth = header('authorization');
        const m = /^Bearer\s+(\S+)$/i.exec(String(auth || '').trim());
        if (!m) {
          throw new ApiError(401, 'Request is missing required authentication credential. Expected OAuth 2 ' +
              'access token, login cookie or other valid authentication credential. See ' +
              'https://developers.google.com/identity/sign-in/web/devconsole-project.',
          {reason: 'required', extraErrorFields: {location: 'Authorization', locationType: 'header'}});
        }
        if (this.options.acceptedTokens && !this.options.acceptedTokens.includes(m[1])) {
          throw new ApiError(401, 'Request had invalid authentication credentials. Expected OAuth 2 access ' +
              'token, login cookie or other valid authentication credential. See ' +
              'https://developers.google.com/identity/sign-in/web/devconsole-project.', {reason: 'authError'});
        }
      }
      // Body parsing.
      let body;
      const raw = req.payload;
      if ((method === 'POST' || method === 'PUT') && raw !== undefined && raw !== null && raw !== '') {
        const ct = String(header('content-type') || '');
        if (this.options.strictContentType && !/application\/json/i.test(ct)) {
          const name = String(raw).slice(0, 100);
          throw new ApiError(400, `Invalid JSON payload received. Unknown name "${name}": Cannot bind query ` +
              `parameter. Field '${name}' could not be found in request message.`, {reason: 'invalid'});
        }
        try {
          body = JSON.parse(raw);
        } catch (e) {
          throw new ApiError(400, `Invalid JSON payload received. ${e.message}`, {reason: 'parseError'});
        }
        entry.body = body;
      }
      const result = this._route(method, apiPath, query, q1, body);
      return respond(result.code || 200, result.body);
    } catch (e) {
      if (e instanceof ApiError) return fail(e);
      throw e;
    }
  }

  /**
   * Dispatches to the handler for METHOD + path.
   * @return {{code: (number|undefined), body: *}}
   */
  _route(method, p, query, q1, body) {
    const ID = '([^/:]+)';
    const C = `^accounts/${ID}/containers/${ID}`;
    const W = `${C}/workspaces/${ID}`;
    const COLL = `(${Object.values(ENTITY_KINDS).map((i) => i.collection).join('|')})`;
    const routes = [
      ['GET', '^accounts$', () => this._page([...this._accounts.values()], 'account', q1)],
      ['GET', '^accounts/containers:lookup$', () => this._lookup(q1)],
      ['GET', `^accounts/${ID}$`, (a) => ({body: this._getAccount(a)})],
      ['PUT', `^accounts/${ID}$`, (a) => ({body: this._updateAccount(a, body, q1('fingerprint'))})],
      ['GET', `^accounts/${ID}/containers$`, (a) => {
        this._getAccount(a);
        return this._page([...this._containers.values()].filter((r) => r.data.accountId === a)
            .map((r) => r.data), 'container', q1);
      }],
      ['POST', `^accounts/${ID}/containers$`, (a) => ({body: this._createContainer(a, body)})],
      ['GET', `${C}$`, (a, c) => ({body: clone(this._containerRec(a, c).data)})],
      ['PUT', `${C}$`, (a, c) => ({body: this._updateContainer(a, c, body, q1('fingerprint'))})],
      ['DELETE', `${C}$`, (a, c) => this._deleteContainer(a, c)],
      ['GET', `${C}/destinations$`, (a, c) => {
        this._containerRec(a, c);
        return this._page([...(this._destinations.get(`${a}/${c}`) || new Map()).values()], 'destination', q1);
      }],
      ['POST', `${C}/destinations:link$`, (a, c) => {
        const id = q1('destinationId');
        if (!id) throw new ApiError(400, 'destinationId is required.');
        return {body: clone(this._link(this._containerRec(a, c), id, {}))};
      }],
      ['GET', `${C}/destinations/${ID}$`, (a, c, d) => {
        this._containerRec(a, c);
        const dest = (this._destinations.get(`${a}/${c}`) || new Map()).get(d);
        if (!dest) throw new ApiError(404, NOT_FOUND_MESSAGE);
        return {body: clone(dest)};
      }],
      ['GET', `${C}/workspaces$`, (a, c) => {
        this._containerRec(a, c);
        return this._page([...this._workspaces.values()].filter((r) => r.data.accountId === a &&
            r.data.containerId === c).map((r) => r.data), 'workspace', q1);
      }],
      ['POST', `${C}/workspaces$`, (a, c) => ({body: this._createWorkspace(a, c, body)})],
      ['GET', `${W}$`, (a, c, w) => ({body: clone(this._workspaceRec(a, c, w).data)})],
      ['PUT', `${W}$`, (a, c, w) => ({body: this._updateWorkspace(a, c, w, body, q1('fingerprint'))})],
      ['DELETE', `${W}$`, (a, c, w) => {
        this._workspaceRec(a, c, w);
        this._workspaces.delete(`${a}/${c}/${w}`);
        return {code: 204};
      }],
      ['GET', `${W}/status$`, (a, c, w) => ({body: this._status(this._workspaceRec(a, c, w))})],
      ['POST', `${W}/bulk_update$`, (a, c, w) => ({body: this._bulkUpdate(this._workspaceRec(a, c, w), body)})],
      ['GET', `${W}/built_in_variables$`, (a, c, w) =>
        this._page([...this._workspaceRec(a, c, w).builtIns.values()], 'builtInVariable', q1)],
      ['POST', `${W}/built_in_variables$`, (a, c, w) =>
        ({body: this._enableBuiltIns(this._workspaceRec(a, c, w), query.type || [])})],
      ['DELETE', `${W}/built_in_variables$`, (a, c, w) => {
        this._disableBuiltIns(this._workspaceRec(a, c, w), query.type || []);
        return {code: 204};
      }],
      ['POST', `${W}/built_in_variables:revert$`, (a, c, w) =>
        ({body: this._revertBuiltIn(this._workspaceRec(a, c, w), q1('type'))})],
      ['GET', `${W}/${COLL}$`, (a, c, w, coll) => {
        const rec = this._workspaceRec(a, c, w);
        const kind = COLLECTION_TO_KIND[coll];
        return this._page(FakeGtmApi._sorted(rec.entities[kind]), ENTITY_KINDS[kind].listKey, q1);
      }],
      ['POST', `${W}/${COLL}$`, (a, c, w, coll) => {
        const rec = this._workspaceRec(a, c, w);
        return {body: clone(this._create(rec, rec, COLLECTION_TO_KIND[coll], body || {}))};
      }],
      ['GET', `${W}/${COLL}/${ID}$`, (a, c, w, coll, id) => {
        const e = this._workspaceRec(a, c, w).entities[COLLECTION_TO_KIND[coll]].get(id);
        if (!e) throw new ApiError(404, NOT_FOUND_MESSAGE);
        return {body: clone(e)};
      }],
      ['PUT', `${W}/${COLL}/${ID}$`, (a, c, w, coll, id) => {
        const rec = this._workspaceRec(a, c, w);
        return {body: clone(this._update(rec, rec, COLLECTION_TO_KIND[coll], id, body || {}, q1('fingerprint')))};
      }],
      ['DELETE', `${W}/${COLL}/${ID}$`, (a, c, w, coll, id) => {
        this._delete(this._workspaceRec(a, c, w), COLLECTION_TO_KIND[coll], id);
        return {code: 204};
      }],
      ['POST', `${W}/${COLL}/${ID}:revert$`, (a, c, w, coll, id) =>
        ({body: this._revert(this._workspaceRec(a, c, w), COLLECTION_TO_KIND[coll], id, q1('fingerprint'))})],
      ['POST', `${W}/folders/${ID}:entities$`, (a, c, w, id) =>
        ({body: this._folderEntities(this._workspaceRec(a, c, w), id)})],
      ['POST', `${W}/folders/${ID}:move_entities_to_folder$`, (a, c, w, id) => {
        this._moveToFolder(this._workspaceRec(a, c, w), id, query);
        return {code: 204};
      }],
    ];
    let pathMatched = false;
    for (const [m, pattern, fn] of routes) {
      const match = new RegExp(pattern).exec(p);
      if (!match) continue;
      pathMatched = true;
      if (m !== method) continue;
      return fn(...match.slice(1));
    }
    throw new ApiError(pathMatched ? 405 : 501, `FakeGtmApi does not implement ${method} ${p}`,
        pathMatched ? {status: 'METHOD_NOT_ALLOWED', reason: 'methodNotAllowed'} : {});
  }

  /** @return {!Array} Map values sorted by numeric id. */
  static _sorted(map) {
    return [...map.entries()].sort((x, y) => {
      const nx = Number(x[0]);
      const ny = Number(y[0]);
      if (!Number.isNaN(nx) && !Number.isNaN(ny)) return nx - ny;
      return x[0] < y[0] ? -1 : 1;
    }).map(([, v]) => v);
  }

  /**
   * Paginates a list response; omits the list key when empty.
   * @return {{body: !Object}}
   */
  _page(items, listKey, q1) {
    const size = this.options.pageSize;
    let offset = 0;
    const token = q1('pageToken');
    if (token) {
      let decoded = null;
      try {
        decoded = JSON.parse(Buffer.from(token, 'base64url').toString('utf8'));
      } catch (e) {
        decoded = null;
      }
      if (!decoded || decoded.k !== listKey || typeof decoded.o !== 'number') {
        throw new ApiError(400, 'Invalid page token.', {reason: 'invalidParameter'});
      }
      offset = decoded.o;
    }
    const page = items.slice(offset, offset + size).map(clone);
    const body = {};
    if (page.length) body[listKey] = page;
    if (offset + size < items.length) {
      body.nextPageToken = Buffer.from(JSON.stringify({o: offset + size, k: listKey})).toString('base64url');
    }
    return {body};
  }

  // --- accounts & containers -------------------------------------------------------

  _getAccount(a) {
    const account = this._accounts.get(a);
    if (!account) throw new ApiError(404, NOT_FOUND_MESSAGE);
    return clone(account);
  }

  _updateAccount(a, body, fingerprint) {
    const account = this._accounts.get(a);
    if (!account) throw new ApiError(404, NOT_FOUND_MESSAGE);
    this._checkFingerprint(account, fingerprint);
    if (this.options.validateSchema) checkSchema(body, 'Account', 'account');
    if (!body || !body.name) throw new ApiError(400, 'account.name: The name field is required.');
    Object.assign(account, {name: body.name, shareData: Boolean(body.shareData), fingerprint: this._fingerprint()});
    return clone(account);
  }

  _lookup(q1) {
    const tagId = q1('tagId');
    const destinationId = q1('destinationId');
    if (Boolean(tagId) === Boolean(destinationId)) {
      throw new ApiError(400, 'Exactly one of tagId or destinationId must be set.');
    }
    for (const rec of this._containers.values()) {
      const d = rec.data;
      if (tagId && (d.publicId === tagId || (d.tagIds || []).includes(tagId))) return {body: clone(d)};
      if (destinationId) {
        const dests = this._destinations.get(`${d.accountId}/${d.containerId}`);
        if ((dests && [...dests.values()].some((x) => x.destinationId === destinationId)) ||
            (d.tagIds || []).includes(destinationId)) {
          return {body: clone(d)};
        }
      }
    }
    throw new ApiError(404, NOT_FOUND_MESSAGE);
  }

  _createContainer(a, body) {
    this._getAccount(a);
    if (this.options.validateSchema) checkSchema(body, 'Container', 'container');
    if (!body || !body.name) throw new ApiError(400, 'container.name: The name field is required.');
    if (!Array.isArray(body.usageContext) || body.usageContext.length === 0) {
      throw new ApiError(400, 'container.usageContext: The usage context is required.');
    }
    const rec = this._seedContainer({...body, accountId: a, containerId: this._nextContainerId(),
      publicId: undefined});
    this._seedWorkspace({accountId: a, containerId: rec.data.containerId, workspaceId: '1',
      name: 'Default Workspace'});
    return clone(rec.data);
  }

  _updateContainer(a, c, body, fingerprint) {
    const rec = this._containerRec(a, c);
    this._checkFingerprint(rec.data, fingerprint);
    if (this.options.validateSchema) checkSchema(body, 'Container', 'container');
    if (!body || !body.name) throw new ApiError(400, 'container.name: The name field is required.');
    const keep = ['accountId', 'containerId', 'publicId', 'path', 'tagManagerUrl', 'tagIds'];
    const next = {...clone(body)};
    for (const k of keep) next[k] = rec.data[k];
    next.fingerprint = this._fingerprint();
    rec.data = next;
    return clone(next);
  }

  _deleteContainer(a, c) {
    this._containerRec(a, c);
    this._containers.delete(`${a}/${c}`);
    this._destinations.delete(`${a}/${c}`);
    for (const key of [...this._workspaces.keys()]) if (key.startsWith(`${a}/${c}/`)) this._workspaces.delete(key);
    return {code: 204};
  }

  /**
   * Links a destination to a container, unlinking it elsewhere.
   * @return {!Object} The destination.
   */
  _link(cRec, destinationId, extra) {
    for (const [key, map] of this._destinations) {
      for (const [linkId, dest] of map) {
        if (dest.destinationId === destinationId) {
          map.delete(linkId);
          const other = this._containers.get(key);
          if (other) other.data.tagIds = (other.data.tagIds || []).filter((t) => t !== destinationId);
        }
      }
    }
    const {accountId, containerId} = cRec.data;
    const key = `${accountId}/${containerId}`;
    if (!this._destinations.has(key)) this._destinations.set(key, new Map());
    const linkId = String(++cRec.counters.destinationLink);
    const containerPath = `accounts/${accountId}/containers/${containerId}`;
    const dest = {
      accountId, containerId, destinationId,
      destinationLinkId: linkId,
      name: extra.name || destinationId,
      path: `${containerPath}/destinations/${linkId}`,
      fingerprint: this._fingerprint(),
      tagManagerUrl: `${UI_BASE}/container/${containerPath}/destinations/${linkId}?apiLink=destination`,
    };
    this._destinations.get(key).set(linkId, dest);
    if (!(cRec.data.tagIds || []).includes(destinationId)) {
      cRec.data.tagIds = [...(cRec.data.tagIds || []), destinationId];
    }
    return dest;
  }

  // --- workspaces ------------------------------------------------------------------

  _createWorkspace(a, c, body) {
    const cRec = this._containerRec(a, c);
    if (this.options.validateSchema) checkSchema(body, 'Workspace', 'workspace');
    if (!body || !body.name) throw new ApiError(400, 'workspace.name: The name field is required.');
    const existing = [...this._workspaces.values()].filter((r) => r.data.accountId === a && r.data.containerId === c);
    if (existing.some((r) => r.data.name === body.name)) {
      throw new ApiError(400, `Found entity with duplicate name: a workspace named "${body.name}" already exists.`);
    }
    if (existing.length >= this.options.maxWorkspacesPerContainer) {
      throw new ApiError(400, `The maximum number of workspaces (${this.options.maxWorkspacesPerContainer}) ` +
          'for this container has been reached.', {reason: 'limitExceeded'});
    }
    const rec = this._seedWorkspace({accountId: a, containerId: c, name: body.name,
      description: body.description, workspaceId: String(cRec.counters.workspace + 1)});
    if (cRec.baseEntities) {
      // A new workspace starts from the container's latest version.
      const copy = FakeGtmApi._cloneState(cRec.baseEntities);
      for (const [kind, map] of Object.entries(copy.entities)) {
        for (const [id, e] of map) {
          map.set(id, this._finalize(rec, kind, {...e}, id));
        }
      }
      rec.entities = copy.entities;
      rec.builtIns = new Map([...copy.builtIns.keys()].map((t) => [t, this._builtInEntity(rec, t)]));
      rec.base = FakeGtmApi._cloneState(rec);
    }
    return clone(rec.data);
  }

  _updateWorkspace(a, c, w, body, fingerprint) {
    const rec = this._workspaceRec(a, c, w);
    this._checkFingerprint(rec.data, fingerprint);
    if (this.options.validateSchema) checkSchema(body, 'Workspace', 'workspace');
    if (!body || !body.name) throw new ApiError(400, 'workspace.name: The name field is required.');
    rec.data = {...rec.data, name: body.name, description: body.description, fingerprint: this._fingerprint()};
    if (rec.data.description === undefined) delete rec.data.description;
    return clone(rec.data);
  }

  /**
   * Diffs current entities against the base.
   * @param {!Object} rec
   * @return {!Array<!Object>}
   */
  _diff(rec) {
    const changes = [];
    const base = rec.base || {entities: FakeGtmApi._emptyEntities(), builtIns: new Map()};
    for (const kind of Object.keys(ENTITY_KINDS)) {
      const cur = rec.entities[kind];
      const old = base.entities[kind];
      const ids = [...new Set([...cur.keys(), ...old.keys()])]
          .sort((x, y) => (Number(x) - Number(y)) || (x < y ? -1 : 1));
      for (const id of ids) {
        const now = cur.get(id);
        const was = old.get(id);
        let status = null;
        if (now && !was) status = 'added';
        else if (!now && was) status = 'deleted';
        else if (now.fingerprint !== was.fingerprint) status = 'updated';
        if (status) changes.push({kind, id, name: (now || was).name, changeStatus: status, entity: now || was});
      }
    }
    const types = [...new Set([...rec.builtIns.keys(), ...base.builtIns.keys()])].sort();
    for (const t of types) {
      const now = rec.builtIns.get(t);
      const was = base.builtIns.get(t);
      if (now && !was) changes.push({kind: 'builtInVariable', id: t, name: now.name, changeStatus: 'added', entity: now});
      if (!now && was) changes.push({kind: 'builtInVariable', id: t, name: was.name, changeStatus: 'deleted', entity: was});
    }
    return changes;
  }

  _status(rec) {
    const changes = this._diff(rec).map((c) => ({
      [c.kind === 'builtInVariable' ? 'builtInVariable' : ENTITY_KINDS[c.kind].entityKey]: clone(c.entity),
      changeStatus: c.changeStatus,
    }));
    return changes.length ? {workspaceChange: changes} : {};
  }

  _revert(rec, kind, id, fingerprint) {
    const cur = rec.entities[kind].get(id);
    const was = rec.base ? rec.base.entities[kind].get(id) : undefined;
    if (!cur && !was) throw new ApiError(404, NOT_FOUND_MESSAGE);
    if (cur) this._checkFingerprint(cur, fingerprint);
    const key = ENTITY_KINDS[kind].entityKey === 'customTemplate' ? 'template' : ENTITY_KINDS[kind].listKey;
    if (!was) {
      rec.entities[kind].delete(id);
      return {};
    }
    rec.entities[kind].set(id, clone(was));
    return {[key]: clone(was)};
  }

  _folderEntities(rec, folderId) {
    if (!rec.entities.folder.has(folderId)) throw new ApiError(404, NOT_FOUND_MESSAGE);
    const out = {};
    for (const kind of ['tag', 'trigger', 'variable']) {
      const items = FakeGtmApi._sorted(rec.entities[kind]).filter((e) => e.parentFolderId === folderId).map(clone);
      if (items.length) out[kind] = items;
    }
    return out;
  }

  _moveToFolder(rec, folderId, query) {
    if (!rec.entities.folder.has(folderId)) throw new ApiError(404, NOT_FOUND_MESSAGE);
    const plan = [['tag', query.tagId || []], ['trigger', query.triggerId || []], ['variable', query.variableId || []]];
    for (const [kind, ids] of plan) {
      for (const id of ids) if (!rec.entities[kind].has(id)) throw new ApiError(404, NOT_FOUND_MESSAGE);
    }
    for (const [kind, ids] of plan) {
      for (const id of ids) {
        const e = rec.entities[kind].get(id);
        rec.entities[kind].set(id, {...e, parentFolderId: folderId, fingerprint: this._fingerprint()});
      }
    }
  }

  // --- built-in variables ------------------------------------------------------------

  _enableBuiltIns(rec, types) {
    if (!types.length) throw new ApiError(400, 'At least one built-in variable type is required.');
    const created = [];
    for (const t of types) {
      if (!BUILT_IN_VARIABLE_TYPES.includes(t)) {
        throw new ApiError(400, `Invalid value at 'type' (TYPE_ENUM), "${t}"`, {reason: 'invalid'});
      }
    }
    for (const t of types) {
      if (!rec.builtIns.has(t)) rec.builtIns.set(t, this._builtInEntity(rec, t));
      created.push(clone(rec.builtIns.get(t)));
    }
    return {builtInVariable: created};
  }

  _disableBuiltIns(rec, types) {
    if (!types.length) throw new ApiError(400, 'At least one built-in variable type is required.');
    for (const t of types) if (!rec.builtIns.has(t)) throw new ApiError(404, NOT_FOUND_MESSAGE);
    for (const t of types) rec.builtIns.delete(t);
  }

  _revertBuiltIn(rec, type) {
    if (!type) throw new ApiError(400, 'type is required.');
    const enabled = rec.base && rec.base.builtIns.has(type);
    if (enabled) rec.builtIns.set(type, clone(rec.base.builtIns.get(type)));
    else rec.builtIns.delete(type);
    return {enabled: Boolean(enabled)};
  }

  // --- bulk update -----------------------------------------------------------------

  /**
   * Atomic bulk update with "new_N" ids. Validates everything against a copy
   * of the workspace; on the first error nothing is applied and a 400 naming
   * the failing change index is returned.
   * @param {!Object} rec Workspace record.
   * @param {*} body {changes: [{tag|trigger|...: {...}, changeStatus}]}
   * @return {!Object} {changes: [...]} (added/updated entities in request order).
   */
  _bulkUpdate(rec, body) {
    if (this.options.validateSchema && body && typeof body === 'object') {
      for (const key of Object.keys(body)) {
        if (key !== 'changes') {
          throw new ApiError(400, `Invalid JSON payload received. Unknown name "${key}": Cannot find field.`,
              {reason: 'invalid'});
        }
      }
    }
    if (!body || !Array.isArray(body.changes)) {
      throw new ApiError(400, 'changes: must be a list of entity changes.');
    }
    const cRec = this._containers.get(`${rec.data.accountId}/${rec.data.containerId}`);
    const sim = FakeGtmApi._cloneState(rec);
    let nextId = cRec.counters.entity;
    const newIds = new Map();  // 'new_1' -> {kind, realId}
    const fail = (i, msg) => {
      throw new ApiError(400, `changes[${i}]: ${msg}`);
    };

    // Pass 1: shape checks and id allocation for added entities.
    const parsed = body.changes.map((change, i) => {
      if (!change || typeof change !== 'object') fail(i, 'change must be an object.');
      if (this.options.validateSchema) {
        for (const key of Object.keys(change)) {
          if (!GTM_SCHEMAS.Entity[key]) {
            throw new ApiError(400, `Invalid JSON payload received. Unknown name "${key}" at 'changes[${i}]': ` +
                'Cannot find field.', {reason: 'invalid'});
          }
        }
      }
      const keys = Object.keys(change).filter((k) => k !== 'changeStatus' && ENTITY_KEY_TO_KIND[k] &&
          change[k] !== undefined && change[k] !== null);
      if (keys.length !== 1) {
        fail(i, `exactly one entity (${Object.keys(ENTITY_KEY_TO_KIND).join(', ')}) must be set; got ` +
            `${keys.length ? keys.join(', ') : 'none'}.`);
      }
      const kind = ENTITY_KEY_TO_KIND[keys[0]];
      const status = change.changeStatus;
      if (!['added', 'updated', 'deleted'].includes(status)) {
        fail(i, `changeStatus must be one of added, updated, deleted (got ${JSON.stringify(status)}).`);
      }
      const entity = clone(change[keys[0]]);
      if (kind === 'builtInVariable') {
        if (status === 'updated') fail(i, 'built-in variables cannot be updated.');
        if (!entity.type) fail(i, 'builtInVariable.type is required.');
        return {i, kind, status, entity, id: entity.type, where: `changes[${i}].built_in_variable`};
      }
      const idField = ENTITY_KINDS[kind].idField;
      const rawId = entity[idField] === undefined || entity[idField] === null ? '' : String(entity[idField]);
      if (status === 'added') {
        if (!/^new_\d+$/.test(rawId)) {
          fail(i, `${kind}.${idField} for an added entity must be "new_" followed by a number (got ` +
              `${JSON.stringify(rawId)}).`);
        }
        if (newIds.has(rawId)) fail(i, `duplicate new entity id "${rawId}".`);
        newIds.set(rawId, {kind, realId: String(++nextId)});
      } else if (!rawId) {
        fail(i, `${kind}.${idField} is required for ${status} changes.`);
      }
      return {i, kind, status, entity, id: rawId, where: `changes[${i}].${snake(keys[0])}`};
    });

    // Pass 2: apply in request order on the copy (reference checks deferred so
    // entities may reference "new_N" ids that are created by later changes).
    const results = [];
    const applied = [];
    const wrap = (p, e) => {
      if (!(e instanceof ApiError)) return e;
      return new ApiError(e.code === 404 ? 404 : 400,
          `changes[${p.i}] (${p.status} ${kindLabel(p.kind)} ${p.id}): ${e.message}`,
          {status: e.status, reason: e.reason});
    };
    for (const p of parsed) {
      try {
        if (p.kind === 'builtInVariable') {
          if (this.options.validateSchema) checkSchema(p.entity, 'BuiltInVariable', p.where);
          if (!BUILT_IN_VARIABLE_TYPES.includes(p.id)) throw new ApiError(400, `Invalid built-in variable type "${p.id}".`);
          if (p.status === 'added') {
            sim.builtIns.set(p.id, this._builtInEntity(rec, p.id));
            results.push({builtInVariable: clone(sim.builtIns.get(p.id)), changeStatus: 'added'});
          } else {
            if (!sim.builtIns.has(p.id)) throw new ApiError(404, NOT_FOUND_MESSAGE);
            sim.builtIns.delete(p.id);
          }
          continue;
        }
        for (const [field, refKind] of Object.entries(REFERENCE_FIELDS)) {
          if (p.entity[field] === undefined || p.entity[field] === null) continue;
          const remap = (ref) => {
            const s = String(ref);
            if (!/^new_\d+$/.test(s)) return s;
            const target = newIds.get(s);
            if (!target) throw new ApiError(400, `${field} references unknown new entity id "${s}".`);
            if (target.kind !== refKind) {
              throw new ApiError(400, `${field} references "${s}", which is a ${target.kind}, not a ${refKind}.`);
            }
            return target.realId;
          };
          p.entity[field] = Array.isArray(p.entity[field]) ? p.entity[field].map(remap) : remap(p.entity[field]);
        }
        const key = ENTITY_KINDS[p.kind].entityKey;
        if (p.status === 'added') {
          const created = this._create(rec, sim, p.kind, p.entity, newIds.get(p.id).realId, true, p.where);
          results.push({[key]: created, changeStatus: 'added'});
          applied.push({p, entity: created});
        } else if (p.status === 'updated') {
          const updated = this._update(rec, sim, p.kind, p.id, p.entity, undefined, true, p.where);
          results.push({[key]: updated, changeStatus: 'updated'});
          applied.push({p, entity: updated});
        } else {
          this._delete(sim, p.kind, p.id);
        }
      } catch (e) {
        throw wrap(p, e);
      }
    }
    // Pass 3: references must resolve in the final state.
    for (const {p, entity} of applied) {
      try {
        if (sim.entities[p.kind].get(entity[ENTITY_KINDS[p.kind].idField]) === entity) {
          this._validateRefs(sim, p.kind, entity);
        }
      } catch (e) {
        throw wrap(p, e);
      }
    }
    // Commit atomically.
    rec.entities = sim.entities;
    rec.builtIns = sim.builtIns;
    cRec.counters.entity = nextId;
    return results.length ? {changes: clone(results)} : {};
  }
}

/**
 * Creates a fake GTM API (see FakeGtmApi constructor for options).
 * @param {!Object=} options
 * @return {!FakeGtmApi}
 */
function createFakeGtmApi(options) {
  return new FakeGtmApi(options);
}

module.exports = {
  createFakeGtmApi,
  FakeGtmApi,
  ApiError,
  GTM_API_BASE,
  ENTITY_KINDS,
  BUILT_IN_TRIGGER_IDS,
  BUILT_IN_VARIABLE_TYPES,
  normalizeKind,
};
