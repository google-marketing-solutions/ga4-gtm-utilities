/**
 * Copyright 2026 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * @fileoverview Client for the Google Tag Manager REST API v2.
 *
 * The API is called with UrlFetchApp (instead of the TagManager advanced
 * service) so that every API method is available (e.g. bulk_update and
 * gtag_config), errors can be reported clearly and rate limits can be handled
 * with retries. The TagManager advanced service stays enabled in the manifest
 * because that enables the API in the script's Google Cloud project.
 */

const GTM_API_BASE_URL = 'https://tagmanager.googleapis.com/tagmanager/v2/';

/**
 * Workspace entity kinds: REST collection, ID field, list response field and
 * workspaces.bulk_update field.
 */
const GTM_ENTITY_KINDS = {
  tag: {collection: 'tags', idField: 'tagId', listKey: 'tag',
        bulkKey: 'tag', label: 'Tag'},
  trigger: {collection: 'triggers', idField: 'triggerId', listKey: 'trigger',
            bulkKey: 'trigger', label: 'Trigger'},
  variable: {collection: 'variables', idField: 'variableId',
             listKey: 'variable', bulkKey: 'variable', label: 'Variable'},
  folder: {collection: 'folders', idField: 'folderId', listKey: 'folder',
           bulkKey: 'folder', label: 'Folder'},
  client: {collection: 'clients', idField: 'clientId', listKey: 'client',
           bulkKey: 'client', label: 'Client'},
  transformation: {collection: 'transformations',
                   idField: 'transformationId', listKey: 'transformation',
                   bulkKey: 'transformation', label: 'Transformation'},
  template: {collection: 'templates', idField: 'templateId',
             listKey: 'template', bulkKey: 'customTemplate',
             label: 'Template'},
  gtagConfig: {collection: 'gtag_config', idField: 'gtagConfigId',
               listKey: 'gtagConfig', bulkKey: 'gtagConfig',
               label: 'Google tag setting'},
  zone: {collection: 'zones', idField: 'zoneId', listKey: 'zone',
         bulkKey: 'zone', label: 'Zone'},
  builtInVariable: {collection: 'built_in_variables', idField: 'type',
                    listKey: 'builtInVariable', bulkKey: 'builtInVariable',
                    label: 'Built-in variable'}
};

/** Time when the current script execution started. */
const EXECUTION_STATE = {startTime: Date.now()};

/**
 * Restarts the execution timer. Called at the beginning of every menu action.
 */
function startExecutionTimer() {
  EXECUTION_STATE.startTime = Date.now();
}

/**
 * Returns the remaining time of the execution budget in milliseconds.
 * @return {number}
 */
function getRemainingExecutionMs() {
  return LIMITS.executionBudgetMs - (Date.now() - EXECUTION_STATE.startTime);
}

/**
 * Error returned by the GTM API.
 */
class GtmApiError extends Error {
  /**
   * @param {string} message
   * @param {number} code HTTP status code.
   * @param {string=} status API status, e.g. INVALID_ARGUMENT.
   */
  constructor(message, code, status) {
    super(message);
    this.name = 'GtmApiError';
    this.code = code;
    this.status = status || '';
  }
}

/**
 * Waits if needed so that the number of API requests stays within the
 * per-project rate limit. Request times are shared between executions with
 * the user cache.
 */
function waitForApiRateLimit_() {
  const cache = CacheService.getUserCache();
  let times = [];
  try {
    times = JSON.parse(cache.get(PROPERTY_KEYS.apiCallTimes) || '[]');
  } catch (e) {
    times = [];
  }
  let now = Date.now();
  times = times.filter(time => now - time < LIMITS.apiWindowMs);
  if (times.length >= LIMITS.apiMaxRequestsPerWindow) {
    const waitMs = LIMITS.apiWindowMs - (now - times[0]) + 200;
    if (waitMs > getRemainingExecutionMs()) {
      throw new UserError('The Tag Manager API rate limit was reached and ' +
          'there is not enough time left in this run. Please run the action ' +
          'again in a minute.');
    }
    if (waitMs > 3000) {
      toast_(`Waiting ${Math.ceil(waitMs / 1000)} seconds for the Tag ` +
          'Manager API rate limit…');
    }
    Utilities.sleep(waitMs);
    now = Date.now();
    times = times.filter(time => now - time < LIMITS.apiWindowMs);
  }
  times.push(now);
  cache.put(PROPERTY_KEYS.apiCallTimes, JSON.stringify(times), 600);
}

/**
 * Shows a toast in the active spreadsheet, ignoring errors (e.g. when no UI
 * is available).
 * @param {string} message
 * @param {number=} seconds
 */
function toast_(message, seconds) {
  try {
    SpreadsheetApp.getActive().toast(message, APP_NAME, seconds || 5);
  } catch (e) {
    console.log(message);
  }
}

/**
 * Builds a query string.
 * @param {!Object<string, *>|undefined} query Values can be arrays (the key is
 *     repeated).
 * @return {string}
 */
function buildQueryString_(query) {
  if (!query) return '';
  const parts = [];
  Object.keys(query).forEach(key => {
    const value = query[key];
    if (value === undefined || value === null || value === '') return;
    (Array.isArray(value) ? value : [value]).forEach(item => {
      parts.push(encodeURIComponent(key) + '=' +
          encodeURIComponent(String(item)));
    });
  });
  return parts.length ? '?' + parts.join('&') : '';
}

/**
 * Returns true if a failed request can be retried.
 * @param {number} code
 * @param {!Object} error The error object of the response.
 * @return {boolean}
 */
function isRetryableApiError_(code, error) {
  if (code === 429 || code === 500 || code === 502 || code === 503 ||
      code === 504) {
    return true;
  }
  if (code === 403) {
    const reasons = (error.errors || []).map(e => e.reason);
    return error.status === 'RESOURCE_EXHAUSTED' ||
        reasons.some(r => /rateLimitExceeded|userRateLimitExceeded/.test(r));
  }
  return false;
}

/**
 * Makes a request to the GTM API.
 * @param {string} method HTTP method.
 * @param {string} path Path relative to the API base URL, e.g.
 *     "accounts/1/containers/2".
 * @param {{query: (!Object|undefined), body: (!Object|undefined)}=} options
 * @return {!Object} The parsed JSON response ({} for empty responses).
 */
function gtmApiRequest(method, path, options) {
  const opts = options || {};
  const url = GTM_API_BASE_URL + path + buildQueryString_(opts.query);
  const params = {
    method: method.toLowerCase(),
    headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()},
    muteHttpExceptions: true
  };
  if (opts.body !== undefined) {
    params.contentType = 'application/json';
    params.payload = JSON.stringify(opts.body);
  }
  let attempt = 0;
  while (true) {
    waitForApiRateLimit_();
    const response = UrlFetchApp.fetch(url, params);
    const code = response.getResponseCode();
    const text = response.getContentText();
    if (code >= 200 && code < 300) {
      if (!text) return {};
      try {
        return JSON.parse(text);
      } catch (e) {
        return {};
      }
    }
    let error = {};
    try {
      error = JSON.parse(text).error || {};
    } catch (e) {
      error = {message: text};
    }
    attempt++;
    const backoffMs = Math.min(60000, 5000 * Math.pow(2, attempt - 1));
    if (isRetryableApiError_(code, error) && attempt <= LIMITS.apiMaxRetries &&
        backoffMs + 5000 < getRemainingExecutionMs()) {
      console.log(`GTM API ${code} for ${method} ${path}; retrying in ` +
          `${backoffMs} ms`);
      Utilities.sleep(backoffMs);
      continue;
    }
    throw new GtmApiError(formatApiErrorMessage_(code, error), code,
        error.status);
  }
}

/**
 * Builds a readable message for an API error.
 * @param {number} code
 * @param {!Object} error
 * @return {string}
 */
function formatApiErrorMessage_(code, error) {
  const message = String(error.message || 'Unknown error').trim();
  if (code === 401) {
    return 'Not authorized to use the Tag Manager API. Reload the ' +
        'spreadsheet and authorize the script again.';
  }
  if (code === 403 && /has not been used|is disabled/i.test(message)) {
    return 'The Tag Manager API is not enabled for this script. ' + message;
  }
  if (code === 403) {
    return 'Permission denied: ' + message + ' (check that you have ' +
        'Edit access to the container).';
  }
  if (code === 404) return 'Not found: ' + message;
  if (code === 429 || error.status === 'RESOURCE_EXHAUSTED') {
    return 'Tag Manager API quota exceeded: ' + message +
        ' Please try again in a few minutes.';
  }
  return message.replace(/^vendorTemplate\.parameter\./gm, 'Setting ')
      .replace(/\n+/g, ' ');
}

/**
 * Lists all items of a collection, following the pagination.
 * @param {string} path Collection path, e.g. ".../workspaces/3/tags".
 * @param {string} listKey Field of the list response that holds the items.
 * @param {!Object=} query Additional query parameters.
 * @return {!Array<!Object>}
 */
function gtmApiListAll(path, listKey, query) {
  let items = [];
  let pageToken;
  do {
    const pageQuery = Object.assign({}, query || {});
    if (pageToken) pageQuery.pageToken = pageToken;
    const response = gtmApiRequest('get', path, {query: pageQuery});
    items = items.concat(response[listKey] || []);
    pageToken = response.nextPageToken;
  } while (pageToken);
  return items;
}

/**
 * Lists all entities of a kind in a workspace.
 * @param {string} workspacePath
 * @param {string} kind A key of GTM_ENTITY_KINDS.
 * @return {!Array<!Object>}
 */
function gtmListEntities(workspacePath, kind) {
  const info = GTM_ENTITY_KINDS[kind];
  return gtmApiListAll(`${workspacePath}/${info.collection}`, info.listKey);
}

/**
 * Creates an entity in a workspace.
 * @param {string} workspacePath
 * @param {string} kind
 * @param {!Object} entity
 * @return {!Object} The created entity.
 */
function gtmCreateEntity(workspacePath, kind, entity) {
  const info = GTM_ENTITY_KINDS[kind];
  return gtmApiRequest('post', `${workspacePath}/${info.collection}`,
      {body: entity});
}

/**
 * Updates an entity.
 * @param {string} path The entity path.
 * @param {!Object} entity
 * @return {!Object} The updated entity.
 */
function gtmUpdateEntity(path, entity) {
  return gtmApiRequest('put', path, {body: entity});
}

/**
 * Deletes an entity.
 * @param {string} path The entity path.
 */
function gtmDeleteEntity(path) {
  gtmApiRequest('delete', path);
}

/**
 * Applies multiple entity changes to a workspace in one call. The call is
 * atomic: either all or none of the changes are applied.
 * @param {string} workspacePath
 * @param {!Array<!Object>} changes Entity changes (see the API reference).
 * @return {!Array<!Object>} The added and updated entities.
 */
function gtmBulkUpdate(workspacePath, changes) {
  const response = gtmApiRequest('post', `${workspacePath}/bulk_update`,
      {body: {changes: changes}});
  return response.changes || [];
}

/**
 * Returns the entity of a bulk_update response change.
 * @param {!Object} change
 * @return {{kind: string, entity: !Object}|null}
 */
function getBulkChangeEntity(change) {
  for (const kind of Object.keys(GTM_ENTITY_KINDS)) {
    const key = GTM_ENTITY_KINDS[kind].bulkKey;
    if (change[key]) return {kind: kind, entity: change[key]};
  }
  return null;
}
