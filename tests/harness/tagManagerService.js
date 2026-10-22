/**
 * @fileoverview Mock of the `TagManager` advanced service (Tag Manager API v2)
 * built on top of the UrlFetchApp mock, so calls are routed to the same
 * fetchHandler (e.g. the fake GTM API) or real HTTP transport and show up in
 * runtime.fetchLog.
 *
 * It is installed only when the manifest (appsscript.json) enables the
 * `tagmanager` advanced service (or options.advancedServices says so), which
 * mirrors Apps Script: the global does not exist unless enabled.
 *
 * Errors are thrown like advanced services do:
 *   name: 'GoogleJsonResponseException'
 *   message: 'API call to tagmanager.<method> failed with error: <message>'
 *   details: {code, message, errors}
 */
'use strict';

const API_BASE = 'https://tagmanager.googleapis.com/tagmanager/v2/';

/**
 * Builds a query string; arrays become repeated keys.
 * @param {?Object} query
 * @return {string}
 */
function queryString(query) {
  const parts = [];
  for (const [k, v] of Object.entries(query || {})) {
    if (v === undefined || v === null) continue;
    for (const item of Array.isArray(v) ? v : [v]) {
      parts.push(`${encodeURIComponent(k)}=${encodeURIComponent(String(item))}`);
    }
  }
  return parts.length ? `?${parts.join('&')}` : '';
}

/**
 * Creates the TagManager advanced service object.
 * @param {!Object} env Runtime environment ({realm, getOAuthToken}).
 * @param {function(string, !Object): !Object} fetchInternal UrlFetchApp.fetch equivalent.
 * @return {!Object}
 */
function createTagManagerService(env, fetchInternal) {
  const realm = env.realm;

  /**
   * @param {string} methodId e.g. 'accounts.containers.workspaces.tags.list'
   * @param {string} httpMethod
   * @param {string} path API-relative path.
   * @param {?Object=} query
   * @param {*=} body
   * @return {*} Parsed (context-realm) response, or undefined for empty bodies.
   */
  function call(methodId, httpMethod, path, query, body) {
    if (typeof path !== 'string' || !path) {
      throw realm.error(`Invalid argument: path (tagmanager.${methodId})`);
    }
    const params = {
      method: httpMethod,
      headers: {Authorization: `Bearer ${env.getOAuthToken()}`},
      muteHttpExceptions: true,
    };
    if (body !== undefined) {
      params.contentType = 'application/json';
      params.payload = JSON.stringify(realm.toHost(body));
    }
    const res = fetchInternal(`${API_BASE}${path.replace(/^\/+/, '')}${queryString(query)}`, params);
    const code = res.getResponseCode();
    const text = res.getContentText();
    if (code >= 400) {
      let error = {};
      try {
        error = JSON.parse(text).error || {};
      } catch (e) {
        error = {};
      }
      const message = error.message || text || `HTTP ${code}`;
      throw realm.error(`API call to tagmanager.${methodId} failed with error: ${message}`,
          'GoogleJsonResponseException',
          {details: {code: error.code || code, message, errors: error.errors || []}});
    }
    if (!text || !text.trim()) return undefined;
    return realm.toContext(JSON.parse(text));
  }

  /**
   * Standard CRUD resource under a workspace/container.
   * @param {string} id Method id prefix.
   * @param {string} collection URL collection segment.
   * @return {!Object}
   */
  const crud = (id, collection) => ({
    list: (parent, opt) => call(`${id}.list`, 'GET', `${parent}/${collection}`, opt),
    get: (path, opt) => call(`${id}.get`, 'GET', path, opt),
    create: (resource, parent, opt) => call(`${id}.create`, 'POST', `${parent}/${collection}`, opt, resource),
    update: (resource, path, opt) => call(`${id}.update`, 'PUT', path, opt, resource),
    remove: (path, opt) => call(`${id}.delete`, 'DELETE', path, opt),
    revert: (path, opt) => call(`${id}.revert`, 'POST', `${path}:revert`, opt),
  });

  const ws = 'accounts.containers.workspaces';
  const Workspaces = {
    ...crud(ws, 'workspaces'),
    getStatus: (path) => call(`${ws}.getStatus`, 'GET', `${path}/status`),
    sync: (path) => call(`${ws}.sync`, 'POST', `${path}:sync`),
    create_version: (resource, path) => call(`${ws}.create_version`, 'POST', `${path}:create_version`,
        null, resource || {}),
    quick_preview: (path) => call(`${ws}.quick_preview`, 'POST', `${path}:quick_preview`),
    resolve_conflict: (resource, path, opt) => call(`${ws}.resolve_conflict`, 'POST',
        `${path}:resolve_conflict`, opt, resource),
    bulk_update: (resource, path) => call(`${ws}.bulk_update`, 'POST', `${path}/bulk_update`, null, resource),
    Tags: crud(`${ws}.tags`, 'tags'),
    Triggers: crud(`${ws}.triggers`, 'triggers'),
    Variables: crud(`${ws}.variables`, 'variables'),
    Folders: {
      ...crud(`${ws}.folders`, 'folders'),
      entities: (path, opt) => call(`${ws}.folders.entities`, 'POST', `${path}:entities`, opt),
      move_entities_to_folder: (resource, path, opt) => call(`${ws}.folders.move_entities_to_folder`,
          'POST', `${path}:move_entities_to_folder`, opt, resource || {}),
    },
    Clients: crud(`${ws}.clients`, 'clients'),
    Transformations: crud(`${ws}.transformations`, 'transformations'),
    Templates: {
      ...crud(`${ws}.templates`, 'templates'),
      import_from_gallery: (parent, opt) => call(`${ws}.templates.import_from_gallery`, 'POST',
          `${parent}/templates:import_from_gallery`, opt),
    },
    Zones: crud(`${ws}.zones`, 'zones'),
    Gtag_config: crud(`${ws}.gtag_config`, 'gtag_config'),
    Built_in_variables: {
      list: (parent, opt) => call(`${ws}.built_in_variables.list`, 'GET', `${parent}/built_in_variables`, opt),
      create: (parent, opt) => call(`${ws}.built_in_variables.create`, 'POST', `${parent}/built_in_variables`, opt),
      remove: (path, opt) => call(`${ws}.built_in_variables.delete`, 'DELETE', `${path}/built_in_variables`, opt),
      revert: (path, opt) => call(`${ws}.built_in_variables.revert`, 'POST', `${path}/built_in_variables:revert`, opt),
    },
  };

  const Containers = {
    ...crud('accounts.containers', 'containers'),
    lookup: (opt) => call('accounts.containers.lookup', 'GET', 'accounts/containers:lookup', opt),
    snippet: (path) => call('accounts.containers.snippet', 'GET', `${path}:snippet`),
    combine: (path, opt) => call('accounts.containers.combine', 'POST', `${path}:combine`, opt),
    move_tag_id: (path, opt) => call('accounts.containers.move_tag_id', 'POST', `${path}:move_tag_id`, opt),
    Workspaces,
    Destinations: {
      list: (parent) => call('accounts.containers.destinations.list', 'GET', `${parent}/destinations`),
      get: (path) => call('accounts.containers.destinations.get', 'GET', path),
      link: (parent, opt) => call('accounts.containers.destinations.link', 'POST', `${parent}/destinations:link`, opt),
    },
    Environments: crud('accounts.containers.environments', 'environments'),
    Versions: {
      get: (path, opt) => call('accounts.containers.versions.get', 'GET', path, opt),
      live: (parent) => call('accounts.containers.versions.live', 'GET', `${parent}/versions:live`),
      publish: (path, opt) => call('accounts.containers.versions.publish', 'POST', `${path}:publish`, opt),
      remove: (path) => call('accounts.containers.versions.delete', 'DELETE', path),
      update: (resource, path, opt) => call('accounts.containers.versions.update', 'PUT', path, opt, resource),
    },
    Version_headers: {
      list: (parent, opt) => call('accounts.containers.version_headers.list', 'GET', `${parent}/version_headers`, opt),
      latest: (parent) => call('accounts.containers.version_headers.latest', 'GET', `${parent}/version_headers:latest`),
    },
  };

  return {
    Accounts: {
      list: (opt) => call('accounts.list', 'GET', 'accounts', opt),
      get: (path, opt) => call('accounts.get', 'GET', path, opt),
      update: (resource, path, opt) => call('accounts.update', 'PUT', path, opt, resource),
      Containers,
      User_permissions: crud('accounts.user_permissions', 'user_permissions'),
    },
    toString: () => 'AdvancedServiceIdentifier',
  };
}

module.exports = {createTagManagerService, queryString, API_BASE};
