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
 * @fileoverview The "Clients" sheet (server containers).
 */

/** Client types that can be created in the Clients sheet. */
const CREATABLE_CLIENT_TYPES = ['gaaw_client', 'gtm_client', 'mpaw_client',
                                'ga4f_client'];

/** Cookie management options of the GA4 client. */
const COOKIE_MANAGEMENT_LABELS = {server: 'Server managed',
                                  js: 'JavaScript managed'};

/**
 * Returns the definition of the Clients sheet.
 * @return {!Object}
 */
function clientsSheetDef_() {
  const ga4 = 'GA4 (WEB) CLIENT';
  const gtm = 'GTM WEB CONTAINER CLIENT';
  const both = 'GA4 & GTM CLIENTS';
  return {
    key: 'clients',
    sheetName: 'Clients',
    title: 'CLIENTS (SERVER)',
    tabColor: TAB_COLORS.server,
    mode: 'entity',
    kind: 'client',
    contexts: ['server'],
    itemLabel: 'client',
    itemLabelPlural: 'clients',
    idHeader: 'Client ID',
    nameColumn: 'name',
    sort: (a, b) => (b.priority || 0) - (a.priority || 0) ||
        compareText(a.name, b.name),
    columns: [
      COLUMNS.name('Client', {group: 'CLIENT'}),
      COLUMNS.type(CLIENT_TYPE_NAMES, 'Client', CREATABLE_CLIENT_TYPES,
          {group: 'CLIENT'}),
      clientPriorityColumn_({group: 'CLIENT'}),
      // GA4 client.
      COLUMNS.boolParam('activateDefaultPaths', 'Default GA4 Paths', {
        group: ga4, types: ['gaaw_client'], defaultValue: true, width: 100
      }),
      COLUMNS.selectParam('cookieManagement', 'Cookie Management',
          COOKIE_MANAGEMENT_LABELS, {
            group: ga4, types: ['gaaw_client'], defaultValue: 'server',
            width: 150
          }),
      COLUMNS.param('cookieName', 'Cookie Name', {
        group: ga4, types: ['gaaw_client'], width: 110,
        note: 'Default: FPID.'
      }),
      COLUMNS.param('cookieDomain', 'Cookie Domain', {
        group: ga4, types: ['gaaw_client'], width: 120,
        note: 'Default: auto.'
      }),
      COLUMNS.param('cookiePath', 'Cookie Path', {
        group: ga4, types: ['gaaw_client'], width: 100, note: 'Default: /'
      }),
      COLUMNS.param('cookieMaxAgeInSec', 'Cookie Max Age (seconds)', {
        group: ga4, types: ['gaaw_client'], width: 120,
        note: 'Default: 63072000 (2 years).'
      }),
      COLUMNS.selectParam('cookieSamesite', 'Cookie SameSite',
          {'': 'Not set', lax: 'Lax', strict: 'Strict'}, {
            group: ga4, types: ['gaaw_client'], width: 100
          }),
      COLUMNS.boolParam('migrateFromJsClientId',
          'Migrate from JavaScript Managed Client ID', {
            group: ga4, types: ['gaaw_client'], width: 140
          }),
      // GTM web container client.
      COLUMNS.tableParam('allowedContainerIds',
          'Allowed Container IDs (ID = path)', ['containerId', 'path'], {
            group: gtm, types: ['gtm_client'], width: 260,
            note: 'Web containers that can be served by this client, one per ' +
                'line: GTM-XXXXXXX = /path (the path is optional).'
          }),
      // Settings of both clients.
      COLUMNS.boolParam('activateResponseCompression',
          'Compress HTTP Responses', {
            group: both, types: ['gaaw_client', 'gtm_client'],
            defaultValue: true, width: 110
          }),
      COLUMNS.boolParam('activateGeoResolution',
          'Enable Region-Specific Settings', {
            group: both, types: ['gaaw_client', 'gtm_client'],
            defaultValue: true, width: 120
          }),
      COLUMNS.param('region', 'Region Variable', {
        group: both, types: ['gaaw_client', 'gtm_client'], width: 140,
        dropdown: 'Variables', note: 'A "Visitor Region" variable.'
      }),
      COLUMNS.param('activationPath', 'Activation Path', {
        group: 'MEASUREMENT PROTOCOL CLIENT', types: ['mpaw_client'],
        width: 140, note: 'Default: /mp/collect'
      }),
      COLUMNS.folder({group: 'CLIENT SETTINGS'}),
      COLUMNS.notes({group: 'CLIENT SETTINGS'})
    ],
    newEntity: values => ({
      type: resolveTypeLabel(CLIENT_TYPE_NAMES, values.type),
      parameter: []
    }),
    validate: draft => {
      if (draft.type === 'gtm_client' &&
          !getTableRows(draft, 'allowedContainerIds', ['containerId']).length) {
        return ['Allowed Container IDs: add at least one web container ID.'];
      }
      return [];
    }
  };
}

/**
 * Column: priority of a client.
 * @param {!Object=} options
 * @return {!Object}
 */
function clientPriorityColumn_(options) {
  return buildColumn_({
    key: 'priority',
    header: 'Priority',
    width: 80,
    note: 'Clients with a higher priority run first. Whole number.',
    get: entity => entity.priority !== undefined ? String(entity.priority) :
        '',
    set: (draft, value) => {
      const text = toText(value);
      if (!text) {
        delete draft.priority;
      } else if (/^-?\d+$/.test(text)) {
        draft.priority = parseInt(text, 10);
      } else {
        throw new UserError('Priority must be a whole number.');
      }
    }
  }, options);
}
