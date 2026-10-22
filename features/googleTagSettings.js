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
 * @fileoverview The "Google Tag Settings" sheet: container-wide Google tag
 * settings of the unified Google tag / Tag Manager experience (the
 * "Settings" tab of Tag Manager, e.g. automatic event detection, domains for
 * cross-domain measurement, user-provided data collection, internal traffic,
 * unwanted referrals, session timeout and cookie settings). The linked
 * Google destinations are shown in the title row.
 */

/** Google tag setting types that are managed by the Google products. */
const READ_ONLY_GTAG_CONFIG_TYPES = ['ogt_google_signals', 'ogt_ads_datatos',
                                     'ogt_ga_datatos', 'ogt_ga_gam_link'];

/** Google tag setting types that can be created in the sheet. */
const CREATABLE_GTAG_CONFIG_TYPES = ['ogt_auto_events', 'ogt_cross_domain',
                                     'ogt_1p_data_v2', 'ogt_ip_mark',
                                     'ogt_referral_exclusion',
                                     'ogt_session_timeout',
                                     'ogt_cookie_settings', 'ogt_dtc',
                                     'ogt_consent_defaults'];

/**
 * Returns the definition of the Google Tag Settings sheet.
 * @return {!Object}
 */
function googleTagSettingsSheetDef_() {
  return {
    key: 'googleTagSettings',
    sheetName: 'Google Tag Settings',
    title: 'GOOGLE TAG SETTINGS (CONTAINER-WIDE)',
    tabColor: TAB_COLORS.web,
    mode: 'item',
    kinds: ['gtagConfig'],
    contexts: ['web', 'googleTag'],
    itemLabel: 'setting',
    itemLabelPlural: 'Google tag settings',
    idHeader: 'Setting ID',
    nameColumn: 'configType',
    columns: [
      {key: 'configType', header: 'Setting', group: 'GOOGLE TAG SETTING',
       width: 240, createOnly: true,
       dropdown: CREATABLE_GTAG_CONFIG_TYPES.map(
           type => GTAG_CONFIG_TYPE_NAMES[type]),
       note: 'Group of Google tag settings, as shown in the Settings tab of ' +
           'Tag Manager. To add a setting group that does not exist yet, ' +
           'add a row with Action "Create".'},
      {key: 'configTypeCode', header: 'Setting Type', group: 'GOOGLE TAG SETTING',
       width: 160, readOnly: true,
       note: 'Type of the setting group in the Tag Manager API.'},
      {key: 'param', header: 'Parameter Key', group: 'PARAMETER', width: 220,
       note: 'Key of the setting, e.g. enablePageView, isEnabled, rules.'},
      {key: 'value', header: 'Value', group: 'PARAMETER', width: 360,
       kind: 'multiline',
       note: 'true / false for switches; lists and tables are shown as JSON, ' +
           'e.g. [{"matchType":"CONTAINS","matchValue":"example.com"}].'},
      {key: 'valueType', header: 'Value Type', group: 'PARAMETER', width: 90,
       readOnly: true}
    ],
    isEligible: () => true,
    explode: (entity, kind) => {
      const base = {
        configType: getTypeLabel(GTAG_CONFIG_TYPE_NAMES, entity.type),
        configTypeCode: entity.type
      };
      const rows = explodeParamItems_(entity, kind, base);
      if (!rows.length) {
        rows.push({values: Object.assign({}, base, {param: '', value: '',
                                                    valueType: ''}),
                   meta: {k: kind, placeholder: true}});
      }
      return rows;
    },
    findEntity: (row, data) => findGtagConfig_(row, data),
    updateItem: (draft, row) => {
      const meta = row.sync.m || {};
      return meta.placeholder ? createParamItem_(draft, row, 'gtagConfig') :
          updateParamItem_(draft, row);
    },
    createItem: (draft, row) => createParamItem_(draft, row, 'gtagConfig'),
    deleteItem: (draft, row) => deleteParamItem_(draft, row),
    validateEntity: draft => READ_ONLY_GTAG_CONFIG_TYPES.indexOf(draft.type) !==
        -1 ? ['This setting is managed in the Google product (e.g. Google ' +
              'Analytics) and can\'t be changed here.'] : [],
    listStatusExtra: (data, ctx) => {
      if (ctx.isGoogleTag) return '';
      try {
        const destinations = gtmApiListAll(`${ctx.containerPath}/destinations`,
            'destination');
        if (!destinations.length) return ' · No linked destinations';
        return ' · Destinations: ' + destinations.map(d =>
          d.destinationId + (d.name ? ` (${d.name})` : '')).join(', ');
      } catch (e) {
        return '';
      }
    }
  };
}

/**
 * Finds the Google tag setting group of a row. A new group is prepared when
 * a group of the chosen type does not exist yet.
 * @param {!Object} row
 * @param {!WorkspaceData} data
 * @return {{kind: string, entity: ?Object, newEntity: (!Object|undefined),
 *     newKey: (string|undefined)}}
 */
function findGtagConfig_(row, data) {
  if (row.id) {
    const entity = data.getById('gtagConfig', row.id);
    if (!entity) {
      throw new UserError(`The setting with ID ${row.id} was not found. ` +
          'List the sheet again.');
    }
    return {kind: 'gtagConfig', entity: entity};
  }
  const type = resolveTypeLabel(GTAG_CONFIG_TYPE_NAMES,
      row.valuesByKey.configType);
  if (!type) throw new UserError('Choose a Setting.');
  const existing = data.list('gtagConfig').filter(c => c.type === type);
  if (existing.length === 1) return {kind: 'gtagConfig', entity: existing[0]};
  if (existing.length > 1) {
    throw new UserError('There are several settings of this type. Copy the ' +
        'Setting ID of the one you want to change to this row.');
  }
  return {kind: 'gtagConfig', entity: null, newKey: type,
          newEntity: {type: type, parameter: []}};
}
