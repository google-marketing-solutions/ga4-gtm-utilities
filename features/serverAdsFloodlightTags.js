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
 * @fileoverview The "Server Ads & Floodlight Tags" sheet (server containers).
 */

/** Tag types managed in the Server Ads & Floodlight Tags sheet. */
const SERVER_ADS_FLOODLIGHT_TAG_TYPES = ['sgtmadsct', 'sgtmadsremarket',
                                         'sgtmawud', 'sgtmadscl', 'sgtmflc',
                                         'sgtmfls'];

/** SameSite values of the server Conversion Linker cookies. */
const SAME_SITE_LABELS = {unset: 'Not set', lax: 'Lax', strict: 'Strict'};

/**
 * Returns the definition of the Server Ads & Floodlight Tags sheet.
 * @return {!Object}
 */
function serverAdsFloodlightTagsSheetDef_() {
  const ads = 'GOOGLE ADS';
  const linker = 'CONVERSION LINKER';
  const floodlight = 'FLOODLIGHT';
  const floodlightTypes = ['sgtmflc', 'sgtmfls'];
  return {
    key: 'serverAdsFloodlightTags',
    sheetName: 'Server Ads & Floodlight Tags',
    title: 'GOOGLE ADS & FLOODLIGHT TAGS (SERVER)',
    tabColor: TAB_COLORS.server,
    mode: 'entity',
    kind: 'tag',
    contexts: ['server'],
    itemLabel: 'tag',
    itemLabelPlural: 'Google Ads and Floodlight tags',
    idHeader: 'Tag ID',
    nameColumn: 'name',
    filter: tag => SERVER_ADS_FLOODLIGHT_TAG_TYPES.indexOf(tag.type) !== -1,
    columns: [
      COLUMNS.name('Tag', {group: 'TAG'}),
      COLUMNS.type(TAG_TYPE_NAMES, 'Tag', SERVER_ADS_FLOODLIGHT_TAG_TYPES,
          {group: 'TAG'}),
      // Google Ads.
      COLUMNS.param('conversionId', 'Conversion ID', {
        group: ads, types: ['sgtmadsct', 'sgtmadsremarket', 'sgtmawud'],
        width: 130
      }),
      COLUMNS.param('conversionLabel', 'Conversion Label', {
        group: ads, types: ['sgtmadsct', 'sgtmadsremarket'], width: 150
      }),
      COLUMNS.param('conversionValue', 'Conversion Value', {
        group: ads, types: ['sgtmadsct'], width: 130, dropdown: 'Variables',
        note: 'Leave empty to use the value of the incoming event.'
      }),
      COLUMNS.param('currencyCode', 'Currency Code', {
        group: ads, types: ['sgtmadsct'], width: 100,
        note: 'Leave empty to use the currency of the incoming event.'
      }),
      COLUMNS.boolParam('enableProductReporting', 'Product-Level Sales Data', {
        group: ads, types: ['sgtmadsct'], width: 110
      }),
      remarketingCustomParamsColumn_({group: ads, types: ['sgtmadsremarket']}),
      // Conversion Linker.
      COLUMNS.boolParam('enableLinkerParams',
          'Accept Incoming Linker Parameters', {
            group: linker, types: ['sgtmadscl'], width: 130
          }),
      COLUMNS.boolParam('enableCookieOverrides', 'Override Cookie Settings', {
        group: linker, types: ['sgtmadscl'], width: 110
      }),
      COLUMNS.param('cookieDomain', 'Cookie Domain', {
        group: linker, types: ['sgtmadscl'], width: 120
      }),
      COLUMNS.param('cookiePath', 'Cookie Path', {
        group: linker, types: ['sgtmadscl'], width: 100
      }),
      COLUMNS.selectParam('cookieSamesite', 'Cookie SameSite',
          SAME_SITE_LABELS, {
            group: linker, types: ['sgtmadscl'], width: 100
          }),
      // Floodlight.
      COLUMNS.param('advertiserId', 'Advertiser ID', {
        group: floodlight, types: floodlightTypes, width: 120
      }),
      COLUMNS.param('groupTag', 'Group Tag String (type=)', {
        group: floodlight, types: floodlightTypes, width: 150
      }),
      COLUMNS.param('activityTag', 'Activity Tag String (cat=)', {
        group: floodlight, types: floodlightTypes, width: 150
      }),
      floodlightCountingColumn_({group: floodlight, types: floodlightTypes}),
      COLUMNS.tableParam('customVariable', 'Custom Variables',
          ['key', 'value'], {
            group: floodlight, types: floodlightTypes,
            note: 'One custom variable per line, e.g. u1 = {{Product ID}}'
          })
    ].concat(commonTagColumns_()),
    newEntity: values => {
      const type = resolveTypeLabel(TAG_TYPE_NAMES, values.type);
      const entity = {type: type, parameter: []};
      if (type === 'sgtmflc') setParamValue(entity, 'ordinalType', 'STANDARD');
      if (type === 'sgtmfls') {
        setParamValue(entity, 'ordinalType', 'TRANSACTIONS');
      }
      return entity;
    },
    validate: draft => validateAdsFloodlightTag_(draft)
  };
}
