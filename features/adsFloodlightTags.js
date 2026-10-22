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
 * @fileoverview The "Google Ads & Floodlight Tags" sheet (web containers).
 */

/** Tag types managed in the Google Ads & Floodlight Tags sheet. */
const ADS_FLOODLIGHT_TAG_TYPES = ['awct', 'sp', 'awud', 'awcc', 'gclidw', 'flc',
                                  'fls'];

/** Floodlight counting methods. */
const FLOODLIGHT_COUNTING_LABELS = {
  STANDARD: 'Standard',
  UNIQUE: 'Unique',
  SESSION: 'Per Session',
  PER_SESSION: 'Per Session',
  TRANSACTIONS: 'Transactions',
  ITEM_SOLD: 'Items Sold'
};

/**
 * Returns the definition of the Google Ads & Floodlight Tags sheet.
 * @return {!Object}
 */
function adsFloodlightTagsSheetDef_() {
  const ads = 'GOOGLE ADS';
  const linker = 'CONVERSION LINKER';
  const floodlight = 'FLOODLIGHT';
  return {
    key: 'adsFloodlightTags',
    sheetName: 'Google Ads & Floodlight Tags',
    title: 'GOOGLE ADS & FLOODLIGHT TAGS (WEB)',
    tabColor: TAB_COLORS.web,
    mode: 'entity',
    kind: 'tag',
    contexts: ['web'],
    itemLabel: 'tag',
    itemLabelPlural: 'Google Ads and Floodlight tags',
    idHeader: 'Tag ID',
    nameColumn: 'name',
    filter: tag => ADS_FLOODLIGHT_TAG_TYPES.indexOf(tag.type) !== -1,
    columns: [
      COLUMNS.name('Tag', {group: 'TAG'}),
      COLUMNS.type(TAG_TYPE_NAMES, 'Tag', ADS_FLOODLIGHT_TAG_TYPES,
          {group: 'TAG'}),
      // Google Ads.
      COLUMNS.param('conversionId', 'Conversion ID', {
        group: ads, types: ['awct', 'sp', 'awud', 'awcc'], width: 130,
        note: 'Google Ads conversion ID (numbers only, without "AW-").'
      }),
      COLUMNS.param('conversionLabel', 'Conversion Label', {
        group: ads, types: ['awct', 'sp', 'awcc'], width: 150
      }),
      COLUMNS.param('conversionValue', 'Conversion Value', {
        group: ads, types: ['awct'], width: 130, dropdown: 'Variables'
      }),
      COLUMNS.param('orderId', 'Transaction ID', {
        key: 'ads:orderId', group: ads, types: ['awct'], width: 140,
        dropdown: 'Variables'
      }),
      COLUMNS.param('currencyCode', 'Currency Code', {
        group: ads, types: ['awct'], width: 100
      }),
      COLUMNS.param('phoneConversionNumber', 'Phone Number to Replace', {
        group: ads, types: ['awcc'], width: 150
      }),
      COLUMNS.param('userDataVariable', 'User-Provided Data Variable', {
        group: ads, types: ['awud'], width: 190, dropdown: 'Variables',
        note: 'A "User-Provided Data" variable.'
      }),
      COLUMNS.boolParam('enableProductReporting', 'Product-Level Sales Data', {
        group: ads, types: ['awct'], width: 110
      }),
      COLUMNS.param('awMerchantId', 'Merchant Center ID', {
        group: ads, types: ['awct'], width: 140
      }),
      COLUMNS.param('awFeedCountry', 'Feed Country', {
        group: ads, types: ['awct'], width: 100
      }),
      COLUMNS.param('awFeedLanguage', 'Feed Language', {
        group: ads, types: ['awct'], width: 100
      }),
      remarketingCustomParamsColumn_({group: ads, types: ['sp']}),
      COLUMNS.param('userId', 'User ID', {
        group: ads, types: ['sp'], width: 130, dropdown: 'Variables'
      }),
      // Conversion Linker.
      COLUMNS.boolParam('enableUrlPassthrough',
          'Enable Linking on All Page URLs', {
            group: linker, types: ['gclidw'], width: 130
          }),
      conditionalParamColumn_('enableCrossDomain', 'linkerDomains',
          'Auto-Link Domains (Cross-Domain)', {
            group: linker, types: ['gclidw'], width: 200,
            note: 'Comma-separated domains for cross-domain linking. ' +
                'Leave empty to disable linking across domains.'
          }),
      conditionalParamColumn_('enableCookieOverrides', 'cookiePrefix',
          'Cookie Prefix (Override)', {
            group: linker, types: ['gclidw'], width: 150,
            note: 'Name prefix of the conversion linker cookies. Leave empty ' +
                'to use the default cookie settings.'
          }),
      // Floodlight.
      COLUMNS.param('advertiserId', 'Advertiser ID', {
        group: floodlight, types: ['flc', 'fls'], width: 120
      }),
      COLUMNS.param('groupTag', 'Group Tag String (type=)', {
        group: floodlight, types: ['flc', 'fls'], width: 150
      }),
      COLUMNS.param('activityTag', 'Activity Tag String (cat=)', {
        group: floodlight, types: ['flc', 'fls'], width: 150
      }),
      floodlightCountingColumn_({group: floodlight, types: ['flc', 'fls']}),
      COLUMNS.param('revenue', 'Revenue', {
        group: floodlight, types: ['flc', 'fls'], width: 120,
        dropdown: 'Variables'
      }),
      COLUMNS.param('orderId', 'Order ID', {
        key: 'fl:orderId', group: floodlight, types: ['flc', 'fls'],
        width: 120, dropdown: 'Variables'
      }),
      COLUMNS.param('quantity', 'Quantity', {
        group: floodlight, types: ['flc', 'fls'], width: 100,
        dropdown: 'Variables'
      }),
      COLUMNS.tableParam('customVariable', 'Custom Variables',
          ['key', 'value'], {
            group: floodlight, types: ['flc', 'fls'],
            note: 'One custom variable per line, e.g.\nu1 = {{Product ID}}\n' +
                'u2 = {{Category}}'
          })
    ].concat(commonTagColumns_()),
    newEntity: values => {
      const type = resolveTypeLabel(TAG_TYPE_NAMES, values.type);
      const entity = {type: type, parameter: []};
      if (type === 'flc') setParamValue(entity, 'ordinalType', 'STANDARD');
      if (type === 'fls') setParamValue(entity, 'countingMethod', 'TRANSACTIONS');
      return entity;
    },
    validate: draft => validateAdsFloodlightTag_(draft)
  };
}

/**
 * Validates the required settings of Google Ads and Floodlight tags.
 * @param {!Object} draft
 * @return {!Array<string>} Errors.
 */
function validateAdsFloodlightTag_(draft) {
  const required = {
    awct: [['conversionId', 'Conversion ID'],
           ['conversionLabel', 'Conversion Label']],
    sp: [['conversionId', 'Conversion ID']],
    awud: [['conversionId', 'Conversion ID'],
           ['userDataVariable', 'User-Provided Data Variable']],
    awcc: [['conversionId', 'Conversion ID'],
           ['conversionLabel', 'Conversion Label']],
    flc: [['advertiserId', 'Advertiser ID'], ['groupTag', 'Group Tag String'],
          ['activityTag', 'Activity Tag String']],
    fls: [['advertiserId', 'Advertiser ID'], ['groupTag', 'Group Tag String'],
          ['activityTag', 'Activity Tag String']],
    sgtmadsct: [['conversionId', 'Conversion ID'],
                ['conversionLabel', 'Conversion Label']],
    sgtmadsremarket: [['conversionId', 'Conversion ID']],
    sgtmawud: [['conversionId', 'Conversion ID']],
    sgtmflc: [['advertiserId', 'Advertiser ID'],
              ['groupTag', 'Group Tag String'],
              ['activityTag', 'Activity Tag String']],
    sgtmfls: [['advertiserId', 'Advertiser ID'],
              ['groupTag', 'Group Tag String'],
              ['activityTag', 'Activity Tag String']]
  };
  const missing = (required[draft.type] || [])
      .filter(item => !getParamValue(draft, item[0]))
      .map(item => item[1]);
  return missing.length ? [`Required: ${missing.join(', ')}.`] : [];
}

/**
 * Column: a text parameter that is only active when a checkbox parameter is
 * enabled. An empty cell disables the checkbox parameter.
 * @param {string} enableKey The boolean parameter.
 * @param {string} valueKey The text parameter.
 * @param {string} header
 * @param {!Object=} options
 * @return {!Object}
 */
function conditionalParamColumn_(enableKey, valueKey, header, options) {
  return buildColumn_({
    key: 'p:' + valueKey,
    header: header,
    width: 180,
    get: entity => getBoolParam(entity, enableKey) ?
        getParamValue(entity, valueKey) : '',
    set: (draft, value) => {
      const text = toText(value);
      setBoolParam(draft, enableKey, Boolean(text));
      if (text) {
        setParamValue(draft, valueKey, text);
      } else {
        removeParam(draft, valueKey);
      }
    }
  }, options);
}

/**
 * Column: counting method of Floodlight tags (ordinalType for counter and
 * server tags, countingMethod for web sales tags).
 * @param {!Object=} options
 * @return {!Object}
 */
function floodlightCountingColumn_(options) {
  const labels = FLOODLIGHT_COUNTING_LABELS;
  const paramKey = entity => entity.type === 'fls' ? 'countingMethod' :
      'ordinalType';
  return buildColumn_({
    key: 'floodlightCounting',
    header: 'Counting Method',
    width: 130,
    dropdown: ['Standard', 'Unique', 'Per Session', 'Transactions',
               'Items Sold'],
    note: 'Counter tags: Standard, Unique, Per Session, Transactions or ' +
        'Items Sold. Sales tags: Transactions or Items Sold.',
    get: entity => {
      const value = getParamValue(entity, paramKey(entity));
      return labels[value] || value;
    },
    set: (draft, value) => {
      const text = toText(value).toLowerCase();
      let apiValue = Object.keys(labels).find(key =>
        labels[key].toLowerCase() === text || key.toLowerCase() === text);
      if (!apiValue) {
        throw new UserError(`Unknown counting method "${toText(value)}".`);
      }
      if (apiValue === 'SESSION' || apiValue === 'PER_SESSION') {
        apiValue = /^sgtm/.test(draft.type) ? 'PER_SESSION' : 'SESSION';
      }
      setParamValue(draft, paramKey(draft), apiValue);
    }
  }, options);
}

/**
 * Column: custom parameters of a Google Ads remarketing tag. Setting custom
 * parameters also selects "Manually specify" as their source.
 * @param {!Object=} options
 * @return {!Object}
 */
function remarketingCustomParamsColumn_(options) {
  const isServer = entity => /^sgtm/.test(entity.type);
  return buildColumn_({
    key: 'p:customParams',
    header: 'Custom Parameters',
    width: 240,
    kind: 'table',
    note: 'One custom parameter per line: name = value.',
    get: entity => formatTable(getTableRows(entity, 'customParams',
        ['key', 'value'])),
    set: (draft, value) => {
      const rows = parseTable(value);
      setTableRows(draft, 'customParams', ['key', 'value'], rows);
      if (isServer(draft)) {
        setBoolParam(draft, 'enableCustomParams', rows.length > 0);
        if (rows.length) {
          setParamValue(draft, 'customParamsFormat', 'USER_SPECIFIED');
        }
      } else {
        setParamValue(draft, 'customParamsFormat',
            rows.length ? 'USER_SPECIFIED' : 'NONE');
      }
    }
  }, options);
}
