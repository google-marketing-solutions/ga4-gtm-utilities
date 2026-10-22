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
 * @fileoverview The "GA4 Event Tags" sheet (web containers).
 */

/**
 * Returns the definition of the GA4 Event Tags sheet.
 * @return {!Object}
 */
function ga4EventTagsSheetDef_() {
  return {
    key: 'ga4EventTags',
    sheetName: 'GA4 Event Tags',
    title: 'GA4 EVENT TAGS (WEB)',
    tabColor: TAB_COLORS.web,
    mode: 'entity',
    kind: 'tag',
    contexts: ['web'],
    itemLabel: 'GA4 event tag',
    itemLabelPlural: 'GA4 event tags',
    idHeader: 'Tag ID',
    nameColumn: 'name',
    filter: tag => tag.type === 'gaawe',
    columns: [
      COLUMNS.name('Tag', {group: 'TAG'}),
      COLUMNS.param('eventName', 'Event Name', {
        group: 'GA4 EVENT', required: true, width: 180,
        note: 'Name of the GA4 event, e.g. purchase or {{Event}}.'
      }),
      COLUMNS.param('measurementIdOverride', 'Measurement ID', {
        group: 'GA4 EVENT', required: true, width: 200,
        dropdown: 'Variables',
        note: 'GA4 measurement ID (G-XXXXXXXXXX) or a variable, e.g. ' +
            '{{GA4 Measurement ID}}.'
      }),
      COLUMNS.tableParam('eventSettingsTable', 'Event Parameters',
          ['parameter', 'parameterValue'], {
            group: 'GA4 EVENT',
            note: 'One parameter per line: name = value, e.g.\n' +
                'page_type = {{Page Type}}\nvalue = {{DLV - value}}\n' +
                'Tip: the "Event Parameters & User Properties" sheet lists ' +
                'one parameter per row, which is easier for bulk changes.'
          }),
      COLUMNS.tableParam('userProperties', 'User Properties',
          ['name', 'value'], {
            group: 'GA4 EVENT',
            note: 'One user property per line: name = value.'
          }),
      COLUMNS.param('eventSettingsVariable', 'Event Settings Variable', {
        group: 'GA4 EVENT', dropdown: 'Variables',
        note: 'A "Google Tag: Event Settings" variable, e.g. ' +
            '{{Event Settings}}.'
      }),
      COLUMNS.boolParam('sendEcommerceData', 'Send Ecommerce Data', {
        group: 'ECOMMERCE'
      }),
      ecommerceSourceColumn_({group: 'ECOMMERCE'})
    ].concat(commonTagColumns_()),
    newEntity: () => ({type: 'gaawe', parameter: []}),
    validate: draft => {
      const errors = [];
      if (!getParamValue(draft, 'eventName')) {
        errors.push('Event Name is required.');
      }
      if (!getParamValue(draft, 'measurementIdOverride') &&
          !getParamValue(draft, 'measurementId')) {
        errors.push('Measurement ID is required.');
      }
      return errors;
    }
  };
}

/**
 * Column: where the ecommerce data of a GA4 event tag comes from.
 * @param {!Object=} options
 * @return {!Object}
 */
function ecommerceSourceColumn_(options) {
  return buildColumn_({
    key: 'ecommerceSource',
    header: 'Ecommerce Data Source',
    width: 180,
    dropdown: ['dataLayer'],
    note: '"dataLayer" or a variable that returns the ecommerce object, ' +
        'e.g. {{Ecommerce Object}}. Only used when "Send Ecommerce Data" is ' +
        'ticked.',
    get: entity => {
      if (!getBoolParam(entity, 'sendEcommerceData')) return '';
      return getParamValue(entity, 'getEcommerceDataFrom') === 'customObject' ?
          getParamValue(entity, 'ecommerceMacroData') : 'dataLayer';
    },
    set: (draft, value) => {
      const text = toText(value);
      if (!text || text.toLowerCase() === 'datalayer') {
        setParamValue(draft, 'getEcommerceDataFrom', 'dataLayer');
        removeParam(draft, 'ecommerceMacroData');
      } else {
        setParamValue(draft, 'getEcommerceDataFrom', 'customObject');
        setParamValue(draft, 'ecommerceMacroData', text);
      }
    }
  }, options);
}
