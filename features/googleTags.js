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
 * @fileoverview The "Google Tags" sheet (web containers): the Google tag and
 * the Google destination tags of the unified Google tag / GTM experience
 * (one tag per Google Analytics, Google Ads, Floodlight or Merchant Center
 * destination).
 */

/** Tag types managed in the Google Tags sheet. */
const GOOGLE_TAG_SHEET_TYPES = ['googtag', 'dest_ga', 'dest_aw', 'dest_dc',
                                'dest_mc'];

/** Expected ID prefixes of the destination tags. */
const DESTINATION_ID_PATTERNS = {
  dest_ga: /^G-[A-Z0-9]+$/i,
  dest_aw: /^AW-\d+$/i,
  dest_dc: /^DC-\d+$/i,
  dest_mc: /^MC-[A-Z0-9]+$/i
};

/**
 * Returns the definition of the Google Tags sheet.
 * @return {!Object}
 */
function googleTagsSheetDef_() {
  const destinationTypes = ['dest_ga', 'dest_aw', 'dest_dc', 'dest_mc'];
  return {
    key: 'googleTags',
    sheetName: 'Google Tags',
    title: 'GOOGLE TAG & DESTINATION TAGS (WEB)',
    tabColor: TAB_COLORS.web,
    mode: 'entity',
    kind: 'tag',
    contexts: ['web'],
    itemLabel: 'Google tag',
    itemLabelPlural: 'Google tags and destination tags',
    idHeader: 'Tag ID',
    nameColumn: 'name',
    filter: tag => GOOGLE_TAG_SHEET_TYPES.indexOf(tag.type) !== -1,
    columns: [
      COLUMNS.name('Tag', {group: 'TAG'}),
      COLUMNS.type(TAG_TYPE_NAMES, 'Tag', GOOGLE_TAG_SHEET_TYPES, {
        group: 'TAG',
        note: 'Google Tag, or a destination tag (Google Analytics, Google ' +
            'Ads, Floodlight or Merchant Center destination).'
      }),
      googleTagIdColumn_({group: 'GOOGLE TAG SETTINGS'}),
      COLUMNS.param('configSettingsVariable',
          'Configuration Settings Variable', {
            group: 'GOOGLE TAG SETTINGS', dropdown: 'Variables',
            note: 'A "Google Tag: Configuration Settings" variable.'
          }),
      COLUMNS.tableParam('configSettingsTable', 'Configuration Parameters',
          ['parameter', 'parameterValue'], {
            group: 'GOOGLE TAG SETTINGS',
            note: 'One parameter per line: name = value, e.g.\n' +
                'send_page_view = false\n' +
                'server_container_url = https://sgtm.example.com'
          }),
      COLUMNS.param('eventSettingsVariable', 'Shared Event Settings Variable', {
        group: 'GOOGLE TAG SETTINGS', dropdown: 'Variables',
        note: 'A "Google Tag: Event Settings" variable.'
      }),
      COLUMNS.tableParam('eventSettingsTable', 'Shared Event Parameters',
          ['parameter', 'parameterValue'], {
            group: 'GOOGLE TAG SETTINGS',
            note: 'Event parameters sent with every event, one per line: ' +
                'name = value, e.g. user_data = {{User-Provided Data}}'
          }),
      COLUMNS.boolParam('ignoreDupeGtagConfigCalls',
          'Ignore Duplicate On-Page Configuration', {
            group: 'GOOGLE TAG SETTINGS', types: destinationTypes, width: 140
          })
    ].concat(commonTagColumns_()),
    newEntity: values => ({
      type: resolveTypeLabel(TAG_TYPE_NAMES, values.type) || 'googtag',
      parameter: []
    }),
    validate: draft => {
      const id = getGoogleTagId_(draft);
      if (!id) return ['Tag / Destination ID is required.'];
      const pattern = DESTINATION_ID_PATTERNS[draft.type];
      if (pattern && !/\{\{.+\}\}/.test(id) && !pattern.test(id)) {
        return [`"${id}" is not a valid ID for a ` +
            `${TAG_TYPE_NAMES[draft.type]} tag.`];
      }
      return [];
    }
  };
}

/**
 * Returns the tag ID (Google tag) or the destination ID (destination tags).
 * @param {!Object} tag
 * @return {string}
 */
function getGoogleTagId_(tag) {
  return getParamValue(tag, tag.type === 'googtag' ? 'tagId' :
      'destinationId');
}

/**
 * Column: tag ID of a Google tag or destination ID of a destination tag.
 * @param {!Object=} options
 * @return {!Object}
 */
function googleTagIdColumn_(options) {
  return buildColumn_({
    key: 'googleTagId',
    header: 'Tag / Destination ID',
    width: 160,
    required: true,
    note: 'Google tag: tag ID (e.g. G-XXXXXXXXXX, AW-123456789, GT-XXXXXXX).\n' +
        'Destination tags: G-… (Google Analytics), AW-… (Google Ads), ' +
        'DC-… (Floodlight) or MC-… (Merchant Center). A variable can be used.',
    get: entity => getGoogleTagId_(entity),
    set: (draft, value) => {
      setParamValue(draft, draft.type === 'googtag' ? 'tagId' :
          'destinationId', value);
    }
  }, options);
}
