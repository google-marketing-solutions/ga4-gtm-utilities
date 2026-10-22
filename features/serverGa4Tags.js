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
 * @fileoverview The "Server GA4 Tags" sheet (server containers).
 */

/** Values of the "Default parameters to include" settings. */
const SERVER_GA4_INCLUDE_LABELS = {all: 'All', none: 'None'};

/**
 * Returns the definition of the Server GA4 Tags sheet.
 * @return {!Object}
 */
function serverGa4TagsSheetDef_() {
  const event = 'EVENT';
  const params = 'EVENT PARAMETERS';
  const props = 'USER PROPERTIES';
  return {
    key: 'serverGa4Tags',
    sheetName: 'Server GA4 Tags',
    title: 'GA4 TAGS (SERVER)',
    tabColor: TAB_COLORS.server,
    mode: 'entity',
    kind: 'tag',
    contexts: ['server'],
    itemLabel: 'GA4 tag',
    itemLabelPlural: 'GA4 tags',
    idHeader: 'Tag ID',
    nameColumn: 'name',
    filter: tag => tag.type === 'sgtmgaaw',
    columns: [
      COLUMNS.name('Tag', {group: 'TAG'}),
      COLUMNS.param('measurementId', 'Measurement ID', {
        group: event, width: 170, dropdown: 'Variables',
        note: 'G-XXXXXXXXXX. Leave empty to use the measurement ID of the ' +
            'incoming event.'
      }),
      COLUMNS.param('eventName', 'Event Name', {
        group: event, width: 150,
        note: 'Leave empty to use the name of the incoming event.'
      }),
      COLUMNS.param('appId', 'App ID', {
        group: event, width: 120,
        note: 'Firebase app ID. Leave empty to use the app ID of the ' +
            'incoming event.'
      }),
      COLUMNS.boolParam('redactVisitorIp', 'Redact Visitor IP Address', {
        group: event, width: 110
      }),
      COLUMNS.selectParam('epToIncludeDropdown',
          'Default Parameters to Include', SERVER_GA4_INCLUDE_LABELS, {
            group: params, width: 120, defaultValue: 'all',
            note: 'All: send all parameters of the event (except the ' +
                'excluded ones). None: send only the parameters listed in ' +
                '"Parameters to Include".'
          }),
      COLUMNS.singleFieldTableParam('epToInclude', 'Parameters to Include',
          'fieldName', {
            group: params,
            note: 'Comma-separated parameter names (used when "Default ' +
                'Parameters to Include" is None).'
          }),
      COLUMNS.singleFieldTableParam('epToExclude', 'Parameters to Exclude',
          'fieldName', {
            group: params,
            note: 'Comma-separated parameter names (used when "Default ' +
                'Parameters to Include" is All).'
          }),
      COLUMNS.tableParam('eventParameters', 'Parameters to Add / Edit',
          ['fieldName', 'value'], {group: params}),
      COLUMNS.selectParam('upToIncludeDropdown',
          'Default User Properties to Include', SERVER_GA4_INCLUDE_LABELS, {
            group: props, width: 120, defaultValue: 'all'
          }),
      COLUMNS.singleFieldTableParam('upToInclude', 'User Properties to Include',
          'fieldName', {group: props}),
      COLUMNS.singleFieldTableParam('upToExclude', 'User Properties to Exclude',
          'fieldName', {group: props}),
      COLUMNS.tableParam('userProperties', 'User Properties to Add / Edit',
          ['fieldName', 'value'], {group: props})
    ].concat(commonTagColumns_()),
    newEntity: () => ({
      type: 'sgtmgaaw',
      parameter: [
        {type: 'template', key: 'epToIncludeDropdown', value: 'all'},
        {type: 'template', key: 'upToIncludeDropdown', value: 'all'}
      ]
    })
  };
}
