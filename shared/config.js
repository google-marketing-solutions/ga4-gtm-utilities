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
 * @fileoverview Global configuration and constants. This file only contains
 * literal values so that it can be loaded in any order.
 */

/** Name and version of the tool. */
const APP_NAME = 'GA4 GTM Utilities';
const APP_VERSION = 'v3.0';
const APP_GITHUB_REPO = 'google-marketing-solutions/ga4-gtm-utilities';

/** Values allowed in the "Action" column of every feature sheet. */
const ACTIONS = {
  CREATE: 'Create',
  UPDATE: 'Update',
  DELETE: 'Delete'
};

/** Row layout shared by all feature sheets. */
const LAYOUT = {
  GROUP_ROW: 1,
  HEADER_ROW: 2,
  FIRST_DATA_ROW: 3
};

/** Sheet tab colours. */
const TAB_COLORS = {
  general: '#188038',
  web: '#1a73e8',
  server: '#e8710a',
  technical: '#80868b'
};

/** Colours used inside the sheets. */
const COLORS = {
  headerBackground: '#e8eaed',
  headerText: '#202124',
  groupBackground: '#5f6368',
  groupBackgroundAlt: '#80868b',
  groupText: '#ffffff',
  systemHeaderBackground: '#fce8b2',
  notApplicable: '#f1f3f4',
  readOnlyText: '#5f6368',
  success: '#137333',
  error: '#c5221f',
  preview: '#1967d2',
  deleted: '#9aa0a6',
  selected: '#ceead6',
  actionCreate: '#e6f4ea',
  actionUpdate: '#fef7e0',
  actionDelete: '#fce8e6'
};

/**
 * Built-in triggers. They are not returned by the triggers.list API method but
 * can be referenced by their IDs in the firing and blocking triggers of tags.
 */
const BUILT_IN_TRIGGERS = [
  {triggerId: '2147479553', name: 'All Pages', context: 'web'},
  {triggerId: '2147479572', name: 'Consent Initialization - All Pages',
   context: 'web'},
  {triggerId: '2147479573', name: 'Initialization - All Pages', context: 'web'},
  {triggerId: '2147479574', name: 'All Pages', context: 'server'},
  {triggerId: '2147479576',
   name: 'Initialization - Support Legacy Config Behavior', context: 'web'}
];

/** Consent types that can be required by a tag (Additional consent checks). */
const CONSENT_TYPES = [
  'ad_storage',
  'analytics_storage',
  'ad_user_data',
  'ad_personalization',
  'functionality_storage',
  'personalization_storage',
  'security_storage'
];

/** Friendly labels for the tag consent status. */
const CONSENT_STATUS_LABELS = {
  notSet: 'Not set',
  notNeeded: 'No additional consent required',
  needed: 'Require additional consent'
};

/** Friendly labels for the tag firing options. */
const TAG_FIRING_OPTION_LABELS = {
  oncePerEvent: 'Once per event',
  oncePerLoad: 'Once per page',
  unlimited: 'Unlimited'
};

/** Limits used to stay within the GTM API quota and the Apps Script limits. */
const LIMITS = {
  // The GTM API allows 0.25 queries per second per project, enforced on a
  // 100 second window.
  apiMaxRequestsPerWindow: 25,
  apiWindowMs: 100 * 1000,
  apiMaxRetries: 5,
  // Apps Script executions are stopped after 6 minutes. Work is stopped
  // gracefully before that, so that the results can be written to the sheet.
  executionBudgetMs: 5 * 60 * 1000,
  // Maximum number of changes sent in one workspaces.bulk_update call.
  bulkUpdateChunkSize: 200,
  // Google Sheets allows at most 50,000 characters per cell.
  maxCellLength: 49000
};

/** Text appended to values that are too long for a cell. */
const TRUNCATION_SUFFIX = '…[truncated: too long for a cell, edit it in ' +
    'Tag Manager]';

/** Key prefix for document properties used by the tool. */
const PROPERTY_KEYS = {
  listedWorkspace: 'listedWorkspace:',
  containerInfo: 'containerInfo:',
  apiCallTimes: 'gtmApiCallTimes',
  dropdownLists: 'dropdownListsLoaded'
};
