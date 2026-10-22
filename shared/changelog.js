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
 * @fileoverview The Changelog sheet: a record of every change made by the
 * tool.
 */

const CHANGELOG_SHEET = {
  name: 'Changelog',
  headers: ['Date & Time', 'User', 'Container', 'Workspace', 'Item Type',
            'Item Name', 'ID', 'Action', 'Details', 'Link'],
  widths: [140, 180, 220, 160, 120, 260, 60, 80, 420, 120]
};

/**
 * Returns the Changelog sheet, creating it if needed.
 * @return {!SpreadsheetApp.Sheet}
 */
function getChangelogSheet_() {
  let sheet = SpreadsheetApp.getActive().getSheetByName(CHANGELOG_SHEET.name);
  if (!sheet) {
    setupChangelogSheet_();
    sheet = SpreadsheetApp.getActive().getSheetByName(CHANGELOG_SHEET.name);
  }
  return sheet;
}

/**
 * Adds the applied changes to the Changelog sheet.
 * @param {!Array<!Object>} changes
 * @param {!Object} ctx
 */
function logChanges_(changes, ctx) {
  const time = Utilities.formatDate(new Date(), ctx.timeZone,
      'yyyy-MM-dd HH:mm:ss');
  let user = '';
  try {
    user = Session.getActiveUser().getEmail();
  } catch (e) {
    user = '';
  }
  const actions = {added: 'Created', updated: 'Updated', deleted: 'Deleted'};
  const rows = [];
  changes.forEach(change => {
    if (!change.ops.some(op => op.result && op.result.status === 'success')) {
      return;
    }
    const info = GTM_ENTITY_KINDS[change.kind];
    const entity = change.resultEntity || change.entity;
    const name = entity.name ||
        getTypeLabel(GTAG_CONFIG_TYPE_NAMES, entity.type) || '';
    const details = change.changeStatus === 'deleted' ? '' :
        truncateText(formatDiffs_(change.diffs, true), 2000);
    rows.push([time, user, `${ctx.containerName} (${ctx.publicId})`,
               ctx.workspaceName, info.label, name,
               String(entity[info.idField] || ''),
               actions[change.changeStatus], details,
               entity.tagManagerUrl || '']);
  });
  if (!rows.length) return;
  const sheet = getChangelogSheet_();
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length)
      .setValues(rows);
}
