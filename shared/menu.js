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
 * @fileoverview Menu, simple triggers and the entry points of all menu
 * actions.
 */

/**
 * Builds the menu.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const workspaceMenu = ui.createMenu('Select workspace')
      .addItem('1. List GTM accounts', 'listGtmAccounts')
      .addItem('2. List containers (of selected accounts)',
          'listGtmContainers')
      .addItem('3. List workspaces (of selected container)',
          'listGtmWorkspaces')
      .addSeparator()
      .addItem('Find container by ID…', 'findContainerById')
      .addItem('Create new workspace…', 'createGtmWorkspace');
  const toolsMenu = ui.createMenu('Tools')
      .addItem('Refresh dropdown lists (triggers, folders, variables, tags)',
          'refreshDropdownLists')
      .addItem('Clear Result column (this sheet)', 'clearResultsOfActiveSheet')
      .addSeparator()
      .addItem('Set up / repair all sheets', 'setupSheets');
  const helpMenu = ui.createMenu('Help')
      .addItem('How to use', 'showHelp')
      .addItem('Check for updates', 'checkForUpdates')
      .addItem('About', 'showAbout');
  ui.createMenu(APP_NAME)
      .addSubMenu(workspaceMenu)
      .addSeparator()
      .addItem('List from GTM (this sheet)', 'listActiveSheet')
      .addItem('Preview changes (this sheet)', 'previewActiveSheet')
      .addItem('Apply changes to GTM (this sheet)', 'applyActiveSheet')
      .addSeparator()
      .addSubMenu(toolsMenu)
      .addSubMenu(helpMenu)
      .addToUi();
}

/**
 * Simple trigger: keeps the workspace selection consistent.
 * @param {!Object} e The onEdit event.
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheetName = e.range.getSheet().getName();
    if (sheetName === WORKSPACE_SHEET.name) handleWorkspaceSheetEdit_(e);
  } catch (error) {
    console.log('onEdit: ' + error.message);
  }
}

/** Definitions of all feature sheets, cached for the execution. */
let SHEET_DEFS_CACHE_ = null;

/**
 * Returns the definitions of all feature sheets, in the order of the tabs.
 * @return {!Array<!Object>}
 */
function getSheetDefs() {
  if (!SHEET_DEFS_CACHE_) {
    SHEET_DEFS_CACHE_ = [
      tagsSheetDef_(),
      ga4EventTagsSheetDef_(),
      eventParametersSheetDef_(),
      googleTagsSheetDef_(),
      adsFloodlightTagsSheetDef_(),
      serverGa4TagsSheetDef_(),
      serverAdsFloodlightTagsSheetDef_(),
      clientsSheetDef_(),
      transformationsSheetDef_(),
      variablesSheetDef_(),
      triggersSheetDef_(),
      advancedSettingsSheetDef_(),
      googleTagSettingsSheetDef_()
    ];
  }
  return SHEET_DEFS_CACHE_;
}

/**
 * Returns the definition of a feature sheet by sheet name.
 * @param {string} sheetName
 * @return {?Object}
 */
function getSheetDefByName_(sheetName) {
  return getSheetDefs().find(def => def.sheetName === sheetName) || null;
}

/**
 * Runs a menu action with consistent error handling.
 * @param {string} title Title of the action, used in messages.
 * @param {function(): (string|{alert: string}|undefined)} action Returns a
 *     message for the user (shown as a toast, or as an alert).
 */
function runMenuAction_(title, action) {
  startExecutionTimer();
  const ui = SpreadsheetApp.getUi();
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(3000)) {
    ui.alert(title, 'Another action is still running in this spreadsheet. ' +
        'Please wait until it has finished.', ui.ButtonSet.OK);
    return;
  }
  try {
    const result = action();
    if (result && typeof result === 'object' && result.alert) {
      ui.alert(title, result.alert, ui.ButtonSet.OK);
    } else if (result) {
      SpreadsheetApp.getActive().toast(String(result), title, 10);
    }
  } catch (e) {
    if (e instanceof UserError) {
      ui.alert(title, e.message, ui.ButtonSet.OK);
    } else if (e instanceof GtmApiError) {
      ui.alert(title, 'Tag Manager API error: ' + e.message, ui.ButtonSet.OK);
    } else {
      console.error(e && e.stack ? e.stack : e);
      ui.alert(title, 'Unexpected error: ' + (e && e.message ? e.message : e) +
          '\n\nIf the problem persists, please report it on GitHub ' +
          `(github.com/${APP_GITHUB_REPO}/issues).`, ui.ButtonSet.OK);
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * Returns the definition of the active feature sheet, or throws.
 * @return {!Object}
 */
function getActiveSheetDef_() {
  const sheet = SpreadsheetApp.getActive().getActiveSheet();
  const def = getSheetDefByName_(sheet.getName());
  if (!def) {
    const names = getSheetDefs().map(d => '• ' + d.sheetName).join('\n');
    throw new UserError('Open one of the feature sheets first, then use the ' +
        `menu again. Feature sheets:\n${names}`);
  }
  return def;
}

/**
 * Menu action: lists the items of the active sheet from GTM.
 */
function listActiveSheet() {
  runMenuAction_('List from GTM', () =>
    listFeatureSheet_(getActiveSheetDef_()));
}

/**
 * Menu action: previews the changes of the active sheet.
 */
function previewActiveSheet() {
  runMenuAction_('Preview changes', () =>
    applyFeatureSheet_(getActiveSheetDef_(), {preview: true}));
}

/**
 * Menu action: applies the changes of the active sheet to GTM.
 */
function applyActiveSheet() {
  runMenuAction_('Apply changes to GTM', () => ({
    alert: applyFeatureSheet_(getActiveSheetDef_(), {preview: false})
  }));
}

/**
 * Menu action: clears the Result column of the active sheet.
 */
function clearResultsOfActiveSheet() {
  runMenuAction_('Clear results', () => {
    const def = getActiveSheetDef_();
    const sheet = getFeatureSheet_(def);
    const lastRow = sheet.getLastRow();
    if (lastRow >= LAYOUT.FIRST_DATA_ROW) {
      sheet.getRange(LAYOUT.FIRST_DATA_ROW,
          getColumnIndex_(def, SYSTEM_COLUMNS.result.key),
          lastRow - LAYOUT.FIRST_DATA_ROW + 1, 1).clearContent();
    }
    return 'The Result column was cleared.';
  });
}

/**
 * Menu action: shows how to use the tool.
 */
function showHelp() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('How to use ' + APP_NAME,
      '1. Select a workspace in the "GTM Workspace" sheet (menu › Select ' +
      'workspace). Tip: create a new workspace for bulk changes.\n\n' +
      '2. Open a feature sheet and use "List from GTM".\n\n' +
      '3. Edit cells. Edited rows are detected automatically. To create an ' +
      'item, add a row and set the Action to "Create" (copy an existing row ' +
      'to duplicate an item with all its settings). To delete an item, set ' +
      'the Action to "Delete".\n\n' +
      '4. Use "Preview changes" to see what will change in the Result ' +
      'column, then "Apply changes to GTM".\n\n' +
      '5. Review the changes in Tag Manager (the workspace shows every ' +
      'change) and publish them there. The Changelog sheet lists all ' +
      'changes made with this tool.\n\n' +
      'Hover over the column headers to see the expected values. Lists are ' +
      'separated by commas (or new lines); tables use one "name = value" per ' +
      'line.', ui.ButtonSet.OK);
}

/**
 * Menu action: shows information about the tool.
 */
function showAbout() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('About', `${APP_NAME} ${APP_VERSION}\n\n` +
      'Bulk edit Google Tag Manager web and server containers from Google ' +
      'Sheets.\n\nThis is not an officially supported Google product.\n\n' +
      `https://github.com/${APP_GITHUB_REPO}`, ui.ButtonSet.OK);
}

/**
 * Menu action: checks GitHub for a newer release of the tool.
 */
function checkForUpdates() {
  runMenuAction_('Check for updates', () => {
    const response = UrlFetchApp.fetch(
        `https://api.github.com/repos/${APP_GITHUB_REPO}/releases/latest`,
        {muteHttpExceptions: true});
    if (response.getResponseCode() !== 200) {
      return 'Could not check for updates right now.';
    }
    const release = JSON.parse(response.getContentText());
    const latest = String(release.tag_name || '');
    const toNumbers = version => version.replace(/^v/, '').split('.')
        .map(part => parseInt(part, 10) || 0);
    const current = toNumbers(APP_VERSION);
    const available = toNumbers(latest);
    let newer = false;
    for (let i = 0; i < Math.max(current.length, available.length); i++) {
      const a = available[i] || 0;
      const c = current[i] || 0;
      if (a !== c) {
        newer = a > c;
        break;
      }
    }
    return newer ? {alert: `A new version (${latest}) is available:\n\n` +
        `${release.body || ''}\n\n${release.html_url}`} :
        `You are using the latest version (${APP_VERSION}).`;
  });
}
