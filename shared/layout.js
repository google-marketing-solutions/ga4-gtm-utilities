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
 * @fileoverview Creates and repairs the layout of all sheets (headers, notes,
 * dropdowns, formats, frozen rows and columns, tab colours). The layout is
 * generated from the sheet definitions, so the code is the single source of
 * truth for the structure of the spreadsheet.
 */

/** Hidden sheet with the values of the dropdown lists. */
const LISTS_SHEET = {
  name: 'Lists',
  lists: ['Triggers', 'Folders', 'Variables', 'Tags']
};

/** Sheet with the release information. */
const SETTINGS_SHEET = {name: 'Settings'};

/** Sheets of version 2 of the tool, replaced by the current sheets. */
const OBSOLETE_SHEET_NAMES = [
  'Modify Parameter or User Property Settings',
  'Event Tag Settings',
  '(sGTM) Client Settings',
  '(sGTM) GA4 Tag Settings',
  '(sGTM) FL/Ads Tag Settings',
  '(sGTM) Other Tag Settings',
  'Variable Usage',
  'Validation Settings',
  'Tag Data Dictionary'
];

/** Notes of the system columns. */
const SYSTEM_COLUMN_NOTES = {
  _action: 'Choose what to do with the row when you use "Apply changes":\n' +
      '• Create: create a new item. To copy an existing item, copy its row ' +
      '(including the ID), change the values and choose Create.\n' +
      '• Update: update the item. Edited rows are detected automatically, ' +
      'so you only need this to force an update.\n' +
      '• Delete: delete the item.',
  _result: 'Result of the last Preview or Apply.',
  _id: 'ID of the item in GTM. Do not change it.',
  _sync: 'Technical data used to detect edited rows. Do not edit.'
};

/**
 * Menu action: creates or repairs the layout of all sheets.
 */
function setupSheets() {
  runMenuAction_('Set up sheets', () => {
    setupAllSheets_({removeObsolete: 'ask'});
    return 'All sheets are set up.';
  });
}

/**
 * Creates or repairs all sheets.
 * @param {{removeObsolete: string}} options removeObsolete: 'ask', 'yes' or
 *     'no'.
 */
function setupAllSheets_(options) {
  const ss = SpreadsheetApp.getActive();
  migrateChangelogSheet_();
  setupWorkspaceSheet_();
  getSheetDefs().forEach(def => setupFeatureSheet_(def));
  setupChangelogSheet_();
  setupListsSheet_();
  setupSettingsSheet_();

  const obsolete = OBSOLETE_SHEET_NAMES
      .map(name => ss.getSheetByName(name)).filter(Boolean);
  let remove = options.removeObsolete === 'yes';
  if (obsolete.length && options.removeObsolete === 'ask') {
    const ui = SpreadsheetApp.getUi();
    remove = ui.alert('Remove old sheets?',
        'These sheets are from the previous version of the tool and are ' +
        'replaced by the new sheets:\n\n' +
        obsolete.map(s => '• ' + s.getName()).join('\n') +
        '\n\nRemove them now?', ui.ButtonSet.YES_NO) === ui.Button.YES;
  }
  if (remove) obsolete.forEach(sheet => ss.deleteSheet(sheet));
  // Remove the empty default sheet of a new spreadsheet.
  ['Sheet1', 'Sheet 1'].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet && sheet.getLastRow() === 0 && ss.getSheets().length > 1) {
      ss.deleteSheet(sheet);
    }
  });

  // Order of the sheet tabs.
  const order = [WORKSPACE_SHEET.name]
      .concat(getSheetDefs().map(def => def.sheetName))
      .concat([CHANGELOG_SHEET.name, LISTS_SHEET.name, SETTINGS_SHEET.name]);
  order.forEach((name, index) => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(index + 1);
    }
  });
  ss.setActiveSheet(ss.getSheetByName(WORKSPACE_SHEET.name));
  // Technical sheets are hidden after moving them (activating a sheet shows
  // it).
  [LISTS_SHEET.name, SETTINGS_SHEET.name].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet && !sheet.isSheetHidden()) sheet.hideSheet();
  });
}

/**
 * Returns a sheet, creating it if needed.
 * @param {string} name
 * @return {!SpreadsheetApp.Sheet}
 */
function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/**
 * Makes sure that a sheet has at least the given number of rows.
 * @param {!SpreadsheetApp.Sheet} sheet
 * @param {number} rows
 */
function ensureSheetRows_(sheet, rows) {
  const maxRows = sheet.getMaxRows();
  if (maxRows < rows) sheet.insertRowsAfter(maxRows, rows - maxRows);
}

/**
 * Makes sure that a sheet has at least the given number of columns.
 * @param {!SpreadsheetApp.Sheet} sheet
 * @param {number} columns
 */
function ensureSheetColumns_(sheet, columns) {
  const maxColumns = sheet.getMaxColumns();
  if (maxColumns < columns) {
    sheet.insertColumnsAfter(maxColumns, columns - maxColumns);
  }
}

/**
 * Removes merged cells, notes and validations from a range (to rebuild it).
 * @param {!SpreadsheetApp.Range} range
 */
function resetRange_(range) {
  range.breakApart();
  range.clearNote();
  range.clearDataValidations();
}

/**
 * Removes sheet-level formatting that is not part of the cells: conditional
 * format rules, alternating colours (banding) and filters. Sheets of older
 * versions of the tool had conditional format rules that would otherwise
 * apply to the new columns (e.g. hide text).
 * @param {!SpreadsheetApp.Sheet} sheet
 */
function resetSheetFormatting_(sheet) {
  sheet.clearConditionalFormatRules();
  try {
    sheet.getBandings().forEach(banding => banding.remove());
  } catch (e) {
    // Banding is optional.
  }
  try {
    const filter = sheet.getFilter();
    if (filter) filter.remove();
  } catch (e) {
    // Filters are optional.
  }
}

/**
 * Returns the letter(s) of a 1-based column number, e.g. 1 -> A, 27 -> AA.
 * @param {number} column
 * @return {string}
 */
function columnLetter_(column) {
  let letters = '';
  let n = column;
  while (n > 0) {
    const remainder = (n - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

/**
 * Creates or repairs a feature sheet.
 * @param {!Object} def
 */
function setupFeatureSheet_(def) {
  const sheet = getOrCreateSheet_(def.sheetName);
  const columns = getAllColumns_(def);
  const width = columns.length;
  ensureSheetColumns_(sheet, width);
  ensureSheetRows_(sheet, 500);

  // Clear the data if the columns have changed (e.g. after an update).
  const expectedHeaders = columns.map(col => col.header);
  const currentHeaders = sheet.getRange(LAYOUT.HEADER_ROW, 1, 1, width)
      .getValues()[0].map(toText);
  if (sheet.getLastRow() >= LAYOUT.FIRST_DATA_ROW &&
      currentHeaders.join('|') !== expectedHeaders.join('|')) {
    sheet.getRange(LAYOUT.FIRST_DATA_ROW, 1,
        sheet.getLastRow() - LAYOUT.FIRST_DATA_ROW + 1, sheet.getMaxColumns())
        .clearContent();
  }

  // Group row and header row.
  const headerArea = sheet.getRange(1, 1, 2, sheet.getMaxColumns());
  resetRange_(headerArea);
  headerArea.clearContent();
  const groupValues = columns.map(col => col.group || '');
  groupValues[0] = def.title || def.sheetName;
  sheet.getRange(LAYOUT.GROUP_ROW, 1, 1, width).setValues([groupValues]);
  sheet.getRange(LAYOUT.HEADER_ROW, 1, 1, width).setValues([expectedHeaders]);
  sheet.getRange(LAYOUT.GROUP_ROW, 1, 1, width)
      .setBackground(COLORS.groupBackground)
      .setFontColor(COLORS.groupText)
      .setFontWeight('bold')
      .setVerticalAlignment('middle');
  sheet.getRange(LAYOUT.HEADER_ROW, 1, 1, width)
      .setBackground(COLORS.headerBackground)
      .setFontColor(COLORS.headerText)
      .setFontWeight('bold')
      .setWrap(true)
      .setVerticalAlignment('middle');
  sheet.getRange(LAYOUT.HEADER_ROW, 1, 1, 2)
      .setBackground(COLORS.systemHeaderBackground);
  sheet.getRange(LAYOUT.GROUP_ROW, 1, 1, 2).merge()
      .setFontWeight('normal').setFontStyle('italic');
  // Show each group label once, in the first column of the group. (Group
  // cells are not merged: Sheets can't freeze columns that contain only a
  // part of a merged cell.)
  const groupLabels = groupValues.map((group, index) =>
    index < 2 || group !== groupValues[index - 1] ? group : '');
  groupLabels[0] = groupValues[0];
  sheet.getRange(LAYOUT.GROUP_ROW, 1, 1, width).setValues([groupLabels]);
  // Alternate the group colours so that the groups can be told apart.
  let groupIndex = 0;
  const groupColors = groupValues.map((group, index) => {
    if (index >= 2 && group !== groupValues[index - 1]) groupIndex++;
    return index < 2 ? COLORS.groupBackground :
        (groupIndex % 2 ? COLORS.groupBackgroundAlt : COLORS.groupBackground);
  });
  sheet.getRange(LAYOUT.GROUP_ROW, 1, 1, width).setBackgrounds([groupColors]);
  const notes = columns.map(col => {
    let note = SYSTEM_COLUMN_NOTES[col.key] || col.note || '';
    if (col.required) note = (note ? note + '\n' : '') + 'Required.';
    if (col.readOnly) {
      note = (note ? note + '\n' : '') + 'Read-only (changes are ignored).';
    } else if (col.createOnly) {
      note = (note ? note + '\n' : '') +
          'Can only be set when creating a new item.';
    }
    return note;
  });
  sheet.getRange(LAYOUT.HEADER_ROW, 1, 1, width).setNotes([notes]);
  sheet.setRowHeight(LAYOUT.GROUP_ROW, 28);
  sheet.setRowHeight(LAYOUT.HEADER_ROW, 42);

  // Data rows: formats and dropdowns.
  applyDataRowFormats_(def, sheet, LAYOUT.FIRST_DATA_ROW,
      sheet.getMaxRows() - LAYOUT.FIRST_DATA_ROW + 1);

  // Colour the Action cells by value.
  resetSheetFormatting_(sheet);
  const actionRange = sheet.getRange(LAYOUT.FIRST_DATA_ROW, 1,
      sheet.getMaxRows() - LAYOUT.FIRST_DATA_ROW + 1, 1);
  sheet.setConditionalFormatRules([
    [ACTIONS.CREATE, COLORS.actionCreate],
    [ACTIONS.UPDATE, COLORS.actionUpdate],
    [ACTIONS.DELETE, COLORS.actionDelete]
  ].map(item => SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(item[0])
      .setBackground(item[1])
      .setRanges([actionRange])
      .build()));

  // Column widths and visibility.
  columns.forEach((col, index) => {
    sheet.setColumnWidth(index + 1, col.width || 150);
  });
  sheet.showColumns(1, width);
  sheet.hideColumns(getColumnIndex_(def, SYSTEM_COLUMNS.sync.key));
  if (sheet.getMaxColumns() > width) {
    sheet.deleteColumns(width + 1, sheet.getMaxColumns() - width);
  }
  sheet.setFrozenRows(LAYOUT.HEADER_ROW);
  const nameIndex = def.columns.findIndex(col => col.key === def.nameColumn);
  sheet.setFrozenColumns(nameIndex === -1 ? 2 : nameIndex + 3);
  sheet.setTabColor(def.tabColor);
  try {
    const protections = sheet.getProtections(
        SpreadsheetApp.ProtectionType.RANGE);
    protections.forEach(protection => {
      if (protection.getDescription() === 'Header rows') protection.remove();
    });
    sheet.getRange(1, 1, 2, width).protect().setDescription('Header rows')
        .setWarningOnly(true);
  } catch (e) {
    // Protection is optional.
  }
}

/**
 * Applies number formats and dropdowns to data rows.
 * @param {!Object} def
 * @param {!SpreadsheetApp.Sheet} sheet
 * @param {number} firstRow
 * @param {number} numRows
 */
function applyDataRowFormats_(def, sheet, firstRow, numRows) {
  if (numRows <= 0) return;
  const columns = getAllColumns_(def);
  const listsSheet = getOrCreateSheet_(LISTS_SHEET.name);
  columns.forEach((col, index) => {
    const range = sheet.getRange(firstRow, index + 1, numRows, 1);
    if (col.kind !== 'checkbox') {
      range.clearDataValidations();
      range.setNumberFormat('@');
    }
    range.setWrap(false);
    range.setVerticalAlignment('top');
    let rule = null;
    if (col.key === SYSTEM_COLUMNS.action.key) {
      rule = SpreadsheetApp.newDataValidation()
          .requireValueInList([ACTIONS.CREATE, ACTIONS.UPDATE, ACTIONS.DELETE],
              true)
          .setAllowInvalid(false)
          .build();
    } else if (Array.isArray(col.dropdown) && col.dropdown.length) {
      rule = SpreadsheetApp.newDataValidation()
          .requireValueInList(col.dropdown, true)
          .setAllowInvalid(!col.strictDropdown)
          .build();
    } else if (typeof col.dropdown === 'string') {
      const listIndex = LISTS_SHEET.lists.indexOf(col.dropdown);
      if (listIndex !== -1) {
        rule = SpreadsheetApp.newDataValidation()
            .requireValueInRange(listsSheet.getRange(2, listIndex + 1,
                listsSheet.getMaxRows() - 1, 1), true)
            .setAllowInvalid(true)
            .build();
      }
    }
    if (rule) range.setDataValidation(rule);
    if (col.key === SYSTEM_COLUMNS.id.key || col.readOnly) {
      range.setFontColor(COLORS.readOnlyText);
    }
  });
}

/**
 * Creates or repairs the GTM Workspace sheet.
 */
function setupWorkspaceSheet_() {
  const sheet = getOrCreateSheet_(WORKSPACE_SHEET.name);
  const blocks = [WORKSPACE_SHEET.accounts, WORKSPACE_SHEET.containers,
                  WORKSPACE_SHEET.workspaces];
  const panel = WORKSPACE_SHEET.panel;
  const width = panel.column + 1;
  ensureSheetColumns_(sheet, width);
  ensureSheetRows_(sheet, 200);

  // Keep the lists if the layout is already the current one.
  const currentHeaders = sheet.getRange(2, 1, 1, width).getValues()[0]
      .map(toText);
  const isCurrent = blocks.every(block => block.headers.every((header, i) =>
    currentHeaders[block.column - 1 + i] === header));
  if (!isCurrent && sheet.getLastRow() > 0) {
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
        .clearContent();
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
        .clearDataValidations();
  }
  const headerArea = sheet.getRange(1, 1, 2, sheet.getMaxColumns());
  resetRange_(headerArea);
  headerArea.clearContent();
  headerArea.setBackground(null);

  blocks.forEach(block => {
    const blockWidth = block.headers.length;
    sheet.getRange(1, block.column).setValue(block.title);
    sheet.getRange(1, block.column, 1, blockWidth).merge();
    sheet.getRange(2, block.column, 1, blockWidth).setValues([block.headers]);
    sheet.getRange(1, block.column, 1, blockWidth)
        .setBackground(COLORS.groupBackground).setFontColor(COLORS.groupText)
        .setFontWeight('bold');
    sheet.getRange(2, block.column, 1, blockWidth)
        .setBackground(COLORS.headerBackground).setFontWeight('bold');
    block.widths.forEach((columnWidth, i) => {
      sheet.setColumnWidth(block.column + i, columnWidth);
    });
    sheet.getRange(WORKSPACE_SHEET.firstDataRow, block.column,
        sheet.getMaxRows() - WORKSPACE_SHEET.firstDataRow + 1, blockWidth - 1)
        .setNumberFormat('@');
  });
  // Spacer column between the workspaces block and the panel.
  sheet.setColumnWidth(panel.column - 1, 24);
  sheet.getRange(1, panel.column).setValue(panel.title);
  sheet.getRange(1, panel.column, 1, 2).merge()
      .setBackground(TAB_COLORS.general).setFontColor('#ffffff')
      .setFontWeight('bold');
  panel.widths.forEach((columnWidth, i) => {
    sheet.setColumnWidth(panel.column + i, columnWidth);
  });
  sheet.getRange(WORKSPACE_SHEET.firstDataRow, panel.column,
      panel.labels.length, 1).setFontWeight('bold');
  const helpRow = WORKSPACE_SHEET.firstDataRow + panel.labels.length + 1;
  const help = [
    ['HOW TO USE', ''],
    ['1. Select a workspace', 'Menu GA4 GTM Utilities › Select workspace: ' +
     'list accounts, containers and workspaces and tick "Select" (or use ' +
     '"Find container by ID").'],
    ['2. List', 'Open a feature sheet (e.g. Tags) and use menu › List from ' +
     'GTM.'],
    ['3. Edit', 'Edit cells, add rows with Action "Create" (copy a row to ' +
     'duplicate an item) or set Action "Delete".'],
    ['4. Preview', 'Menu › Preview changes shows what will change in the ' +
     'Result column.'],
    ['5. Apply', 'Menu › Apply changes to GTM. Changes are made in the ' +
     'workspace; review and publish them in Tag Manager.'],
    ['Tip', 'Hover over the column headers to see the expected values.']
  ];
  sheet.getRange(helpRow, panel.column, help.length, 2).setValues(help);
  sheet.getRange(helpRow, panel.column, 1, 2).merge()
      .setBackground(TAB_COLORS.general).setFontColor('#ffffff')
      .setFontWeight('bold');
  sheet.getRange(helpRow + 1, panel.column, help.length - 1, 1)
      .setFontWeight('bold');
  sheet.getRange(helpRow + 1, panel.column + 1, help.length - 1, 1)
      .setWrap(true);
  sheet.setFrozenRows(2);
  sheet.setTabColor(TAB_COLORS.general);

  // Highlight the selected account(s), container(s) and workspace.
  resetSheetFormatting_(sheet);
  const numRows = sheet.getMaxRows() - WORKSPACE_SHEET.firstDataRow + 1;
  sheet.setConditionalFormatRules(blocks.map(block => {
    const selectColumn = columnLetter_(block.column + block.headers.length - 1);
    return SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(
            `=$${selectColumn}${WORKSPACE_SHEET.firstDataRow}=TRUE`)
        .setBackground(COLORS.selected)
        .setRanges([sheet.getRange(WORKSPACE_SHEET.firstDataRow, block.column,
                                   numRows, block.headers.length)])
        .build();
  }));
  updateSelectedWorkspacePanel_();
}

/**
 * Converts the Changelog sheet of version 2 to the current format.
 */
function migrateChangelogSheet_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(
      CHANGELOG_SHEET.name);
  if (!sheet || sheet.getLastRow() < 1) return;
  const header = sheet.getRange(1, 1, 1, 2).getValues()[0].map(toText);
  if (header[0] === CHANGELOG_SHEET.headers[0]) return;
  // Version 2: Date, Name, Type, ID, Action Taken, URL, User (no header row
  // in some copies).
  const lastRow = sheet.getLastRow();
  const oldValues = sheet.getRange(1, 1, lastRow, 7).getValues();
  const rows = oldValues
      .filter(row => !isBlank(row[0]) && toText(row[0]) !== 'Date')
      .map(row => [toText(row[0]), toText(row[6]), '', '', toText(row[2]),
                   toText(row[1]), toText(row[3]),
                   truncateText(toText(row[4]), 200), '', toText(row[5])]);
  sheet.clear();
  sheet.getRange(1, 1, 1, CHANGELOG_SHEET.headers.length)
      .setValues([CHANGELOG_SHEET.headers]);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, CHANGELOG_SHEET.headers.length)
        .setValues(rows);
  }
}

/**
 * Creates or repairs the Changelog sheet.
 */
function setupChangelogSheet_() {
  const sheet = getOrCreateSheet_(CHANGELOG_SHEET.name);
  resetSheetFormatting_(sheet);
  const width = CHANGELOG_SHEET.headers.length;
  ensureSheetColumns_(sheet, width);
  sheet.getRange(1, 1, 1, width).setValues([CHANGELOG_SHEET.headers])
      .setBackground(COLORS.headerBackground).setFontWeight('bold');
  CHANGELOG_SHEET.widths.forEach((columnWidth, index) => {
    sheet.setColumnWidth(index + 1, columnWidth);
  });
  sheet.getRange(2, 1, Math.max(1, sheet.getMaxRows() - 1), width)
      .setNumberFormat('@').setWrap(false).setVerticalAlignment('top');
  sheet.setFrozenRows(1);
  sheet.setTabColor(TAB_COLORS.technical);
}

/**
 * Creates or repairs the hidden Lists sheet.
 */
function setupListsSheet_() {
  const sheet = getOrCreateSheet_(LISTS_SHEET.name);
  resetSheetFormatting_(sheet);
  ensureSheetColumns_(sheet, LISTS_SHEET.lists.length);
  ensureSheetRows_(sheet, 1000);
  sheet.getRange(1, 1, 1, LISTS_SHEET.lists.length)
      .setValues([LISTS_SHEET.lists]).setFontWeight('bold');
  sheet.getRange(2, 1, sheet.getMaxRows() - 1, LISTS_SHEET.lists.length)
      .setNumberFormat('@');
  sheet.setTabColor(TAB_COLORS.technical);
  sheet.hideSheet();
}

/**
 * Creates or repairs the Settings sheet.
 */
function setupSettingsSheet_() {
  const sheet = getOrCreateSheet_(SETTINGS_SHEET.name);
  resetSheetFormatting_(sheet);
  const current = sheet.getRange(2, 2).getValue();
  sheet.getRange(1, 1, 2, 2).setValues([
    ['Release Version', APP_VERSION],
    ['Update notification', isBlank(current) ? 'Not Set' : current]
  ]);
  sheet.getRange(1, 1, 2, 1).setFontWeight('bold');
  sheet.setColumnWidth(1, 160);
  sheet.setTabColor(TAB_COLORS.technical);
  sheet.hideSheet();
}

/**
 * Writes a list to the Lists sheet.
 * @param {string} name A name of LISTS_SHEET.lists.
 * @param {!Array<string>} values
 */
function writeDropdownList_(name, values) {
  const sheet = getOrCreateSheet_(LISTS_SHEET.name);
  const column = LISTS_SHEET.lists.indexOf(name) + 1;
  const unique = [];
  values.forEach(value => {
    if (value && unique.indexOf(value) === -1) unique.push(value);
  });
  ensureSheetRows_(sheet, unique.length + 1);
  sheet.getRange(2, column, sheet.getMaxRows() - 1, 1).clearContent();
  if (unique.length) {
    sheet.getRange(2, column, unique.length, 1)
        .setValues(unique.map(value => [value]));
  }
}

/**
 * Updates the dropdown lists (triggers, folders, variables and tags) from the
 * workspace. Only lists that are already loaded are updated, so that no extra
 * API calls are made.
 * @param {!WorkspaceData} data
 * @param {!Object} ctx
 * @param {boolean=} loadAll Load all lists from the API.
 */
function refreshDropdownLists_(data, ctx, loadAll) {
  try {
    // All lists are loaded from the API when the workspace changed or when
    // they were last loaded more than 10 minutes ago.
    const properties = PropertiesService.getDocumentProperties();
    const marker = properties.getProperty(PROPERTY_KEYS.dropdownLists) || '';
    const [markerPath, markerTime] = marker.split('|');
    const stale = markerPath !== ctx.workspacePath ||
        Date.now() - Number(markerTime || 0) > 10 * 60 * 1000;
    const all = Boolean(loadAll) || stale;
    const loaded = kind => all || Boolean(data.cache_[kind]);
    if (loaded('trigger')) {
      writeDropdownList_('Triggers', data.getBuiltInTriggers()
          .map(t => t.name)
          .concat(data.list('trigger').map(t => t.name).sort(compareText)));
    }
    if (loaded('folder')) {
      writeDropdownList_('Folders',
          data.list('folder').map(f => f.name).sort(compareText));
    }
    if (loaded('variable')) {
      const builtIns = all || data.cache_.builtInVariable ?
          data.list('builtInVariable').map(v => v.name) : [];
      writeDropdownList_('Variables', data.list('variable')
          .map(v => v.name).concat(builtIns).sort(compareText)
          .map(name => `{{${name}}}`));
    }
    if (loaded('tag')) {
      writeDropdownList_('Tags',
          data.list('tag').map(t => t.name).sort(compareText));
    }
    if (all) {
      properties.setProperty(PROPERTY_KEYS.dropdownLists,
          `${ctx.workspacePath}|${Date.now()}`);
    }
  } catch (e) {
    // Dropdown lists are a convenience; never fail the main action.
    console.log('Could not refresh dropdown lists: ' + e.message);
  }
}

/**
 * Menu action: refreshes all dropdown lists from the selected workspace.
 */
function refreshDropdownLists() {
  runMenuAction_('Refresh dropdown lists', () => {
    const ctx = getWorkspaceContext();
    const data = new WorkspaceData(ctx);
    refreshDropdownLists_(data, ctx, true);
    return 'The dropdown lists (triggers, folders, variables and tags) were ' +
        'refreshed.';
  });
}
