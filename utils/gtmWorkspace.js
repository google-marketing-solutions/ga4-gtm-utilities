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
 * @fileoverview The "GTM Workspace" sheet: selection of the account,
 * container and workspace that all other sheets work with.
 */

/** Layout of the GTM Workspace sheet. */
const WORKSPACE_SHEET = {
  name: 'GTM Workspace',
  firstDataRow: 3,
  accounts: {
    column: 1,
    title: '1. ACCOUNT',
    headers: ['Account Name', 'Account Path', 'Select'],
    widths: [220, 160, 60]
  },
  containers: {
    column: 4,
    title: '2. CONTAINER',
    headers: ['Container Name', 'Container ID', 'Type', 'Container Path',
              'Select'],
    widths: [240, 120, 90, 230, 60]
  },
  workspaces: {
    column: 9,
    title: '3. WORKSPACE',
    headers: ['Workspace Name', 'Workspace Path', 'Select'],
    widths: [200, 300, 60]
  },
  panel: {
    column: 13,
    title: 'SELECTED WORKSPACE',
    labels: ['Account', 'Container', 'Container ID', 'Container type',
             'Workspace', 'Workspace path', 'Open in Tag Manager'],
    widths: [150, 380]
  }
};

/**
 * Returns the GTM Workspace sheet, creating its layout if needed.
 * @return {!SpreadsheetApp.Sheet}
 */
function getWorkspaceSheet_() {
  let sheet = SpreadsheetApp.getActive().getSheetByName(WORKSPACE_SHEET.name);
  if (!sheet) {
    setupWorkspaceSheet_();
    sheet = SpreadsheetApp.getActive().getSheetByName(WORKSPACE_SHEET.name);
  }
  return sheet;
}

/**
 * Reads the rows of one block (accounts, containers or workspaces).
 * @param {!Object} block A block of WORKSPACE_SHEET.
 * @return {!Array<{row: number, values: !Array<*>, selected: boolean}>}
 */
function readWorkspaceBlock_(block) {
  const sheet = getWorkspaceSheet_();
  const lastRow = sheet.getLastRow();
  const numRows = lastRow - WORKSPACE_SHEET.firstDataRow + 1;
  if (numRows <= 0) return [];
  const width = block.headers.length;
  const values = sheet.getRange(WORKSPACE_SHEET.firstDataRow, block.column,
      numRows, width).getValues();
  const rows = [];
  values.forEach((rowValues, index) => {
    if (isBlank(rowValues[0]) && isBlank(rowValues[width - 2])) return;
    rows.push({
      row: WORKSPACE_SHEET.firstDataRow + index,
      values: rowValues,
      selected: toBool(rowValues[width - 1])
    });
  });
  return rows;
}

/**
 * Writes the rows of one block, replacing its previous content.
 * @param {!Object} block
 * @param {!Array<!Array<*>>} rows Rows without the Select column.
 * @param {number=} selectedIndex Index of the row to select.
 */
function writeWorkspaceBlock_(block, rows, selectedIndex) {
  const sheet = getWorkspaceSheet_();
  const width = block.headers.length;
  const lastRow = Math.max(sheet.getLastRow(), WORKSPACE_SHEET.firstDataRow);
  const oldRows = lastRow - WORKSPACE_SHEET.firstDataRow + 1;
  const clearRange = sheet.getRange(WORKSPACE_SHEET.firstDataRow, block.column,
      oldRows, width);
  clearRange.clearContent();
  clearRange.clearDataValidations();
  if (!rows.length) return;
  const selectRange = sheet.getRange(WORKSPACE_SHEET.firstDataRow,
      block.column + width - 1, rows.length, 1);
  selectRange.insertCheckboxes();
  const values = rows.map((row, index) =>
    row.concat([index === selectedIndex]));
  sheet.getRange(WORKSPACE_SHEET.firstDataRow, block.column, values.length,
      width).setValues(values);
}

/**
 * Returns a readable container type.
 * @param {!Object} container A Container resource.
 * @return {string}
 */
function getContainerTypeLabel_(container) {
  const features = container.features || {};
  const contexts = container.usageContext || [];
  if (features.supportTags === false && features.supportGtagConfigs) {
    return 'Google tag';
  }
  if (contexts.indexOf('server') !== -1) return 'Server';
  if (contexts.indexOf('web') !== -1) return 'Web';
  if (contexts.indexOf('amp') !== -1) return 'AMP';
  if (contexts.indexOf('android') !== -1 || contexts.indexOf('ios') !== -1) {
    return 'Mobile';
  }
  return contexts.join(', ');
}

/**
 * Stores information about a container so that it does not need to be
 * requested again.
 * @param {!Object} container A Container resource.
 */
function storeContainerInfo_(container) {
  const info = {
    name: container.name,
    publicId: container.publicId || '',
    usageContext: container.usageContext || [],
    features: container.features || {},
    tagIds: container.tagIds || [],
    typeLabel: getContainerTypeLabel_(container)
  };
  PropertiesService.getDocumentProperties().setProperty(
      PROPERTY_KEYS.containerInfo + container.path, JSON.stringify(info));
}

/**
 * Returns information about a container, from the cache or the API.
 * @param {string} containerPath
 * @return {!Object}
 */
function getContainerInfo_(containerPath) {
  const properties = PropertiesService.getDocumentProperties();
  const cached = properties.getProperty(
      PROPERTY_KEYS.containerInfo + containerPath);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      // Ignore and request the container again.
    }
  }
  const container = gtmApiRequest('get', containerPath);
  storeContainerInfo_(container);
  return JSON.parse(properties.getProperty(
      PROPERTY_KEYS.containerInfo + containerPath));
}

/**
 * Menu action: lists the GTM accounts of the user.
 */
function listGtmAccounts() {
  runMenuAction_('List GTM accounts', () => {
    // Accounts that only contain Google tags (gtag.js) are not listed, to
    // keep the list short; they can be opened with "Find container by ID".
    const accounts = gtmApiListAll('accounts', 'account');
    accounts.sort((a, b) => compareText(a.name, b.name));
    writeWorkspaceBlock_(WORKSPACE_SHEET.accounts,
        accounts.map(account => [account.name, account.path]),
        accounts.length === 1 ? 0 : undefined);
    writeWorkspaceBlock_(WORKSPACE_SHEET.containers, []);
    writeWorkspaceBlock_(WORKSPACE_SHEET.workspaces, []);
    updateSelectedWorkspacePanel_();
    activateWorkspaceSheet_();
    return accounts.length ?
        `Found ${accounts.length} account(s). Select one or more accounts, ` +
        'then use "2. List containers". (To open a Google tag, use "Find ' +
        'container by ID" with its G- or GT- ID.)' :
        'No GTM accounts were found for your Google account.';
  });
}

/**
 * Menu action: lists the containers of the selected accounts.
 */
function listGtmContainers() {
  runMenuAction_('List GTM containers', () => {
    const accounts = readWorkspaceBlock_(WORKSPACE_SHEET.accounts)
        .filter(row => row.selected);
    if (!accounts.length) {
      throw new UserError('Select at least one account in the "Select" ' +
          'column of the GTM Workspace sheet first.');
    }
    let containers = [];
    accounts.forEach(account => {
      containers = containers.concat(gtmApiListAll(
          `${account.values[1]}/containers`, 'container'));
    });
    containers.forEach(storeContainerInfo_);
    containers.sort((a, b) => compareText(a.name, b.name));
    writeWorkspaceBlock_(WORKSPACE_SHEET.containers, containers.map(c => [
      c.name, c.publicId || '', getContainerTypeLabel_(c), c.path
    ]), containers.length === 1 ? 0 : undefined);
    writeWorkspaceBlock_(WORKSPACE_SHEET.workspaces, []);
    updateSelectedWorkspacePanel_();
    activateWorkspaceSheet_();
    return containers.length ?
        `Found ${containers.length} container(s). Select a container, then ` +
        'use "3. List workspaces".' :
        'No containers were found in the selected account(s).';
  });
}

/**
 * Lists the workspaces of a container. Google tag containers do not have
 * workspaces; a single entry is returned for their settings.
 * @param {string} containerPath
 * @return {!Array<{name: string, path: string}>}
 */
function listContainerWorkspaces_(containerPath) {
  const info = getContainerInfo_(containerPath);
  if (info.typeLabel === 'Google tag') {
    return [{name: 'Google tag settings (applied immediately)',
             path: `${containerPath}/workspaces/0`}];
  }
  return gtmApiListAll(`${containerPath}/workspaces`, 'workspace');
}

/**
 * Menu action: lists the workspaces of the selected containers.
 */
function listGtmWorkspaces() {
  runMenuAction_('List GTM workspaces', () => {
    const containers = readWorkspaceBlock_(WORKSPACE_SHEET.containers)
        .filter(row => row.selected);
    if (!containers.length) {
      throw new UserError('Select a container in the "Select" column of the ' +
          'GTM Workspace sheet first.');
    }
    let workspaces = [];
    containers.forEach(container => {
      workspaces = workspaces.concat(listContainerWorkspaces_(
          toText(container.values[3])));
    });
    writeWorkspaceBlock_(WORKSPACE_SHEET.workspaces,
        workspaces.map(w => [w.name, w.path]),
        workspaces.length === 1 ? 0 : undefined);
    updateSelectedWorkspacePanel_();
    activateWorkspaceSheet_();
    return workspaces.length === 1 ?
        `Workspace "${workspaces[0].name}" is selected. Open a feature ` +
        'sheet and use "List from GTM".' :
        `Found ${workspaces.length} workspaces. Select the workspace you ` +
        'want to work in.';
  });
}

/**
 * Menu action: finds a container by its ID (e.g. GTM-XXXXXXX) and lists its
 * workspaces.
 */
function findContainerById() {
  runMenuAction_('Find container by ID', () => {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt('Find container by ID',
        'Enter a container ID (e.g. GTM-XXXXXXX) or a destination ID ' +
        '(e.g. G-XXXXXXXXXX, AW-123456789):',
        ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) return '';
    const id = toText(response.getResponseText()).toUpperCase();
    if (!id) return '';
    const query = /^(GTM|GT)-/.test(id) ? {tagId: id} : {destinationId: id};
    let container;
    try {
      container = gtmApiRequest('get', 'accounts/containers:lookup',
          {query: query});
    } catch (e) {
      if (e instanceof GtmApiError && (e.code === 404 || e.code === 403 ||
          e.code === 400)) {
        throw new UserError(`No container with the ID ${id} was found that ` +
            'you have access to.');
      }
      throw e;
    }
    storeContainerInfo_(container);
    const accountPath = `accounts/${container.accountId}`;
    let accountName = accountPath;
    try {
      accountName = gtmApiRequest('get', accountPath).name || accountPath;
    } catch (e) {
      // The account name is only informative.
    }
    writeWorkspaceBlock_(WORKSPACE_SHEET.accounts, [[accountName, accountPath]],
        0);
    writeWorkspaceBlock_(WORKSPACE_SHEET.containers, [[
      container.name, container.publicId || '',
      getContainerTypeLabel_(container), container.path
    ]], 0);
    const workspaces = listContainerWorkspaces_(container.path);
    writeWorkspaceBlock_(WORKSPACE_SHEET.workspaces,
        workspaces.map(w => [w.name, w.path]),
        workspaces.length === 1 ? 0 : undefined);
    updateSelectedWorkspacePanel_();
    activateWorkspaceSheet_();
    return `Found container "${container.name}" (${container.publicId}). ` +
        (workspaces.length === 1 ? 'Its only workspace is selected.' :
         'Select the workspace you want to work in.');
  });
}

/**
 * Menu action: creates a new workspace in the container of the selected
 * workspace (or the selected container) and selects it.
 */
function createGtmWorkspace() {
  runMenuAction_('Create workspace', () => {
    const selectedWorkspace = getSelectedWorkspaceRow_();
    const selectedContainer = readWorkspaceBlock_(WORKSPACE_SHEET.containers)
        .find(row => row.selected);
    let containerPath = '';
    if (selectedWorkspace) {
      containerPath = selectedWorkspace.path.split('/').slice(0, 4).join('/');
    } else if (selectedContainer) {
      containerPath = toText(selectedContainer.values[3]);
    }
    if (!containerPath) {
      throw new UserError('Select a container (or a workspace) in the GTM ' +
          'Workspace sheet first.');
    }
    if (getContainerInfo_(containerPath).typeLabel === 'Google tag') {
      throw new UserError('Google tag containers do not have workspaces.');
    }
    const ui = SpreadsheetApp.getUi();
    const defaultName = 'Bulk changes ' + Utilities.formatDate(new Date(),
        getSpreadsheetTimeZone_(), 'yyyy-MM-dd HH:mm');
    const response = ui.prompt('Create workspace',
        `Name of the new workspace (leave empty for "${defaultName}"):`,
        ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) return '';
    const name = toText(response.getResponseText()) || defaultName;
    let workspace;
    try {
      workspace = gtmApiRequest('post', `${containerPath}/workspaces`, {
        body: {name: name, description: `Created by ${APP_NAME}`}
      });
    } catch (e) {
      if (e instanceof GtmApiError && e.code !== 401) {
        throw new UserError('The workspace could not be created: ' +
            e.message + ' (free containers can have up to 3 workspaces).');
      }
      throw e;
    }
    const workspaces = listContainerWorkspaces_(containerPath);
    const index = workspaces.findIndex(w => w.path === workspace.path);
    writeWorkspaceBlock_(WORKSPACE_SHEET.workspaces,
        workspaces.map(w => [w.name, w.path]), index);
    updateSelectedWorkspacePanel_();
    activateWorkspaceSheet_();
    return `Workspace "${name}" was created and selected.`;
  });
}

/**
 * Activates the GTM Workspace sheet.
 */
function activateWorkspaceSheet_() {
  getWorkspaceSheet_().activate();
}

/**
 * Returns the selected workspace row.
 * @return {?{name: string, path: string}}
 */
function getSelectedWorkspaceRow_() {
  const selected = readWorkspaceBlock_(WORKSPACE_SHEET.workspaces)
      .filter(row => row.selected && !isBlank(row.values[1]));
  if (!selected.length) return null;
  return {name: toText(selected[0].values[0]),
          path: toText(selected[0].values[1])};
}

/**
 * Returns the context of the selected workspace.
 * @return {!Object}
 */
function getWorkspaceContext() {
  const workspace = getSelectedWorkspaceRow_();
  if (!workspace) {
    throw new UserError('No workspace is selected. Go to the "GTM Workspace" ' +
        'sheet and select a workspace (menu: Select workspace).');
  }
  const parts = workspace.path.split('/');
  if (parts.length < 6) {
    throw new UserError(`The selected workspace path "${workspace.path}" is ` +
        'not valid. Please list the workspaces again.');
  }
  const accountPath = parts.slice(0, 2).join('/');
  const containerPath = parts.slice(0, 4).join('/');
  const info = getContainerInfo_(containerPath);
  const contexts = info.usageContext || [];
  const isGoogleTag = info.typeLabel === 'Google tag';
  const isServer = !isGoogleTag && contexts.indexOf('server') !== -1;
  const accountRow = readWorkspaceBlock_(WORKSPACE_SHEET.accounts)
      .find(row => toText(row.values[1]) === accountPath);
  return {
    accountPath: accountPath,
    containerPath: containerPath,
    workspacePath: workspace.path,
    accountId: parts[1],
    containerId: parts[3],
    workspaceId: parts[5],
    accountName: accountRow ? toText(accountRow.values[0]) : accountPath,
    containerName: info.name,
    workspaceName: workspace.name,
    publicId: info.publicId,
    usageContext: contexts,
    typeLabel: info.typeLabel,
    isServer: isServer,
    isWeb: !isServer && !isGoogleTag,
    isGoogleTag: isGoogleTag,
    features: info.features || {},
    timeZone: getSpreadsheetTimeZone_()
  };
}

/**
 * Returns the time zone of the spreadsheet (File › Settings), or the time zone
 * of the script if it is not available.
 * @return {string}
 */
function getSpreadsheetTimeZone_() {
  try {
    return SpreadsheetApp.getActive().getSpreadsheetTimeZone() ||
        Session.getScriptTimeZone();
  } catch (e) {
    return Session.getScriptTimeZone();
  }
}

/**
 * Updates the "Selected workspace" panel of the GTM Workspace sheet from the
 * data in the sheet (no API calls, so that it can run in onEdit).
 */
function updateSelectedWorkspacePanel_() {
  const sheet = getWorkspaceSheet_();
  const panel = WORKSPACE_SHEET.panel;
  const values = panel.labels.map(label => [label, '']);
  const workspace = getSelectedWorkspaceRow_();
  if (workspace) {
    const parts = workspace.path.split('/');
    const accountPath = parts.slice(0, 2).join('/');
    const containerPath = parts.slice(0, 4).join('/');
    const account = readWorkspaceBlock_(WORKSPACE_SHEET.accounts)
        .find(row => toText(row.values[1]) === accountPath);
    const container = readWorkspaceBlock_(WORKSPACE_SHEET.containers)
        .find(row => toText(row.values[3]) === containerPath);
    values[0][1] = account ? toText(account.values[0]) : accountPath;
    values[1][1] = container ? toText(container.values[0]) : containerPath;
    values[2][1] = container ? toText(container.values[1]) : '';
    values[3][1] = container ? toText(container.values[2]) : '';
    values[4][1] = workspace.name;
    values[5][1] = workspace.path;
    values[6][1] = parts.length >= 6 && parts[5] !== '0' ?
        `https://tagmanager.google.com/#/container/${workspace.path}` : '';
  }
  sheet.getRange(WORKSPACE_SHEET.firstDataRow, panel.column, values.length, 2)
      .setValues(values);
}

/**
 * Handles edits of the GTM Workspace sheet (called from onEdit): only one
 * workspace can be selected at a time, and the panel is updated.
 * @param {!Object} e The onEdit event.
 */
function handleWorkspaceSheetEdit_(e) {
  const range = e.range;
  const block = WORKSPACE_SHEET.workspaces;
  const selectColumn = block.column + block.headers.length - 1;
  if (range.getColumn() > selectColumn ||
      range.getColumn() + range.getNumColumns() - 1 < selectColumn ||
      range.getRow() + range.getNumRows() - 1 < WORKSPACE_SHEET.firstDataRow) {
    return;
  }
  if (range.getNumRows() === 1 && range.getNumColumns() === 1 &&
      toBool(range.getValue())) {
    const sheet = range.getSheet();
    const numRows = sheet.getLastRow() - WORKSPACE_SHEET.firstDataRow + 1;
    if (numRows > 0) {
      const selectRange = sheet.getRange(WORKSPACE_SHEET.firstDataRow,
          selectColumn, numRows, 1);
      const values = selectRange.getValues();
      let changed = false;
      values.forEach((row, index) => {
        if (WORKSPACE_SHEET.firstDataRow + index !== range.getRow() &&
            row[0] === true) {
          row[0] = false;
          changed = true;
        }
      });
      if (changed) selectRange.setValues(values);
    }
  }
  updateSelectedWorkspacePanel_();
}
