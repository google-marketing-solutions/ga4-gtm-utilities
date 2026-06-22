/**
 * Copyright 2023 Google LLC
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

const ui = SpreadsheetApp.getUi();
const ss = SpreadsheetApp.getActive();

const apiActionTaken = {
  ga4: {
    deleted: 'Deleted',
    created: 'Created',
    archived: 'Archived',
    skipped: 'Skipped',
    error: 'Error',
    updated: 'Updated'
  }
};

const GA4TagTypes = {
  EVENT: 'gaawe',
  CONFIG: 'gaawc'
}

/**
 * list GA4 event tags from a list of all tags in a workspace.
 * @return {!Array} List of GA4 event tags in a workspace.
 */
function listGA4EventTags(tags) {
  return tags.filter(tag => tag.type == GA4TagTypes.EVENT);
}

/**
 * Gets data to a specified sheet and range.
 * @param {string} sheetsMetaField The name of the sheet to read the data from.
 * @param {string} rangeName The name of the range to read the data from.
 * @return {!Array<Array>}
 */
function getDataFromSheet(sheetsMetaField, rangeName) {
  let ranges = null;
  sheetsMeta[sheetsMetaField].ranges.forEach(range => {
    if (range.name == rangeName) {
      ranges = range.read;
    }
  });
  let sheet = ss.getSheetByName(sheetsMeta[sheetsMetaField].sheetName);
  const lastRow = sheet.getLastRow();
  const numRows = lastRow - ranges.row + 1;
  if (numRows > 0) {
    const data = sheet.getRange(ranges.row, ranges.column, numRows, ranges.numColumns).getValues();
    return removeEmptyRows(data);
  }
  return [];
}

/**
 * Writes data to a specified sheet.
 * @param {!Array} data The data to be written to the sheet.
 * @param {string} sheetName The name of the sheet to which the data will be written.
 * @param {string} rangeName The name of the range for the sheet where the data will be written.
 */
function writeToSheet(data, sheetName, rangeName) {
  let ranges = null;
  sheetsMeta[sheetName].ranges.forEach(range => {
    if (range.name == rangeName) {
      ranges = range.write;
    }
  });
  let sheet = ss.getSheetByName(sheetsMeta[sheetName].sheetName);
  if (sheet == null) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
  }
  if (data.length > 0) {
    sheet.getRange(ranges.row, ranges.column, data.length, ranges.numColumns).setValues(data);
  }
}

/**
 * Removes the empty rows for a given range of values.
 * @param {!Array<!Array<string>>} rows Spreadsheet row values.
 * @return {?Array<?Array<string>>} Spreadsheet row values with no empty values.
 */
function removeEmptyRows(rows) {
  return rows.filter(row => {
    return row[0] != '';
  });
}

/**
 * Clear contents of a specific range.
 * @param {string} sheetsMetaField
 * @param {string} rangeName
 */
function clearRangeContent(sheetsMetaField, rangeName) {
  let ranges = null;
  sheetsMeta[sheetsMetaField].ranges.forEach(range => {
    if (range.name == rangeName) {
      ranges = range.read;
    }
  });
  const sheet = ss.getSheetByName(sheetsMeta[sheetsMetaField].sheetName);
  const lastRow = sheet.getLastRow();
  const numRows = lastRow - ranges.row + 1;
  if (numRows > 0) {
    sheet.getRange(ranges.row, ranges.column, numRows, ranges.numColumns).clearContent();
  }
}

/**
 * Builds the map object for fields, custom definitions, etc.
 * @param {string} name The name of the entity being mapped.
 * @param {string} value The value of the entity being mapped.
 * @return {!Object} The map object that can be added to a tag or variable.
 */
function buildParameterMapObject(name, value) {
  return {
    map: [
      {value: name, type: 'template', key: 'parameter'},
      {value: value.toString(), type: 'template', key: 'parameterValue'}
    ],
    type: 'map'
  };
}

/**
 * Builds the map object for user properties.
 * @param {string} name The name of the entity being mapped.
 * @param {string} value The value of the entity being mapped.
 * @return {!Object} The map object that can be added to a tag or variable.
 */
function buildUserPropertyMapObject (name, value) {
  return {
    map: [
      {value: name, type: 'template', key: 'name'},
      {value: value.toString(), type: 'template', key: 'value'}
    ],
    type: 'map'
  };
}

/**
 * Adds a row to the changelog sheet to create a record of the modification that
 * was made.
 * @param {string} entityName The name of what was changed.
 * @param {string} entityType The type (trigger, tag, variable, etc.) that was changed.
 * @param {number} entityId The ID of the entity that was chagned
 * @param {string} actionTaken A brief description of how something was changed.
 * @param {string} gtmURL The URL for the entity that was changed.
 */
function logChange(entityName, entityType, entityId, actionTaken, gtmURL) {
  const date = new Date();
  const sheet = ss.getSheetByName('Changelog');
  const currentDateTime = (date.getMonth() + 1) + '/' + date.getDate() + '/' + date.getFullYear();
  const user = Session.getActiveUser().getEmail();
  const loggedChange = [[
		currentDateTime, entityName, entityType, entityId, actionTaken, gtmURL, user
	]];
  sheet.getRange((sheet.getLastRow() + 1), 1, 1, 7).setValues(loggedChange);
}

/**
 * Writes tag names to the validation sheet.
 * @param {!Array} tags An array of tags in a workspace.
 */
function writeTagNamesToValidationSheet(tags) {
  const tagNames = tags.reduce((arr, tag) => {
    arr.push([tag.name]);
    return arr;
  }, []);
  writeToSheet(tagNames, 'validation', 'tag names')
}

/**
 * Writes event setting variable names to the validation sheet.
 * @param {!Array} variables An array of variables in a workspace.
 */
function writeSettingVariableNamesToValidationSheet(variables) {
  const variableNames = variables.reduce((arr, variable) => {
    if (variable.type == 'gtes') {
      arr.push([`{{${variable.name}}}`]);
    }
    return arr;
  }, []);
  writeToSheet(variableNames, 'validation', 'event setting variable names');
}

/**
 * Returns readable error response.
 * @param {!Object} response Error response object
 * @return {string}
 */
function errorResponse(response) {
  return `Error ${response.statusCode}: ' ${response.message}`;
}

/**
 * Checks if the current GTM container is a Server container.
 * @param {string} workspacePath The workspace path string.
 * @return {boolean} True if the container type is 'server'.
 */
function isServerContainer(workspacePath) {
  if (!workspacePath) return false;
  try {
    const parts = workspacePath.split('/');
    if (parts.length >= 4) {
      const containerPath = parts.slice(0, 4).join('/');
      const container = TagManager.Accounts.Containers.get(containerPath);
      return container && container.usageContext && container.usageContext.indexOf('server') !== -1;
    }
  } catch (e) {
    console.log('Error in isServerContainer: ' + e.message);
  }
  return false;
}