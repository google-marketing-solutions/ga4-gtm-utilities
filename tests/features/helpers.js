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
 * @fileoverview Helpers for the feature tests: a test environment with the
 * Apps Script runtime emulator, a fake GTM API and functions that simulate
 * what a user does in the spreadsheet (menu clicks, cell edits).
 */
'use strict';

const assert = require('node:assert/strict');
const {createRuntime} = require('../harness/appsScriptRuntime');
const {createFakeGtmApi} = require('../harness/fakeGtmApi');

/**
 * Builds a parameter list from a compact description.
 * @param {!Object<string, *>} params Values: string (template), boolean,
 *     array of strings (list), or array of [name, value] pairs with
 *     {table: [fieldA, fieldB]} given in tableFields.
 * @param {!Object<string, !Array<string>>=} tableFields
 * @return {!Array<!Object>}
 */
function params(params, tableFields) {
  return Object.keys(params).map(key => {
    const value = params[key];
    if (typeof value === 'boolean') {
      return {type: 'boolean', key, value: String(value)};
    }
    if (Array.isArray(value)) {
      const fields = (tableFields || {})[key];
      if (fields) {
        return {type: 'list', key, list: value.map(row => ({
          type: 'map',
          map: fields.map((field, i) => ({type: 'template', key: field,
                                          value: String(row[i])}))
        }))};
      }
      return {type: 'list', key, list: value.map(v => ({type: 'template',
                                                        value: String(v)}))};
    }
    return {type: 'template', key, value: String(value)};
  });
}

/**
 * Builds a Custom Event trigger.
 * @param {string} triggerId
 * @param {string} name
 * @param {string} eventName
 * @return {!Object}
 */
function customEventTrigger(triggerId, name, eventName) {
  return {
    triggerId, name, type: 'customEvent',
    customEventFilter: [{type: 'equals', parameter: [
      {type: 'template', key: 'arg0', value: '{{_event}}'},
      {type: 'template', key: 'arg1', value: eventName}]}]
  };
}

/**
 * A test environment: fake GTM API + Apps Script runtime with a selected
 * workspace.
 */
class TestEnv {
  /**
   * @param {{workspace: (!Object|undefined), fake: (!Object|undefined),
   *     runtime: (!Object|undefined), select: (boolean|undefined),
   *     preSeed: (!Object|undefined)}=} options
   */
  constructor(options = {}) {
    this.fake = createFakeGtmApi(options.fake || {});
    if (options.preSeed) this.fake.seed(options.preSeed);
    this.ws = this.fake.seedWorkspace(options.workspace || {});
    this.rt = createRuntime(Object.assign({fetchHandler: this.fake},
        options.runtime || {}));
    this.click('setupSheets');
    if (options.select !== false) this.selectWorkspace();
  }

  /**
   * Simulates a menu click (a new Apps Script execution).
   * @param {string} fnName
   * @param {string=} sheetName Sheet to activate first.
   * @return {!Object} The runtime.
   */
  click(fnName, sheetName) {
    this.rt = this.rt.nextExecution();
    this.rt.ui.reset();
    if (sheetName) {
      const ss = this.rt.spreadsheet.active;
      ss.setActiveSheet(ss.getSheetByName(sheetName));
    }
    this.rt.call(fnName);
    assert.deepEqual(this.rt.harnessErrors, [], 'harness errors');
    return this.rt;
  }

  /** Selects the seeded workspace through the menu. */
  selectWorkspace() {
    this.click('listGtmAccounts');
    this.click('listGtmContainers');
    this.click('listGtmWorkspaces');
  }

  /**
   * Lists a feature sheet.
   * @param {string} sheetName
   * @return {!TestEnv}
   */
  list(sheetName) {
    this.click('listActiveSheet', sheetName);
    this.assertNoAlerts();
    return this;
  }

  /**
   * Previews the changes of a sheet.
   * @param {string} sheetName
   * @return {!TestEnv}
   */
  preview(sheetName) {
    this.click('previewActiveSheet', sheetName);
    return this;
  }

  /**
   * Applies the changes of a sheet.
   * @param {string} sheetName
   * @return {!TestEnv}
   */
  apply(sheetName) {
    this.click('applyActiveSheet', sheetName);
    return this;
  }

  /** @return {!Array<string>} Messages of the alerts of the last click. */
  alertMessages() {
    return this.rt.ui.alerts.map(a => `${a.title || ''}: ${a.prompt || ''}`);
  }

  /** @return {!Array<string>} Messages of the toasts of the last click. */
  toastMessages() {
    return this.rt.ui.toasts.map(t => String(t.message));
  }

  /** Asserts that the last click showed no alert. */
  assertNoAlerts() {
    assert.deepEqual(this.alertMessages(), []);
  }

  /**
   * @param {string} sheetName
   * @return {!Array<!Array<*>>}
   */
  values(sheetName) {
    return this.rt.spreadsheet.getSheetValues(sheetName);
  }

  /**
   * Returns the 1-based column of a header (row 2).
   * @param {string} sheetName
   * @param {string} header
   * @return {number}
   */
  col(sheetName, header) {
    const headers = this.values(sheetName)[1] || [];
    const index = headers.indexOf(header);
    assert.notEqual(index, -1, `header "${header}" in ${sheetName}`);
    return index + 1;
  }

  /**
   * Returns the 1-based row whose column has a value (first match).
   * @param {string} sheetName
   * @param {string} header
   * @param {*} value
   * @param {function(!Array<*>): boolean=} filter Additional row filter.
   * @return {number}
   */
  row(sheetName, header, value, filter) {
    const col = this.col(sheetName, header) - 1;
    const values = this.values(sheetName);
    const index = values.findIndex((r, i) => i >= 2 && r[col] === value &&
        (!filter || filter(r)));
    assert.notEqual(index, -1, `row with ${header} = ${value} in ${sheetName}`);
    return index + 1;
  }

  /**
   * Returns a cell value.
   * @param {string} sheetName
   * @param {number} row
   * @param {string} header
   * @return {*}
   */
  get(sheetName, row, header) {
    return this.values(sheetName)[row - 1][this.col(sheetName, header) - 1];
  }

  /**
   * Sets cell values of a row, like a user editing the sheet.
   * @param {string} sheetName
   * @param {number} row
   * @param {!Object<string, *>} values Header -> value.
   * @return {!TestEnv}
   */
  set(sheetName, row, values) {
    const sheet = this.rt.spreadsheet.active.getSheetByName(sheetName);
    Object.keys(values).forEach(header => {
      sheet.getRange(row, this.col(sheetName, header)).setValue(values[header]);
    });
    return this;
  }

  /**
   * Copies a row to the first empty row (like copy & paste).
   * @param {string} sheetName
   * @param {number} row
   * @return {number} The new row.
   */
  copyRow(sheetName, row) {
    const sheet = this.rt.spreadsheet.active.getSheetByName(sheetName);
    const target = sheet.getLastRow() + 1;
    const width = sheet.getLastColumn();
    const values = sheet.getRange(row, 1, 1, width).getValues();
    sheet.getRange(target, 1, 1, width).setValues(values);
    return target;
  }

  /**
   * Adds a row with the given values.
   * @param {string} sheetName
   * @param {!Object<string, *>} values Header -> value.
   * @return {number} The new row.
   */
  addRow(sheetName, values) {
    const sheet = this.rt.spreadsheet.active.getSheetByName(sheetName);
    const row = Math.max(sheet.getLastRow() + 1, 3);
    this.set(sheetName, row, values);
    return row;
  }

  /**
   * Returns the Result cell of a row.
   * @param {string} sheetName
   * @param {number} row
   * @return {string}
   */
  result(sheetName, row) {
    return String(this.get(sheetName, row, 'Result'));
  }

  /**
   * Returns an entity of the fake API.
   * @param {string} kind
   * @param {string} id
   * @return {!Object|undefined}
   */
  entity(kind, id) {
    return this.fake.list(this.ws.workspacePath, kind)
        .find(e => String(e[{tag: 'tagId', trigger: 'triggerId',
          variable: 'variableId', folder: 'folderId', client: 'clientId',
          transformation: 'transformationId', gtagConfig: 'gtagConfigId'}[kind]]) ===
          String(id));
  }

  /**
   * Returns an entity of the fake API by name.
   * @param {string} kind
   * @param {string} name
   * @return {!Object|undefined}
   */
  entityByName(kind, name) {
    return this.fake.list(this.ws.workspacePath, kind).find(e => e.name === name);
  }
}

/**
 * Returns the value of a parameter of an entity.
 * @param {!Object} entity
 * @param {string} key
 * @return {*}
 */
function paramValue(entity, key) {
  const param = (entity.parameter || []).find(p => p.key === key);
  if (!param) return undefined;
  if (param.type === 'list') {
    return param.list.map(item => item.type === 'map' ?
        Object.fromEntries(item.map.map(e => [e.key, e.value])) : item.value);
  }
  return param.value;
}

module.exports = {TestEnv, params, paramValue, customEventTrigger};
