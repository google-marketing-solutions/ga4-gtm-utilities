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
 * @fileoverview Live tests against the real Tag Manager API.
 *
 * For every container in GTM_LIVE_CONTAINERS (comma-separated public IDs),
 * a temporary workspace is created, the tool's code is run against it (list,
 * create, update, delete on test items named "[TEST] ...") and the workspace
 * is deleted at the end. Nothing is published and the other workspaces of
 * the container are not changed.
 *
 * Requires a token (see liveAuth.js). Skipped otherwise.
 *   GTM_LIVE_CONTAINERS=GTM-AAAA,GTM-BBBB node tests/live/live.test.js
 */
'use strict';

const {test, describe, after} = require('node:test');
const assert = require('node:assert/strict');
const {execFileSync} = require('node:child_process');
const {createRuntime} = require('../harness/appsScriptRuntime');
const {getLiveAccessToken, getLiveExtraHeaders} = require('./liveAuth');

const containers = (process.env.GTM_LIVE_CONTAINERS || '').split(',')
    .map(id => id.trim()).filter(Boolean);
let token = null;
try {
  token = containers.length ? getLiveAccessToken() : null;
} catch (e) {
  console.log('Live tests skipped: ' + e.message);
}
const skip = !token || !containers.length;
const BASE = 'https://tagmanager.googleapis.com/tagmanager/v2/';
const RUN_ID = new Date().toISOString().slice(0, 16).replace(/[-:T]/g, '');

/**
 * Calls the API directly (for setup, verification and clean-up).
 * @param {string} method
 * @param {string} path
 * @param {!Object=} body
 * @return {{code: number, body: *}}
 */
function api(method, path, body) {
  const headers = Object.assign({Authorization: 'Bearer ' + getLiveAccessToken(),
                                 'Content-Type': 'application/json'},
                                getLiveExtraHeaders());
  const config = [`url = "${BASE}${path}"`, `request = "${method}"`,
                  'silent', 'show-error', 'write-out = "\\n%{http_code}"']
      .concat(Object.keys(headers).map(k => `header = "${k}: ${headers[k]}"`));
  if (body) config.push('data-binary = "@-"');
  const args = ['-K', '-'];
  const input = config.join('\n') + '\n';
  // The payload is passed through a second stdin read is not possible with
  // -K, so write it to a temporary file when needed.
  let output;
  if (body) {
    const fs = require('node:fs');
    const os = require('node:os');
    const path_ = require('node:path');
    const file = path_.join(os.tmpdir(), `gtm-live-${process.pid}.json`);
    fs.writeFileSync(file, JSON.stringify(body));
    output = execFileSync('curl', args, {input: input.replace('"@-"',
        `"@${file}"`), encoding: 'utf8'});
    fs.unlinkSync(file);
  } else {
    output = execFileSync('curl', args, {input, encoding: 'utf8'});
  }
  const index = output.lastIndexOf('\n');
  const code = Number(output.slice(index + 1));
  let parsed = output.slice(0, index);
  try {
    parsed = parsed ? JSON.parse(parsed) : {};
  } catch (e) {
    // Keep the text.
  }
  return {code, body: parsed};
}

/**
 * Lists the entities of a kind in a workspace.
 * @param {string} workspacePath
 * @param {string} collection
 * @param {string} key
 * @return {!Array<!Object>}
 */
function listEntities(workspacePath, collection, key) {
  let items = [];
  let pageToken = '';
  do {
    const response = api('GET', `${workspacePath}/${collection}` +
        (pageToken ? `?pageToken=${encodeURIComponent(pageToken)}` : ''));
    assert.equal(response.code, 200, JSON.stringify(response.body));
    items = items.concat(response.body[key] || []);
    pageToken = response.body.nextPageToken;
  } while (pageToken);
  return items;
}

/**
 * Drives the tool in a temporary workspace of a container.
 */
class LiveEnv {
  /** @param {string} publicId */
  constructor(publicId) {
    this.publicId = publicId;
    const lookup = api('GET', `accounts/containers:lookup?tagId=${publicId}`);
    assert.equal(lookup.code, 200, JSON.stringify(lookup.body));
    this.container = lookup.body;
    const workspace = api('POST', `${this.container.path}/workspaces`, {
      name: `[TEST] ga4-gtm-utilities ${RUN_ID}`,
      description: 'Temporary workspace of the GA4 GTM Utilities live tests. ' +
          'It is deleted automatically.'
    });
    assert.equal(workspace.code, 200,
        `Could not create a test workspace in ${publicId}: ` +
        JSON.stringify(workspace.body));
    this.workspace = workspace.body;
    this.rt = createRuntime({
      realHttp: true,
      oauthToken: () => getLiveAccessToken(),
      realHttpHeaders: getLiveExtraHeaders(),
      realSleep: true,
      virtualTime: false
    });
    this.click('setupSheets');
    // Select the temporary workspace: find the container by ID, then tick it.
    this.rt = this.rt.nextExecution();
    this.rt.ui.queueResponse(publicId);
    this.rt.call('findContainerById');
    const sheet = this.rt.spreadsheet.active.getSheetByName('GTM Workspace');
    const values = this.rt.spreadsheet.getSheetValues('GTM Workspace');
    values.forEach((row, i) => {
      if (i >= 2 && row[9]) {
        sheet.getRange(i + 1, 11).setValue(row[9] === this.workspace.path);
      }
    });
    const ctx = this.rt.call('getWorkspaceContext');
    assert.equal(ctx.workspacePath, this.workspace.path);
  }

  click(fnName, sheetName) {
    this.rt = this.rt.nextExecution();
    this.rt.ui.reset();
    if (sheetName) {
      const ss = this.rt.spreadsheet.active;
      ss.setActiveSheet(ss.getSheetByName(sheetName));
    }
    this.rt.call(fnName);
    return this.rt;
  }

  alerts() {
    return this.rt.ui.alerts.map(a => `${a.title}: ${a.prompt}`);
  }

  list(sheetName) {
    this.click('listActiveSheet', sheetName);
    assert.deepEqual(this.alerts(), [], `alerts when listing ${sheetName}`);
  }

  apply(sheetName) {
    this.click('applyActiveSheet', sheetName);
  }

  values(sheetName) {
    return this.rt.spreadsheet.getSheetValues(sheetName);
  }

  col(sheetName, header) {
    const index = this.values(sheetName)[1].indexOf(header);
    assert.notEqual(index, -1, header);
    return index + 1;
  }

  row(sheetName, header, value) {
    const col = this.col(sheetName, header) - 1;
    const index = this.values(sheetName).findIndex((r, i) => i >= 2 &&
        r[col] === value);
    assert.notEqual(index, -1, `${header} = ${value} in ${sheetName}`);
    return index + 1;
  }

  set(sheetName, row, values) {
    const sheet = this.rt.spreadsheet.active.getSheetByName(sheetName);
    Object.keys(values).forEach(header =>
      sheet.getRange(row, this.col(sheetName, header)).setValue(values[header]));
  }

  addRow(sheetName, values) {
    const sheet = this.rt.spreadsheet.active.getSheetByName(sheetName);
    const row = Math.max(sheet.getLastRow() + 1, 3);
    this.set(sheetName, row, values);
    return row;
  }

  result(sheetName, row) {
    return String(this.values(sheetName)[row - 1][1]);
  }

  assertApplied(sheetName, rows) {
    rows.forEach(row => assert.match(this.result(sheetName, row), /^✔/,
        `${sheetName} row ${row}: ${this.result(sheetName, row)}`));
  }

  entities(collection, key) {
    return listEntities(this.workspace.path, collection, key);
  }

  cleanUp() {
    if (!this.workspace) return;
    const response = api('DELETE', this.workspace.path);
    console.log(`Deleted test workspace ${this.workspace.path}: ${response.code}`);
  }
}

const envs = [];
after(() => envs.forEach(env => env.cleanUp()));

const param = (entity, key) => {
  const p = (entity.parameter || []).find(x => x.key === key);
  return p ? (p.value !== undefined ? p.value : p) : undefined;
};

describe('live GTM API', {skip}, () => {
  for (const publicId of containers) {
    test(`container ${publicId}`, {timeout: 40 * 60 * 1000}, () => {
      const env = new LiveEnv(publicId);
      envs.push(env);
      const ctx = env.rt.call('getWorkspaceContext');
      console.log(`${publicId}: ${ctx.containerName} (${ctx.typeLabel})`);
      // Every sheet of the container type can be listed.
      env.rt.call('getSheetDefs').forEach(def => {
        const contexts = def.contexts;
        const type = ctx.isServer ? 'server' : 'web';
        if (contexts.includes(type)) env.list(def.sheetName);
      });
      if (ctx.isServer) {
        runServerScenario(env);
      } else {
        runWebScenario(env);
      }
    });
  }
});

/**
 * Create / update / delete scenario for web containers.
 * @param {!LiveEnv} env
 */
function runWebScenario(env) {
  // Triggers and variables.
  env.list('Triggers');
  const trigger = env.addRow('Triggers', {Action: 'Create',
    'Trigger Name': '[TEST] CE - test_event', 'Trigger Type': 'Custom Event',
    'Event Name (Custom Event)': 'test_event',
    'Conditions (fires on)': '{{Page Path}} does not contain /admin',
    Folder: '[TEST] Folder'});
  env.apply('Triggers');
  env.assertApplied('Triggers', [trigger]);
  env.list('Variables');
  const constant = env.addRow('Variables', {Action: 'Create',
    'Variable Name': '[TEST] GA4 ID', 'Variable Type': 'Constant',
    'Value / Main Setting': 'G-TEST000001'});
  const dlv = env.addRow('Variables', {Action: 'Create',
    'Variable Name': '[TEST] DLV - method', 'Variable Type': 'Data Layer Variable',
    'Value / Main Setting': 'method', 'Default Value': 'unknown'});
  const lookup = env.addRow('Variables', {Action: 'Create',
    'Variable Name': '[TEST] Lookup', 'Variable Type': 'Lookup Table',
    'Value / Main Setting': '{{Page Hostname}}',
    'Lookup Table (input = output)': 'a.com = 1\nb.com = 2'});
  env.apply('Variables');
  env.assertApplied('Variables', [constant, dlv, lookup]);

  // GA4 event tag.
  env.list('GA4 Event Tags');
  const ga4 = env.addRow('GA4 Event Tags', {Action: 'Create',
    'Tag Name': '[TEST] GA4 - test_event', 'Event Name': 'test_event',
    'Measurement ID': '{{[TEST] GA4 ID}}',
    'Event Parameters': 'method = {{[TEST] DLV - method}}\nsource = live_test',
    'User Properties': 'tester = yes',
    'Firing Triggers': '[TEST] CE - test_event', Folder: '[TEST] Folder'});
  env.apply('GA4 Event Tags');
  env.assertApplied('GA4 Event Tags', [ga4]);
  let tag = env.entities('tags', 'tag').find(t => t.name === '[TEST] GA4 - test_event');
  assert.ok(tag, 'GA4 event tag created');
  assert.equal(param(tag, 'eventName'), 'test_event');

  // Event parameters: add, update, delete.
  const sheet = 'Event Parameters & User Properties';
  env.list(sheet);
  const add = env.addRow(sheet, {Action: 'Create',
    'Tag / Variable Name': '[TEST] GA4 - test_event', Setting: 'Event Parameter',
    'Parameter / Property Name': 'added', Value: '1'});
  const update = env.row(sheet, 'Parameter / Property Name', 'source');
  env.set(sheet, update, {Value: 'live_test_updated'});
  const remove = env.row(sheet, 'Parameter / Property Name', 'tester');
  env.set(sheet, remove, {Action: 'Delete'});
  env.apply(sheet);
  env.assertApplied(sheet, [add, update, remove]);
  tag = env.entities('tags', 'tag').find(t => t.name === '[TEST] GA4 - test_event');
  const table = (tag.parameter.find(p => p.key === 'eventSettingsTable') || {}).list || [];
  const rows = table.map(item => Object.fromEntries(item.map.map(e => [e.key, e.value])));
  assert.deepEqual(rows.map(r => [r.parameter, r.parameterValue]), [
    ['method', '{{[TEST] DLV - method}}'], ['source', 'live_test_updated'],
    ['added', '1']]);
  assert.equal(tag.parameter.find(p => p.key === 'userProperties'), undefined);

  // Tags sheet: common settings, copy and delete.
  env.list('Tags');
  const tagRow = env.row('Tags', 'Tag Name', '[TEST] GA4 - test_event');
  env.set('Tags', tagRow, {Paused: true, 'Required Consent Types': 'analytics_storage',
                           'Tag Firing Priority': '5', Notes: 'Live test'});
  env.apply('Tags');
  env.assertApplied('Tags', [tagRow]);
  env.list('Tags');
  const source = env.row('Tags', 'Tag Name', '[TEST] GA4 - test_event');
  const sheetObj = env.rt.spreadsheet.active.getSheetByName('Tags');
  const width = sheetObj.getLastColumn();
  const copy = sheetObj.getLastRow() + 1;
  sheetObj.getRange(copy, 1, 1, width).setValues(
      sheetObj.getRange(source, 1, 1, width).getValues());
  env.set('Tags', copy, {'Tag Name': '[TEST] GA4 - copy', Action: 'Create'});
  env.apply('Tags');
  env.assertApplied('Tags', [copy]);
  tag = env.entities('tags', 'tag').find(t => t.name === '[TEST] GA4 - test_event');
  assert.equal(tag.paused, true);
  assert.equal(tag.consentSettings.consentStatus, 'needed');
  const copied = env.entities('tags', 'tag').find(t => t.name === '[TEST] GA4 - copy');
  assert.ok(copied, 'copy created');
  assert.equal(param(copied, 'eventName'), 'test_event');

  // Google Ads and Floodlight tags.
  const ads = 'Google Ads & Floodlight Tags';
  env.list(ads);
  const awct = env.addRow(ads, {Action: 'Create', 'Tag Name': '[TEST] Ads conversion',
    'Tag Type': 'Google Ads Conversion Tracking', 'Conversion ID': '123456789',
    'Conversion Label': 'TestLabel', 'Conversion Value': '1',
    'Currency Code': 'EUR', 'Firing Triggers': '[TEST] CE - test_event'});
  const flc = env.addRow(ads, {Action: 'Create', 'Tag Name': '[TEST] Floodlight',
    'Tag Type': 'Floodlight Counter', 'Advertiser ID': '1234567',
    'Group Tag String (type=)': 'testg0', 'Activity Tag String (cat=)': 'testa0',
    'Counting Method': 'Standard', 'Custom Variables': 'u1 = test',
    'Firing Triggers': '[TEST] CE - test_event'});
  const linker = env.addRow(ads, {Action: 'Create', 'Tag Name': '[TEST] Linker',
    'Tag Type': 'Conversion Linker', 'Enable Linking on All Page URLs': true,
    'Auto-Link Domains (Cross-Domain)': 'example.com, example.org',
    'Firing Triggers': 'All Pages'});
  env.apply(ads);
  env.assertApplied(ads, [awct, flc, linker]);

  // Google tag / destination tags.
  env.list('Google Tags');
  const dest = env.addRow('Google Tags', {Action: 'Create',
    'Tag Name': '[TEST] GA destination', 'Tag Type': 'Google Analytics Destination',
    'Tag / Destination ID': 'G-TEST000002',
    'Configuration Parameters': 'send_page_view = false',
    'Firing Triggers': 'Initialization - All Pages'});
  env.apply('Google Tags');
  console.log('Google Tags result:', env.result('Google Tags', dest));

  // Advanced settings: edit a parameter of the copy.
  env.list('Advanced Settings');
  const adv = env.values('Advanced Settings').findIndex((r, i) => i >= 2 &&
      r[3] === '[TEST] GA4 - copy' && r[5] === 'eventName') + 1;
  assert.ok(adv > 0);
  env.set('Advanced Settings', adv, {Value: 'copied_event'});
  env.apply('Advanced Settings');
  env.assertApplied('Advanced Settings', [adv]);

  // Google tag settings (container-wide).
  env.list('Google Tag Settings');
  console.log('Google Tag Settings:', env.values('Google Tag Settings').length - 2,
      'rows;', String(env.values('Google Tag Settings')[0][0]));

  // Delete the test items again (in dependency order).
  env.list('Tags');
  ['[TEST] GA4 - copy', '[TEST] GA4 - test_event', '[TEST] Ads conversion',
   '[TEST] Floodlight', '[TEST] Linker'].forEach(name =>
    env.set('Tags', env.row('Tags', 'Tag Name', name), {Action: 'Delete'}));
  env.apply('Tags');
  env.list('Variables');
  const used = env.row('Variables', 'Variable Name', '[TEST] GA4 ID');
  env.set('Variables', used, {Action: 'Delete'});
  env.apply('Variables');
  env.assertApplied('Variables', [used]);
  env.list('Triggers');
  const t = env.row('Triggers', 'Trigger Name', '[TEST] CE - test_event');
  env.set('Triggers', t, {Action: 'Delete'});
  env.apply('Triggers');
  env.assertApplied('Triggers', [t]);
}

/**
 * Create / update / delete scenario for server containers.
 * @param {!LiveEnv} env
 */
function runServerScenario(env) {
  env.list('Triggers');
  const trigger = env.addRow('Triggers', {Action: 'Create',
    'Trigger Name': '[TEST] Event - test_event', 'Trigger Type': 'Custom (Server)',
    'Conditions (fires on)': '{{Event Name}} equals test_event'});
  env.apply('Triggers');
  env.assertApplied('Triggers', [trigger]);
  env.list('Variables');
  const ed = env.addRow('Variables', {Action: 'Create',
    'Variable Name': '[TEST] ED - value', 'Variable Type': 'Event Data',
    'Value / Main Setting': 'value'});
  env.apply('Variables');
  env.assertApplied('Variables', [ed]);

  env.list('Server GA4 Tags');
  const ga4 = env.addRow('Server GA4 Tags', {Action: 'Create',
    'Tag Name': '[TEST] GA4 server', 'Measurement ID': 'G-TEST000003',
    'Parameters to Exclude': 'ip_override',
    'Parameters to Add / Edit': 'source = live_test',
    'Firing Triggers': '[TEST] Event - test_event'});
  env.apply('Server GA4 Tags');
  env.assertApplied('Server GA4 Tags', [ga4]);

  const ads = 'Server Ads & Floodlight Tags';
  env.list(ads);
  const sgtmAds = env.addRow(ads, {Action: 'Create', 'Tag Name': '[TEST] Ads server',
    'Tag Type': 'Google Ads Conversion Tracking (Server)',
    'Conversion ID': '123456789', 'Conversion Label': 'TestLabel',
    'Firing Triggers': '[TEST] Event - test_event'});
  const sgtmFl = env.addRow(ads, {Action: 'Create', 'Tag Name': '[TEST] FL server',
    'Tag Type': 'Floodlight Counter (Server)', 'Advertiser ID': '1234567',
    'Group Tag String (type=)': 'testg0', 'Activity Tag String (cat=)': 'testa0',
    'Counting Method': 'Standard', 'Firing Triggers': '[TEST] Event - test_event'});
  env.apply(ads);
  env.assertApplied(ads, [sgtmAds, sgtmFl]);

  env.list('Clients');
  const client = env.addRow('Clients', {Action: 'Create',
    'Client Name': '[TEST] GTM client', 'Client Type': 'Google Tag Manager: Web Container',
    'Allowed Container IDs (ID = path)': 'GTM-TEST001 = /test', Priority: '1'});
  env.apply('Clients');
  env.assertApplied('Clients', [client]);

  env.list('Transformations');
  const transformation = env.addRow('Transformations', {Action: 'Create',
    'Transformation Name': '[TEST] Exclude', 'Transformation Type': 'Exclude parameters',
    'Excluded Parameters': 'ip_override', 'Affected Tags Mode': 'Only these tags',
    'Affected Tags': '[TEST] GA4 server'});
  env.apply('Transformations');
  env.assertApplied('Transformations', [transformation]);

  const sheet = 'Event Parameters & User Properties';
  env.list(sheet);
  const add = env.addRow(sheet, {Action: 'Create',
    'Tag / Variable Name': '[TEST] GA4 server', Setting: 'User Property to Add / Edit',
    'Parameter / Property Name': 'tester', Value: 'yes'});
  env.apply(sheet);
  env.assertApplied(sheet, [add]);

  env.list('Transformations');
  const tRow = env.row('Transformations', 'Transformation Name', '[TEST] Exclude');
  env.set('Transformations', tRow, {Action: 'Delete'});
  env.apply('Transformations');
  env.assertApplied('Transformations', [tRow]);
}
