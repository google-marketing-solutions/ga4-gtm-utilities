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
 * @fileoverview End-to-end tests of the feature sheets with a web container:
 * menu clicks and cell edits against the fake GTM API.
 */
'use strict';

const {test, describe} = require('node:test');
const assert = require('node:assert/strict');
const {TestEnv, params, paramValue, customEventTrigger} = require('./helpers');

/** A web workspace with typical content. */
function webWorkspace() {
  return {
    containerName: 'Shop Web', publicId: 'GTM-WEB0001',
    folders: [{folderId: '90', name: 'GA4'}],
    triggers: [
      customEventTrigger('10', 'CE - purchase', 'purchase'),
      {triggerId: '11', name: 'PV - Checkout', type: 'pageview', filter: [{
        type: 'contains', parameter: [
          {type: 'template', key: 'arg0', value: '{{Page Path}}'},
          {type: 'template', key: 'arg1', value: '/checkout'}]}]},
      customEventTrigger('12', 'CE - unused', 'unused_event')
    ],
    variables: [
      {variableId: '20', name: 'GA4 ID', type: 'c',
       parameter: params({value: 'G-OLD0000001'})},
      {variableId: '21', name: 'DLV - value', type: 'v',
       parameter: params({name: 'ecommerce.value'}).concat(
           [{type: 'integer', key: 'dataLayerVersion', value: '2'}])},
      {variableId: '22', name: 'Unused constant', type: 'c',
       parameter: params({value: 'x'})},
      {variableId: '23', name: 'Event Settings', type: 'gtes',
       parameter: params({eventSettingsTable: [['site', 'shop']]},
           {eventSettingsTable: ['parameter', 'parameterValue']})}
    ],
    tags: [
      {tagId: '1', name: 'GA4 - purchase', type: 'gaawe', firingTriggerId: ['10'],
       parentFolderId: '90',
       parameter: params({
         eventName: 'purchase', measurementIdOverride: '{{GA4 ID}}',
         sendEcommerceData: true, getEcommerceDataFrom: 'dataLayer',
         eventSettingsTable: [['value', '{{DLV - value}}'],
                              ['currency', 'EUR']],
         userProperties: [['customer_type', 'new']]
       }, {eventSettingsTable: ['parameter', 'parameterValue'],
           userProperties: ['name', 'value']})},
      {tagId: '2', name: 'GA4 - page_view', type: 'gaawe',
       firingTriggerId: ['2147479553'],
       parameter: params({eventName: 'page_view',
                          measurementIdOverride: 'G-ABC1234567'})},
      {tagId: '3', name: 'Custom HTML - chat', type: 'html',
       firingTriggerId: ['11'],
       parameter: params({html: '<script>chat()</script>',
                          supportDocumentWrite: false})},
      {tagId: '4', name: 'Google Tag', type: 'googtag',
       firingTriggerId: ['2147479573'],
       parameter: params({tagId: 'G-ABC1234567',
                          configSettingsTable: [['send_page_view', 'false']]},
           {configSettingsTable: ['parameter', 'parameterValue']})},
      {tagId: '5', name: 'Ads - Purchase', type: 'awct', firingTriggerId: ['10'],
       parameter: params({conversionId: '123456789', conversionLabel: 'abcDEF',
                          conversionValue: '{{DLV - value}}',
                          currencyCode: 'EUR'})},
      {tagId: '6', name: 'Floodlight - Purchase', type: 'flc',
       firingTriggerId: ['10'],
       parameter: params({advertiserId: '1234567', groupTag: 'sales',
                          activityTag: 'purch0', ordinalType: 'STANDARD',
                          customVariable: [['u1', '{{DLV - value}}']]},
           {customVariable: ['key', 'value']})},
      {tagId: '7', name: 'Meta Pixel', type: 'cvt_9000001_70',
       firingTriggerId: ['2147479553'],
       parameter: params({pixelId: '111', eventName: 'PageView'})}
    ],
    templates: [{templateId: '70', name: 'Meta Pixel (template)',
                 templateData: '___INFO___'}],
    gtagConfigs: [{gtagConfigId: '80', type: 'ogt_auto_events',
                   parameter: params({enablePageView: true,
                                      enableScroll: false})}]
  };
}

describe('workspace selection and setup', () => {
  test('all sheets are created and the workspace is selected', () => {
    const env = new TestEnv({workspace: webWorkspace()});
    const names = env.rt.spreadsheet.sheetNames();
    ['GTM Workspace', 'Tags', 'GA4 Event Tags',
     'Event Parameters & User Properties', 'Google Tags',
     'Google Ads & Floodlight Tags', 'Server GA4 Tags',
     'Server Ads & Floodlight Tags', 'Clients', 'Transformations',
     'Variables', 'Triggers', 'Advanced Settings', 'Google Tag Settings',
     'Changelog', 'Lists', 'Settings'].forEach(name =>
      assert.ok(names.includes(name), name));
    assert.ok(!names.includes('Sheet1'));
    const panel = env.values('GTM Workspace').slice(2, 9).map(r => r.slice(12, 14));
    assert.deepEqual(panel[1], ['Container', 'Shop Web']);
    assert.deepEqual(panel[2], ['Container ID', 'GTM-WEB0001']);
    assert.deepEqual(panel[3], ['Container type', 'Web']);
    assert.deepEqual(panel[4], ['Workspace', 'Default Workspace']);
    const ctx = env.rt.call('getWorkspaceContext');
    assert.equal(ctx.isWeb, true);
    assert.equal(ctx.publicId, 'GTM-WEB0001');
  });

  test('menu items point to existing functions', () => {
    const env = new TestEnv({workspace: webWorkspace(), select: false});
    env.rt.callTrigger('onOpen');
    assert.deepEqual(env.rt.missingMenuFunctions(), []);
    assert.ok(env.rt.ui.menuFunctionNames().includes('applyActiveSheet'));
  });

  test('find container by ID selects the container and workspace', () => {
    const env = new TestEnv({workspace: webWorkspace(), select: false});
    env.rt = env.rt.nextExecution();
    env.rt.ui.queueResponse('GTM-WEB0001');
    env.rt.call('findContainerById');
    assert.equal(env.rt.call('getWorkspaceContext').containerName, 'Shop Web');
  });

  test('only one workspace can be selected (onEdit)', () => {
    const env = new TestEnv({workspace: webWorkspace(), select: false});
    env.fake.seed({workspaces: [{accountId: '6000001', containerId: '9000001',
                                 workspaceId: '2', name: 'Second'}]});
    env.click('listGtmAccounts');
    env.click('listGtmContainers');
    env.click('listGtmWorkspaces');
    const sheet = env.rt.spreadsheet.active.getSheetByName('GTM Workspace');
    sheet.getRange(3, 11).setValue(true);
    env.rt.callTrigger('onEdit', {range: sheet.getRange(3, 11)});
    sheet.getRange(4, 11).setValue(true);
    env.rt.callTrigger('onEdit', {range: sheet.getRange(4, 11)});
    const selected = env.values('GTM Workspace').slice(2, 4).map(r => r[10]);
    assert.deepEqual(selected, [false, true]);
    assert.equal(env.values('GTM Workspace')[6][13], 'Second');
  });

  test('feature sheets require a selected workspace', () => {
    const env = new TestEnv({workspace: webWorkspace(), select: false});
    env.click('listActiveSheet', 'Tags');
    assert.match(env.alertMessages()[0], /No workspace is selected/);
  });

  test('setup removes conditional formats of older versions', () => {
    const env = new TestEnv({workspace: webWorkspace(), select: false});
    env.rt = env.rt.nextExecution();
    // A rule of version 2 that hides text in columns G:Z.
    env.rt.run(`(() => {
      const sheet = SpreadsheetApp.getActive().getSheetByName('GTM Workspace');
      sheet.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied('= $I2<> false')
          .setFontColor('#ffffff').setBackground('#ffffff')
          .setRanges([sheet.getRange('G2:Z1000')]).build()]);
    })()`);
    env.rt.call('setupSheets');
    const rules = JSON.parse(env.rt.run(`JSON.stringify(SpreadsheetApp
        .getActive().getSheetByName('GTM Workspace').getConditionalFormatRules()
        .map(rule => rule.describe()))`));
    assert.deepEqual(rules.map(rule => [rule.values[0], rule.ranges[0],
                                        rule.format.background]), [
      ['=$C3=TRUE', 'A3:C1000', '#ceead6'],
      ['=$H3=TRUE', 'D3:H1000', '#ceead6'],
      ['=$K3=TRUE', 'I3:K1000', '#ceead6']]);
    const tagRules = JSON.parse(env.rt.run(`JSON.stringify(SpreadsheetApp
        .getActive().getSheetByName('Tags').getConditionalFormatRules()
        .map(rule => rule.describe()))`));
    assert.deepEqual(tagRules.map(rule => rule.values[0]),
        ['Create', 'Update', 'Delete']);
  });
});

describe('Tags sheet', () => {
  test('lists all tags with readable values', () => {
    const env = new TestEnv({workspace: webWorkspace()}).list('Tags');
    const row = env.row('Tags', 'Tag Name', 'GA4 - purchase');
    assert.equal(env.get('Tags', row, 'Tag Type'), 'GA4 Event');
    assert.equal(env.get('Tags', row, 'Folder'), 'GA4');
    assert.equal(env.get('Tags', row, 'Firing Triggers'), 'CE - purchase');
    assert.equal(env.get('Tags', row, 'Additional Consent Checks'), 'Not set');
    const meta = env.row('Tags', 'Tag Name', 'Meta Pixel');
    assert.equal(env.get('Tags', meta, 'Tag Type'), 'Meta Pixel (template)');
    assert.equal(env.values('Tags').length, 2 + 7);
    assert.match(String(env.values('Tags')[0][0]), /7 listed from GTM-WEB0001/);
  });

  test('edited rows are detected, previewed and applied', () => {
    const env = new TestEnv({workspace: webWorkspace()}).list('Tags');
    const row = env.row('Tags', 'Tag Name', 'Custom HTML - chat');
    env.set('Tags', row, {
      'Tag Name': 'Custom HTML - chat widget', 'Paused': true,
      'Firing Triggers': 'All Pages, PV - Checkout',
      'Exception (Blocking) Triggers': 'CE - purchase',
      'Required Consent Types': 'functionality_storage',
      'Tag Firing Option': 'Once per page', 'Tag Firing Priority': '10',
      'Setup Tag (fires before)': 'GA4 - page_view',
      "Don't Fire if Setup Tag Fails": true,
      'Schedule Start': '2030-01-02 09:30',
      'Only Fire in Published Containers': true,
      'Notes': 'Reviewed'
    });
    env.preview('Tags');
    assert.match(env.result('Tags', row), /^Edited, will be updated: /);
    assert.equal(env.entity('tag', '3').name, 'Custom HTML - chat');
    env.apply('Tags');
    assert.match(env.result('Tags', row), /^✔ Updated/);
    assert.equal(env.get('Tags', row, 'Action'), '');
    const tag = env.entity('tag', '3');
    assert.equal(tag.name, 'Custom HTML - chat widget');
    assert.equal(tag.paused, true);
    assert.deepEqual(tag.firingTriggerId, ['2147479553', '11']);
    assert.deepEqual(tag.blockingTriggerId, ['10']);
    assert.deepEqual(tag.consentSettings, {consentStatus: 'needed',
      consentType: {type: 'list', list: [{type: 'template',
                                          value: 'functionality_storage'}]}});
    assert.equal(tag.tagFiringOption, 'oncePerLoad');
    assert.deepEqual(tag.priority, {type: 'integer', value: '10'});
    assert.deepEqual(tag.setupTag, [{tagName: 'GA4 - page_view',
                                     stopOnSetupFailure: true}]);
    assert.ok(Number(tag.scheduleStartMs) > 0);
    assert.equal(tag.liveOnly, true);
    assert.equal(tag.notes, 'Reviewed');
    // Untouched settings are kept.
    assert.equal(paramValue(tag, 'html'), '<script>chat()</script>');
    // The changelog has the change.
    const log = env.values('Changelog');
    assert.equal(log.length, 2);
    assert.equal(log[1][5], 'Custom HTML - chat widget');
    assert.equal(log[1][7], 'Updated');
    // Nothing is left to apply.
    env.preview('Tags');
    assert.match(env.toastMessages()[0], /No changes found/);
  });

  test('a copied row with Action Create duplicates the tag', () => {
    const env = new TestEnv({workspace: webWorkspace()}).list('Tags');
    const source = env.row('Tags', 'Tag Name', 'Meta Pixel');
    const copy = env.copyRow('Tags', source);
    env.set('Tags', copy, {'Tag Name': 'Meta Pixel - copy', 'Action': 'Create',
                           'Paused': true});
    env.apply('Tags');
    assert.match(env.result('Tags', copy), /^✔ Created \(ID \d+\)/);
    const created = env.entityByName('tag', 'Meta Pixel - copy');
    assert.equal(created.type, 'cvt_9000001_70');
    assert.equal(paramValue(created, 'pixelId'), '111');
    assert.equal(created.paused, true);
    assert.equal(String(env.get('Tags', copy, 'Tag ID')), created.tagId);
    // The source is unchanged.
    assert.equal(env.entity('tag', '7').name, 'Meta Pixel');
  });

  test('a copied row without Action is rejected as a duplicate', () => {
    const env = new TestEnv({workspace: webWorkspace()}).list('Tags');
    const source = env.row('Tags', 'Tag Name', 'Meta Pixel');
    const copy = env.copyRow('Tags', source);
    env.set('Tags', copy, {'Tag Name': 'Meta Pixel - renamed?'});
    env.apply('Tags');
    assert.match(env.result('Tags', copy), /refer to the same item/);
    assert.equal(env.entity('tag', '7').name, 'Meta Pixel');
  });

  test('tags are deleted', () => {
    const env = new TestEnv({workspace: webWorkspace()}).list('Tags');
    const row = env.row('Tags', 'Tag Name', 'Meta Pixel');
    env.set('Tags', row, {Action: 'Delete'});
    env.preview('Tags');
    assert.equal(env.result('Tags', row), 'Will be deleted.');
    env.apply('Tags');
    assert.equal(env.result('Tags', row), '✔ Deleted.');
    assert.equal(env.entity('tag', '7'), undefined);
  });

  test('invalid values are reported per row', () => {
    const env = new TestEnv({workspace: webWorkspace()}).list('Tags');
    const a = env.row('Tags', 'Tag Name', 'GA4 - purchase');
    const b = env.row('Tags', 'Tag Name', 'GA4 - page_view');
    const c = env.row('Tags', 'Tag Name', 'Custom HTML - chat');
    env.set('Tags', a, {'Firing Triggers': 'Does not exist'});
    env.set('Tags', b, {'Tag Name': 'GA4 - purchase'});
    env.set('Tags', c, {'Tag Firing Priority': 'high', Paused: true});
    env.apply('Tags');
    assert.match(env.result('Tags', a), /Trigger not found: Does not exist/);
    assert.match(env.result('Tags', b), /already used/);
    assert.match(env.result('Tags', c), /whole number/);
    assert.equal(env.entity('tag', '3').paused, undefined);
  });

  test('creating a tag from scratch requires a copied row', () => {
    const env = new TestEnv({workspace: webWorkspace()}).list('Tags');
    const row = env.addRow('Tags', {'Tag Name': 'New', Action: 'Create'});
    env.apply('Tags');
    assert.match(env.result('Tags', row), /copy the row of an existing one/);
  });
});

describe('GA4 Event Tags sheet', () => {
  test('lists GA4 event tags only', () => {
    const env = new TestEnv({workspace: webWorkspace()}).list('GA4 Event Tags');
    assert.equal(env.values('GA4 Event Tags').length, 2 + 2);
    const row = env.row('GA4 Event Tags', 'Tag Name', 'GA4 - purchase');
    assert.equal(env.get('GA4 Event Tags', row, 'Event Parameters'),
        'value = {{DLV - value}}\ncurrency = EUR');
    assert.equal(env.get('GA4 Event Tags', row, 'User Properties'),
        'customer_type = new');
    assert.equal(env.get('GA4 Event Tags', row, 'Send Ecommerce Data'), true);
    assert.equal(env.get('GA4 Event Tags', row, 'Ecommerce Data Source'),
        'dataLayer');
  });

  test('creates a GA4 event tag in a new folder', () => {
    const env = new TestEnv({workspace: webWorkspace()}).list('GA4 Event Tags');
    const row = env.addRow('GA4 Event Tags', {
      Action: 'Create', 'Tag Name': 'GA4 - sign_up', 'Event Name': 'sign_up',
      'Measurement ID': '{{GA4 ID}}',
      'Event Parameters': 'method = {{DLV - method}}\nlogin_type = email',
      'Firing Triggers': 'CE - purchase', 'Folder': 'GA4 - New'
    });
    env.apply('GA4 Event Tags');
    assert.match(env.result('GA4 Event Tags', row), /^✔ Created/);
    const tag = env.entityByName('tag', 'GA4 - sign_up');
    assert.equal(tag.type, 'gaawe');
    assert.equal(paramValue(tag, 'eventName'), 'sign_up');
    assert.deepEqual(paramValue(tag, 'eventSettingsTable'), [
      {parameter: 'method', parameterValue: '{{DLV - method}}'},
      {parameter: 'login_type', parameterValue: 'email'}]);
    const folder = env.entityByName('folder', 'GA4 - New');
    assert.ok(folder);
    assert.equal(tag.parentFolderId, folder.folderId);
  });

  test('updates only the changed settings', () => {
    const env = new TestEnv({workspace: webWorkspace()}).list('GA4 Event Tags');
    const row = env.row('GA4 Event Tags', 'Tag Name', 'GA4 - purchase');
    env.set('GA4 Event Tags', row, {
      'Measurement ID': 'G-NEW0000001',
      'Ecommerce Data Source': '{{Ecommerce Object}}'
    });
    env.apply('GA4 Event Tags');
    const tag = env.entity('tag', '1');
    assert.equal(paramValue(tag, 'measurementIdOverride'), 'G-NEW0000001');
    assert.equal(paramValue(tag, 'getEcommerceDataFrom'), 'customObject');
    assert.equal(paramValue(tag, 'ecommerceMacroData'), '{{Ecommerce Object}}');
    assert.equal(paramValue(tag, 'eventSettingsTable').length, 2);
    assert.equal(tag.parentFolderId, '90');
  });

  test('required values are validated', () => {
    const env = new TestEnv({workspace: webWorkspace()}).list('GA4 Event Tags');
    const row = env.addRow('GA4 Event Tags', {Action: 'Create',
                                              'Tag Name': 'GA4 - broken'});
    env.apply('GA4 Event Tags');
    assert.match(env.result('GA4 Event Tags', row),
        /Missing required value\(s\): Event Name, Measurement ID/);
    assert.equal(env.entityByName('tag', 'GA4 - broken'), undefined);
  });
});

describe('Event Parameters & User Properties sheet', () => {
  const sheet = 'Event Parameters & User Properties';

  test('lists one row per parameter and placeholders', () => {
    const env = new TestEnv({workspace: webWorkspace()}).list(sheet);
    const values = env.values(sheet).slice(2);
    const rows = values.map(r => [r[3], r[6], r[7], r[8]]);
    assert.deepEqual(rows.filter(r => r[0] === 'GA4 - purchase'), [
      ['GA4 - purchase', 'Event Parameter', 'value', '{{DLV - value}}'],
      ['GA4 - purchase', 'Event Parameter', 'currency', 'EUR'],
      ['GA4 - purchase', 'User Property', 'customer_type', 'new']]);
    // Placeholder for the GA4 tag without parameters.
    assert.deepEqual(rows.filter(r => r[0] === 'GA4 - page_view'),
        [['GA4 - page_view', 'Event Parameter', '', '']]);
    // Google tag, Ads, Floodlight and variables are included too.
    assert.ok(rows.some(r => r[0] === 'Google Tag' &&
        r[1] === 'Configuration Parameter' && r[2] === 'send_page_view'));
    assert.ok(rows.some(r => r[0] === 'Event Settings' && r[2] === 'site'));
  });

  test('updates, adds, moves and deletes parameters', () => {
    const env = new TestEnv({workspace: webWorkspace()}).list(sheet);
    const value = env.row(sheet, 'Parameter / Property Name', 'value');
    const currency = env.row(sheet, 'Parameter / Property Name', 'currency');
    const customer = env.row(sheet, 'Parameter / Property Name',
        'customer_type');
    const placeholder = env.row(sheet, 'Tag / Variable Name', 'GA4 - page_view');
    env.set(sheet, value, {Value: '{{DLV - revenue}}'});
    env.set(sheet, currency, {Action: 'Delete'});
    env.set(sheet, customer, {Setting: 'Event Parameter',
                              'Parameter / Property Name': 'customer_kind'});
    env.set(sheet, placeholder, {'Parameter / Property Name': 'page_type',
                                 Value: '{{Page Type}}'});
    const added = env.addRow(sheet, {Action: 'Create', 'Item Type': 'Tag',
      'Tag / Variable Name': 'GA4 - purchase', Setting: 'User Property',
      'Parameter / Property Name': 'loyalty', Value: 'gold'});
    const addedVar = env.addRow(sheet, {Action: 'Create',
      'Tag / Variable Name': 'Event Settings', Setting: 'User Property',
      'Parameter / Property Name': 'plan', Value: 'pro'});
    env.preview(sheet);
    assert.match(env.result(sheet, currency), /^Will be removed/);
    assert.match(env.result(sheet, added), /^Will be added/);
    env.apply(sheet);
    [value, currency, customer, placeholder, added, addedVar].forEach(row =>
      assert.match(env.result(sheet, row), /^✔/, `row ${row}`));
    const tag = env.entity('tag', '1');
    assert.deepEqual(paramValue(tag, 'eventSettingsTable'), [
      {parameter: 'value', parameterValue: '{{DLV - revenue}}'},
      {parameter: 'customer_kind', parameterValue: 'new'}]);
    assert.deepEqual(paramValue(tag, 'userProperties'),
        [{name: 'loyalty', value: 'gold'}]);
    assert.deepEqual(paramValue(env.entity('tag', '2'), 'eventSettingsTable'),
        [{parameter: 'page_type', parameterValue: '{{Page Type}}'}]);
    assert.deepEqual(paramValue(env.entity('variable', '23'), 'userProperties'),
        [{name: 'plan', value: 'pro'}]);
    // Rows applied together for one tag cause one update of the tag.
    const bulk = env.fake.requestsTo('POST', /bulk_update/).slice(-1)[0];
    const tagChanges = bulk.body.changes.filter(c => c.tag && c.tag.tagId === '1');
    assert.equal(tagChanges.length, 1);
    // The rows can be edited again without listing.
    env.set(sheet, customer, {Value: 'returning'});
    env.apply(sheet);
    assert.match(env.result(sheet, customer), /^✔ Updated/);
    assert.deepEqual(paramValue(env.entity('tag', '1'), 'eventSettingsTable')[1],
        {parameter: 'customer_kind', parameterValue: 'returning'});
  });

  test('stale rows and invalid settings are reported', () => {
    const env = new TestEnv({workspace: webWorkspace()}).list(sheet);
    const value = env.row(sheet, 'Parameter / Property Name', 'value');
    // The parameter is changed in GTM after listing.
    const tag = env.entity('tag', '1');
    tag.parameter.find(p => p.key === 'eventSettingsTable').list[0].map[1].value =
        'changed in GTM';
    env.fake.seed({entities: {[env.ws.workspacePath]: {tag: [tag]}}});
    env.set(sheet, value, {Value: 'mine'});
    const bad = env.addRow(sheet, {Action: 'Create',
      'Tag / Variable Name': 'Custom HTML - chat', Setting: 'Event Parameter',
      'Parameter / Property Name': 'x', Value: 'y'});
    env.apply(sheet);
    assert.match(env.result(sheet, value), /was not found in GTM anymore/);
    assert.match(env.result(sheet, bad), /can't be used in this sheet/);
  });
});

describe('Google Tags sheet', () => {
  test('creates a destination tag and validates IDs', () => {
    const env = new TestEnv({workspace: webWorkspace()}).list('Google Tags');
    assert.equal(env.values('Google Tags').length, 2 + 1);
    const good = env.addRow('Google Tags', {Action: 'Create',
      'Tag Name': 'Google Analytics G-NEW1234567',
      'Tag Type': 'Google Analytics Destination',
      'Tag / Destination ID': 'G-NEW1234567',
      'Configuration Parameters': 'server_container_url = https://sgtm.shop.com',
      'Firing Triggers': 'Initialization - All Pages'});
    const bad = env.addRow('Google Tags', {Action: 'Create',
      'Tag Name': 'Ads destination', 'Tag Type': 'Google Ads Destination',
      'Tag / Destination ID': 'G-WRONG'});
    env.apply('Google Tags');
    assert.match(env.result('Google Tags', good), /^✔ Created/);
    assert.match(env.result('Google Tags', bad), /not a valid ID/);
    const tag = env.entityByName('tag', 'Google Analytics G-NEW1234567');
    assert.equal(tag.type, 'dest_ga');
    assert.equal(paramValue(tag, 'destinationId'), 'G-NEW1234567');
    assert.deepEqual(tag.firingTriggerId, ['2147479573']);
  });
});

describe('Google Ads & Floodlight Tags sheet', () => {
  const sheet = 'Google Ads & Floodlight Tags';

  test('lists tags with type-specific columns', () => {
    const env = new TestEnv({workspace: webWorkspace()}).list(sheet);
    const ads = env.row(sheet, 'Tag Name', 'Ads - Purchase');
    const fl = env.row(sheet, 'Tag Name', 'Floodlight - Purchase');
    assert.equal(env.get(sheet, ads, 'Conversion Label'), 'abcDEF');
    assert.equal(env.get(sheet, ads, 'Advertiser ID'), '');
    assert.equal(env.get(sheet, fl, 'Counting Method'), 'Standard');
    assert.equal(env.get(sheet, fl, 'Custom Variables'), 'u1 = {{DLV - value}}');
    const cell = env.rt.spreadsheet.getCell(sheet,
        String.fromCharCode(64 + env.col(sheet, 'Advertiser ID')) + ads);
    assert.ok(/f1f3f4/i.test(JSON.stringify(cell.format)));
  });

  test('creates and updates Ads and Floodlight tags', () => {
    const env = new TestEnv({workspace: webWorkspace()}).list(sheet);
    const ads = env.row(sheet, 'Tag Name', 'Ads - Purchase');
    env.set(sheet, ads, {'Conversion Label': 'NEWlabel',
                         'Advertiser ID': '999'});
    const fl = env.addRow(sheet, {Action: 'Create',
      'Tag Name': 'Floodlight - Sign up', 'Tag Type': 'Floodlight Counter',
      'Advertiser ID': '1234567', 'Group Tag String (type=)': 'leads',
      'Activity Tag String (cat=)': 'signup0', 'Counting Method': 'Unique',
      'Custom Variables': 'u2 = {{Plan}}', 'Firing Triggers': 'CE - purchase'});
    const missing = env.addRow(sheet, {Action: 'Create',
      'Tag Name': 'Ads - incomplete',
      'Tag Type': 'Google Ads Conversion Tracking',
      'Conversion ID': '111'});
    env.apply(sheet);
    assert.match(env.result(sheet, ads), /^✔ Updated: Conversion Label\. .*does not apply/);
    assert.equal(paramValue(env.entity('tag', '5'), 'conversionLabel'), 'NEWlabel');
    assert.equal(paramValue(env.entity('tag', '5'), 'advertiserId'), undefined);
    assert.match(env.result(sheet, fl), /^✔ Created/);
    const created = env.entityByName('tag', 'Floodlight - Sign up');
    assert.equal(created.type, 'flc');
    assert.equal(paramValue(created, 'ordinalType'), 'UNIQUE');
    assert.deepEqual(paramValue(created, 'customVariable'),
        [{key: 'u2', value: '{{Plan}}'}]);
    assert.match(env.result(sheet, missing), /Conversion Label/);
  });
});

describe('Variables sheet', () => {
  test('lists variables with main value and usage', () => {
    const env = new TestEnv({workspace: webWorkspace()}).list('Variables');
    const ga4 = env.row('Variables', 'Variable Name', 'GA4 ID');
    assert.equal(env.get('Variables', ga4, 'Variable Type'), 'Constant');
    assert.equal(env.get('Variables', ga4, 'Value / Main Setting'),
        'G-OLD0000001');
    assert.equal(env.get('Variables', ga4, 'Used In (count)'), '1');
    assert.equal(env.get('Variables', ga4, 'Used In'), 'Tags: GA4 - purchase');
    const dlv = env.row('Variables', 'Variable Name', 'DLV - value');
    assert.equal(env.get('Variables', dlv, 'Value / Main Setting'),
        'ecommerce.value');
    assert.equal(env.get('Variables', dlv, 'Used In (count)'), '3');
    const unused = env.row('Variables', 'Variable Name', 'Unused constant');
    assert.equal(env.get('Variables', unused, 'Used In (count)'), '0');
  });

  test('creates, updates and deletes variables', () => {
    const env = new TestEnv({workspace: webWorkspace()}).list('Variables');
    const ga4 = env.row('Variables', 'Variable Name', 'GA4 ID');
    env.set('Variables', ga4, {'Value / Main Setting': 'G-NEW0000001'});
    const dlv = env.addRow('Variables', {Action: 'Create',
      'Variable Name': 'DLV - method', 'Variable Type': 'Data Layer Variable',
      'Value / Main Setting': 'method', 'Default Value': 'unknown',
      'Convert Case (Format Value)': 'Lowercase'});
    const lookup = env.addRow('Variables', {Action: 'Create',
      'Variable Name': 'Lookup - Measurement ID', 'Variable Type': 'Lookup Table',
      'Value / Main Setting': '{{Page Hostname}}',
      'Lookup Table (input = output)': 'shop.com = G-AAA\nstaging.shop.com = G-BBB',
      'Default Value': 'G-CCC'});
    const css = env.addRow('Variables', {Action: 'Create',
      'Variable Name': 'DOM - price', 'Variable Type': 'DOM Element',
      'Value / Main Setting': 'CSS: .price span'});
    const query = env.addRow('Variables', {Action: 'Create',
      'Variable Name': 'URL - utm_source', 'Variable Type': 'URL',
      'Value / Main Setting': 'QUERY: utm_source'});
    const unused = env.row('Variables', 'Variable Name', 'Unused constant');
    env.set('Variables', unused, {Action: 'Delete'});
    const used = env.row('Variables', 'Variable Name', 'DLV - value');
    env.set('Variables', used, {Action: 'Delete'});
    env.apply('Variables');
    [ga4, dlv, lookup, css, query, unused].forEach(row =>
      assert.match(env.result('Variables', row), /^✔/, `row ${row}`));
    assert.match(env.result('Variables', used), /can't be deleted because it is used/);
    assert.equal(paramValue(env.entity('variable', '20'), 'value'), 'G-NEW0000001');
    const created = env.entityByName('variable', 'DLV - method');
    assert.equal(created.type, 'v');
    assert.equal(paramValue(created, 'name'), 'method');
    assert.equal(paramValue(created, 'dataLayerVersion'), '2');
    assert.equal(paramValue(created, 'setDefaultValue'), 'true');
    assert.equal(paramValue(created, 'defaultValue'), 'unknown');
    assert.equal(created.formatValue.caseConversionType, 'lowercase');
    const table = env.entityByName('variable', 'Lookup - Measurement ID');
    assert.deepEqual(paramValue(table, 'map'), [
      {key: 'shop.com', value: 'G-AAA'}, {key: 'staging.shop.com', value: 'G-BBB'}]);
    assert.equal(paramValue(table, 'input'), '{{Page Hostname}}');
    const dom = env.entityByName('variable', 'DOM - price');
    assert.equal(paramValue(dom, 'selectorType'), 'CSS');
    assert.equal(paramValue(dom, 'elementSelector'), '.price span');
    const url = env.entityByName('variable', 'URL - utm_source');
    assert.equal(paramValue(url, 'component'), 'QUERY');
    assert.equal(paramValue(url, 'queryKey'), 'utm_source');
    assert.equal(env.entity('variable', '22'), undefined);
    assert.ok(env.entity('variable', '21'));
  });
});

describe('Triggers sheet', () => {
  test('lists triggers with conditions and usage', () => {
    const env = new TestEnv({workspace: webWorkspace()}).list('Triggers');
    const ce = env.row('Triggers', 'Trigger Name', 'CE - purchase');
    assert.equal(env.get('Triggers', ce, 'Event Name (Custom Event)'), 'purchase');
    assert.equal(env.get('Triggers', ce, 'Used By'),
        'GA4 - purchase (firing)\nAds - Purchase (firing)\nFloodlight - Purchase (firing)');
    const pv = env.row('Triggers', 'Trigger Name', 'PV - Checkout');
    assert.equal(env.get('Triggers', pv, 'Conditions (fires on)'),
        '{{Page Path}} contains /checkout');
  });

  test('creates, updates and deletes triggers', () => {
    const env = new TestEnv({workspace: webWorkspace()}).list('Triggers');
    const created = env.addRow('Triggers', {Action: 'Create',
      'Trigger Name': 'CE - sign_up', 'Trigger Type': 'Custom Event',
      'Event Name (Custom Event)': 'sign_up|login', 'Use RegEx Matching': true,
      'Conditions (fires on)':
          '{{Page Path}} does not contain /admin\n{{Page Hostname}} equals shop.com',
      Folder: 'Triggers'});
    const click = env.addRow('Triggers', {Action: 'Create',
      'Trigger Name': 'Link - outbound', 'Trigger Type': 'Click - Just Links',
      'Conditions (fires on)': '{{Click URL}} does not contain shop.com'});
    const pv = env.row('Triggers', 'Trigger Name', 'PV - Checkout');
    env.set('Triggers', pv, {'Conditions (fires on)': '{{Page Path}} starts with /checkout'});
    const unused = env.row('Triggers', 'Trigger Name', 'CE - unused');
    env.set('Triggers', unused, {Action: 'Delete'});
    const used = env.row('Triggers', 'Trigger Name', 'CE - purchase');
    env.set('Triggers', used, {Action: 'Delete'});
    const invalid = env.addRow('Triggers', {Action: 'Create',
      'Trigger Name': 'Broken', 'Trigger Type': 'Page View',
      'Conditions (fires on)': 'Page Path contains x'});
    env.apply('Triggers');
    [created, click, pv, unused].forEach(row =>
      assert.match(env.result('Triggers', row), /^✔/, `row ${row}`));
    assert.match(env.result('Triggers', used), /can't be deleted/);
    assert.match(env.result('Triggers', invalid), /must start with a variable/);
    const trigger = env.entityByName('trigger', 'CE - sign_up');
    assert.equal(trigger.type, 'customEvent');
    assert.equal(trigger.customEventFilter[0].type, 'matchRegex');
    assert.equal(trigger.filter.length, 2);
    assert.equal(trigger.filter[0].type, 'contains');
    assert.ok(trigger.parentFolderId);
    assert.equal(env.entityByName('trigger', 'Link - outbound').type, 'linkClick');
    assert.equal(env.entity('trigger', '11').filter[0].type, 'startsWith');
    assert.equal(env.entity('trigger', '12'), undefined);
  });
});

describe('Advanced Settings sheet', () => {
  const sheet = 'Advanced Settings';

  test('edits any parameter of any item', () => {
    const env = new TestEnv({workspace: webWorkspace()}).list(sheet);
    const html = env.row(sheet, 'Parameter Key', 'html');
    const pixel = env.row(sheet, 'Parameter Key', 'pixelId');
    const table = env.row(sheet, 'Parameter Key', 'eventSettingsTable',
        r => r[3] === 'GA4 - purchase');
    assert.equal(env.get(sheet, table, 'Value Type'), 'list');
    assert.equal(env.get(sheet, table, 'Value'),
        '[{"parameter":"value","parameterValue":"{{DLV - value}}"},' +
        '{"parameter":"currency","parameterValue":"EUR"}]');
    env.set(sheet, html, {Value: '<script>chat("v2")</script>'});
    env.set(sheet, pixel, {Action: 'Delete'});
    env.set(sheet, table, {Value: '[{"parameter":"value","parameterValue":"1"}]'});
    const added = env.addRow(sheet, {Action: 'Create', 'Item Type': 'Tag',
      'Item Name': 'Meta Pixel', 'Parameter Key': 'pixelIds',
      Value: '["222","333"]'});
    const badJson = env.addRow(sheet, {Action: 'Create', 'Item Type': 'Tag',
      'Item Name': 'Meta Pixel', 'Parameter Key': 'config', Value: '[oops'});
    env.apply(sheet);
    [html, pixel, table, added].forEach(row =>
      assert.match(env.result(sheet, row), /^✔/, `row ${row}`));
    assert.match(env.result(sheet, badJson), /valid JSON/);
    assert.equal(paramValue(env.entity('tag', '3'), 'html'),
        '<script>chat("v2")</script>');
    const meta = env.entity('tag', '7');
    assert.equal(paramValue(meta, 'pixelId'), undefined);
    assert.deepEqual(paramValue(meta, 'pixelIds'), ['222', '333']);
    assert.deepEqual(paramValue(env.entity('tag', '1'), 'eventSettingsTable'),
        [{parameter: 'value', parameterValue: '1'}]);
  });
});

describe('Google Tag Settings sheet', () => {
  const sheet = 'Google Tag Settings';

  test('lists and edits container-wide Google tag settings', () => {
    const env = new TestEnv({workspace: webWorkspace()});
    env.fake.seed({destinations: [{accountId: '6000001', containerId: '9000001',
                                   destinationId: 'G-ABC1234567',
                                   name: 'Shop property'}]});
    env.list(sheet);
    assert.match(String(env.values(sheet)[0][0]),
        /Destinations: G-ABC1234567 \(Shop property\)/);
    const scroll = env.row(sheet, 'Parameter Key', 'enableScroll');
    assert.equal(env.get(sheet, scroll, 'Setting'), 'Automatic event detection');
    env.set(sheet, scroll, {Value: 'true'});
    const domains = env.addRow(sheet, {Action: 'Create',
      Setting: 'Configure your domains', 'Parameter Key': 'rules',
      Value: '[{"matchType":"CONTAINS","matchValue":"shop.com"}]'});
    env.apply(sheet);
    assert.match(env.result(sheet, scroll), /^✔ Updated/);
    assert.match(env.result(sheet, domains), /^✔/);
    assert.equal(paramValue(env.entity('gtagConfig', '80'), 'enableScroll'), 'true');
    const created = env.fake.list(env.ws.workspacePath, 'gtagConfig')
        .find(c => c.type === 'ogt_cross_domain');
    assert.deepEqual(paramValue(created, 'rules'),
        [{matchType: 'CONTAINS', matchValue: 'shop.com'}]);
  });

  test('is not available for server containers', () => {
    const env = new TestEnv({workspace: {usageContext: ['server']}});
    env.click('listActiveSheet', sheet);
    assert.match(env.alertMessages()[0], /works with web and Google tag containers/);
  });
});
