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
 * @fileoverview End-to-end tests of the feature sheets with a server
 * container.
 */
'use strict';

const {test, describe} = require('node:test');
const assert = require('node:assert/strict');
const {TestEnv, params, paramValue} = require('./helpers');

/** A server workspace with typical content. */
function serverWorkspace() {
  return {
    containerName: 'Shop Server', publicId: 'GTM-SRV0001',
    usageContext: ['server'],
    triggers: [
      {triggerId: '10', name: 'Client - GA4', type: 'always', filter: [{
        type: 'equals', parameter: [
          {type: 'template', key: 'arg0', value: '{{Client Name}}'},
          {type: 'template', key: 'arg1', value: 'GA4'}]}]},
      {triggerId: '11', name: 'Event - purchase', type: 'always', filter: [{
        type: 'equals', parameter: [
          {type: 'template', key: 'arg0', value: '{{Event Name}}'},
          {type: 'template', key: 'arg1', value: 'purchase'}]}]}
    ],
    variables: [
      {variableId: '20', name: 'ED - value', type: 'ed',
       parameter: params({keyPath: 'value'})}
    ],
    clients: [
      {clientId: '30', name: 'GA4', type: 'gaaw_client', priority: 10,
       parameter: params({activateDefaultPaths: true, cookieManagement: 'server',
                          cookieName: 'FPID'})},
      {clientId: '31', name: 'GTM', type: 'gtm_client', priority: 5,
       parameter: params({allowedContainerIds: [['GTM-WEB0001', '/gtm']],
                          activateResponseCompression: true},
           {allowedContainerIds: ['containerId', 'path']})}
    ],
    tags: [
      {tagId: '1', name: 'GA4 - all events', type: 'sgtmgaaw',
       firingTriggerId: ['10'],
       parameter: params({epToIncludeDropdown: 'all',
                          epToExclude: [['ip_override']],
                          eventParameters: [['source', 'server']],
                          upToIncludeDropdown: 'all'},
           {epToExclude: ['fieldName'],
            eventParameters: ['fieldName', 'value']})},
      {tagId: '2', name: 'Ads - Purchase (server)', type: 'sgtmadsct',
       firingTriggerId: ['11'],
       parameter: params({conversionId: '123', conversionLabel: 'abc'})},
      {tagId: '3', name: 'Conversion Linker (server)', type: 'sgtmadscl',
       firingTriggerId: ['2147479574'],
       parameter: params({enableLinkerParams: true})},
      {tagId: '4', name: 'Meta CAPI', type: 'cvt_9000001_40',
       firingTriggerId: ['10'], parameter: params({pixelId: '999'})}
    ],
    transformations: [
      {transformationId: '50', name: 'Exclude IP', type: 'tf_exclude_params',
       parameter: params({allTagsExcept: true,
                          excludedParamsTable: [['ip_override']]},
           {excludedParamsTable: ['excludedParams']})}
    ],
    templates: [{templateId: '40', name: 'Meta Conversions API',
                 templateData: '___INFO___'}]
  };
}

describe('server container', () => {
  test('context and container-specific sheets', () => {
    const env = new TestEnv({workspace: serverWorkspace()});
    const ctx = env.rt.call('getWorkspaceContext');
    assert.equal(ctx.isServer, true);
    assert.equal(env.values('GTM Workspace')[5][13], 'Server');
    env.click('listActiveSheet', 'GA4 Event Tags');
    assert.match(env.alertMessages()[0], /works with web containers/);
  });

  test('Tags sheet lists server tags with the built-in trigger', () => {
    const env = new TestEnv({workspace: serverWorkspace()}).list('Tags');
    const linker = env.row('Tags', 'Tag Name', 'Conversion Linker (server)');
    assert.equal(env.get('Tags', linker, 'Firing Triggers'), 'All Pages');
    assert.equal(env.get('Tags', linker, 'Tag Type'), 'Conversion Linker (Server)');
    const meta = env.row('Tags', 'Tag Name', 'Meta CAPI');
    assert.equal(env.get('Tags', meta, 'Tag Type'), 'Meta Conversions API');
    // Web-only settings do not apply to server tags.
    assert.equal(env.get('Tags', meta, 'Additional Consent Checks'), '');
    env.set('Tags', meta, {Paused: true, 'Firing Triggers': 'Event - purchase'});
    env.apply('Tags');
    assert.match(env.result('Tags', meta), /^✔ Updated/);
    assert.equal(env.entity('tag', '4').paused, true);
    assert.deepEqual(env.entity('tag', '4').firingTriggerId, ['11']);
  });

  test('Clients sheet', () => {
    const env = new TestEnv({workspace: serverWorkspace()}).list('Clients');
    const ga4 = env.row('Clients', 'Client Name', 'GA4');
    assert.equal(env.get('Clients', ga4, 'Client Type'),
        'Google Analytics: GA4 (Web)');
    assert.equal(env.get('Clients', ga4, 'Cookie Management'), 'Server managed');
    assert.equal(env.get('Clients', ga4, 'Default GA4 Paths'), true);
    env.set('Clients', ga4, {'Cookie Name': '_fpid', Priority: '20',
                             'Cookie SameSite': 'Lax'});
    const gtm = env.addRow('Clients', {Action: 'Create',
      'Client Name': 'GTM 2', 'Client Type': 'Google Tag Manager: Web Container',
      'Allowed Container IDs (ID = path)': 'GTM-AAA1111 = /a\nGTM-BBB2222 = /b',
      'Priority': '1'});
    const broken = env.addRow('Clients', {Action: 'Create',
      'Client Name': 'GTM 3', 'Client Type': 'Google Tag Manager: Web Container'});
    env.apply('Clients');
    assert.match(env.result('Clients', ga4), /^✔ Updated/);
    const client = env.entity('client', '30');
    assert.equal(paramValue(client, 'cookieName'), '_fpid');
    assert.equal(paramValue(client, 'cookieSamesite'), 'lax');
    assert.equal(client.priority, 20);
    assert.match(env.result('Clients', gtm), /^✔ Created/);
    assert.deepEqual(paramValue(env.entityByName('client', 'GTM 2'),
        'allowedContainerIds'), [{containerId: 'GTM-AAA1111', path: '/a'},
                                 {containerId: 'GTM-BBB2222', path: '/b'}]);
    assert.match(env.result('Clients', broken), /Allowed Container IDs/);
  });

  test('Server GA4 Tags sheet', () => {
    const env = new TestEnv({workspace: serverWorkspace()}).list('Server GA4 Tags');
    const row = env.row('Server GA4 Tags', 'Tag Name', 'GA4 - all events');
    assert.equal(env.get('Server GA4 Tags', row, 'Default Parameters to Include'),
        'All');
    assert.equal(env.get('Server GA4 Tags', row, 'Parameters to Exclude'),
        'ip_override');
    assert.equal(env.get('Server GA4 Tags', row, 'Parameters to Add / Edit'),
        'source = server');
    env.set('Server GA4 Tags', row, {
      'Measurement ID': 'G-SRV0000001',
      'Parameters to Exclude': 'ip_override, user_agent',
      'Parameters to Add / Edit': 'source = server\nregion = {{Region}}',
      'Redact Visitor IP Address': true});
    const created = env.addRow('Server GA4 Tags', {Action: 'Create',
      'Tag Name': 'GA4 - purchase only', 'Event Name': 'purchase',
      'Firing Triggers': 'Event - purchase'});
    env.apply('Server GA4 Tags');
    assert.match(env.result('Server GA4 Tags', row), /^✔ Updated/);
    const tag = env.entity('tag', '1');
    assert.equal(paramValue(tag, 'measurementId'), 'G-SRV0000001');
    assert.equal(paramValue(tag, 'redactVisitorIp'), 'true');
    assert.deepEqual(paramValue(tag, 'epToExclude'),
        [{fieldName: 'ip_override'}, {fieldName: 'user_agent'}]);
    assert.deepEqual(paramValue(tag, 'eventParameters'), [
      {fieldName: 'source', value: 'server'},
      {fieldName: 'region', value: '{{Region}}'}]);
    assert.match(env.result('Server GA4 Tags', created), /^✔ Created/);
    const newTag = env.entityByName('tag', 'GA4 - purchase only');
    assert.equal(newTag.type, 'sgtmgaaw');
    assert.equal(paramValue(newTag, 'epToIncludeDropdown'), 'all');
  });

  test('Server Ads & Floodlight Tags sheet', () => {
    const sheet = 'Server Ads & Floodlight Tags';
    const env = new TestEnv({workspace: serverWorkspace()}).list(sheet);
    const ads = env.row(sheet, 'Tag Name', 'Ads - Purchase (server)');
    env.set(sheet, ads, {'Conversion Value': '{{ED - value}}'});
    const fl = env.addRow(sheet, {Action: 'Create',
      'Tag Name': 'Floodlight (server)', 'Tag Type': 'Floodlight Counter (Server)',
      'Advertiser ID': '765', 'Group Tag String (type=)': 'g',
      'Activity Tag String (cat=)': 'a', 'Counting Method': 'Per Session',
      'Firing Triggers': 'Event - purchase'});
    const linker = env.row(sheet, 'Tag Name', 'Conversion Linker (server)');
    env.set(sheet, linker, {'Override Cookie Settings': true,
                            'Cookie Domain': 'shop.com',
                            'Cookie SameSite': 'Strict'});
    env.apply(sheet);
    [ads, fl, linker].forEach(row =>
      assert.match(env.result(sheet, row), /^✔/, `row ${row}`));
    assert.equal(paramValue(env.entity('tag', '2'), 'conversionValue'),
        '{{ED - value}}');
    const created = env.entityByName('tag', 'Floodlight (server)');
    assert.equal(created.type, 'sgtmflc');
    assert.equal(paramValue(created, 'ordinalType'), 'PER_SESSION');
    const tag = env.entity('tag', '3');
    assert.equal(paramValue(tag, 'enableCookieOverrides'), 'true');
    assert.equal(paramValue(tag, 'cookieSamesite'), 'strict');
  });

  test('Transformations sheet', () => {
    const env = new TestEnv({workspace: serverWorkspace()}).list('Transformations');
    const row = env.row('Transformations', 'Transformation Name', 'Exclude IP');
    assert.equal(env.get('Transformations', row, 'Excluded Parameters'),
        'ip_override');
    assert.equal(env.get('Transformations', row, 'Affected Tags Mode'),
        'All tags except');
    env.set('Transformations', row, {'Excluded Parameters': 'ip_override, user_agent'});
    const created = env.addRow('Transformations', {Action: 'Create',
      'Transformation Name': 'Augment - region',
      'Transformation Type': 'Augment event',
      'Parameters to Add / Change': 'region = {{Region}}',
      'Affected Tags Mode': 'Only these tags',
      'Affected Tags': 'GA4 - all events, Meta CAPI', 'Priority': '5'});
    const empty = env.addRow('Transformations', {Action: 'Create',
      'Transformation Name': 'Allow - empty',
      'Transformation Type': 'Allow parameters'});
    env.apply('Transformations');
    assert.match(env.result('Transformations', row), /^✔ Updated/);
    assert.deepEqual(paramValue(env.entity('transformation', '50'),
        'excludedParamsTable'),
        [{excludedParams: 'ip_override'}, {excludedParams: 'user_agent'}]);
    assert.match(env.result('Transformations', created), /^✔ Created/);
    const transformation = env.entityByName('transformation', 'Augment - region');
    assert.equal(paramValue(transformation, 'allTagsExcept'), 'false');
    const affected = transformation.parameter.find(p => p.key === 'affectedTags');
    assert.deepEqual(affected.list.map(i => i.map[0]), [
      {type: 'tagReference', key: 'tagReference', value: 'GA4 - all events',
       isWeakReference: false},
      {type: 'tagReference', key: 'tagReference', value: 'Meta CAPI',
       isWeakReference: false}]);
    assert.match(env.result('Transformations', empty), /at least one parameter/);
  });

  test('Event Parameters sheet includes server GA4 settings', () => {
    const sheet = 'Event Parameters & User Properties';
    const env = new TestEnv({workspace: serverWorkspace()}).list(sheet);
    const exclude = env.row(sheet, 'Parameter / Property Name', 'ip_override');
    assert.equal(env.get(sheet, exclude, 'Setting'), 'Parameter to Exclude');
    const added = env.addRow(sheet, {Action: 'Create',
      'Tag / Variable Name': 'GA4 - all events',
      Setting: 'User Property to Add / Edit',
      'Parameter / Property Name': 'tier', Value: 'gold'});
    env.apply(sheet);
    assert.match(env.result(sheet, added), /^✔ Added/);
    assert.deepEqual(paramValue(env.entity('tag', '1'), 'userProperties'),
        [{fieldName: 'tier', value: 'gold'}]);
  });

  test('Variables and Triggers sheets work with server items', () => {
    const env = new TestEnv({workspace: serverWorkspace()}).list('Variables');
    const ed = env.row('Variables', 'Variable Name', 'ED - value');
    assert.equal(env.get('Variables', ed, 'Variable Type'), 'Event Data');
    assert.equal(env.get('Variables', ed, 'Value / Main Setting'), 'value');
    const created = env.addRow('Variables', {Action: 'Create',
      'Variable Name': 'Header - UA', 'Variable Type': 'Request Header',
      'Value / Main Setting': 'user-agent'});
    env.apply('Variables');
    assert.match(env.result('Variables', created), /^✔ Created/);
    assert.equal(paramValue(env.entityByName('variable', 'Header - UA'),
        'headerName'), 'user-agent');
    env.list('Triggers');
    const trigger = env.addRow('Triggers', {Action: 'Create',
      'Trigger Name': 'Event - sign_up', 'Trigger Type': 'Custom (Server)',
      'Conditions (fires on)': '{{Event Name}} equals sign_up'});
    env.apply('Triggers');
    assert.match(env.result('Triggers', trigger), /^✔ Created/);
    assert.equal(env.entityByName('trigger', 'Event - sign_up').type, 'always');
  });
});
