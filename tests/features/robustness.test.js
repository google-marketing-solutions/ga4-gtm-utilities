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
 * @fileoverview Robustness tests: partial failures of bulk updates, stale
 * sheets, rate limits, the execution time budget, API errors and Google tag
 * containers.
 */
'use strict';

const {test, describe} = require('node:test');
const assert = require('node:assert/strict');
const {TestEnv, params, paramValue} = require('./helpers');

/**
 * A web workspace with n Custom HTML tags.
 * @param {number} n
 * @return {!Object}
 */
function workspaceWithTags(n) {
  const tags = [];
  for (let i = 1; i <= n; i++) {
    tags.push({tagId: String(i), name: `Tag ${i}`, type: 'html',
               firingTriggerId: ['2147479553'],
               parameter: params({html: `<script>${i}</script>`})});
  }
  return {publicId: 'GTM-ROB0001', tags};
}

/**
 * Pauses all listed tags of the Tags sheet.
 * @param {!TestEnv} env
 * @param {number} n
 */
function pauseTags(env, n) {
  for (let i = 1; i <= n; i++) {
    env.set('Tags', env.row('Tags', 'Tag Name', `Tag ${i}`), {Paused: true});
  }
}

describe('robustness', () => {
  test('a rejected change does not block the other changes', () => {
    const env = new TestEnv({workspace: workspaceWithTags(6)}).list('Tags');
    pauseTags(env, 6);
    env.fake.injectError({method: 'POST', path: /bulk_update/, times: 1000,
      match: req => String(req.rawBody || JSON.stringify(req.body)).includes('Tag 4'),
      code: 400, message: 'Invalid tag: Tag 4 is broken'});
    env.apply('Tags');
    for (let i = 1; i <= 6; i++) {
      const row = env.row('Tags', 'Tag Name', `Tag ${i}`);
      if (i === 4) {
        assert.match(env.result('Tags', row), /✖ Invalid tag: Tag 4 is broken/);
        assert.equal(env.entity('tag', '4').paused, undefined);
      } else {
        assert.match(env.result('Tags', row), /^✔ Updated: Paused/, `tag ${i}`);
        assert.equal(env.entity('tag', String(i)).paused, true);
      }
    }
    // The failed row is still detected as edited and can be applied again.
    env.fake.clearInjectedErrors();
    env.apply('Tags');
    assert.equal(env.entity('tag', '4').paused, true);
  });

  test('falls back to single requests when bulk_update is unavailable', () => {
    const env = new TestEnv({workspace: workspaceWithTags(3)}).list('Tags');
    pauseTags(env, 3);
    env.fake.injectError({method: 'POST', path: /bulk_update/, times: 1000,
                          code: 501, message: 'Method not found.'});
    env.apply('Tags');
    for (let i = 1; i <= 3; i++) {
      assert.equal(env.entity('tag', String(i)).paused, true);
    }
    assert.equal(env.fake.requestsTo('PUT', /\/tags\//).length, 3);
  });

  test('rows listed from another container cannot be applied', () => {
    const env = new TestEnv({workspace: workspaceWithTags(2)}).list('Tags');
    env.fake.seedWorkspace({containerId: '9000002', containerName: 'Other',
                            publicId: 'GTM-OTHER01', tags: []});
    env.click('listGtmContainers');
    const sheet = env.rt.spreadsheet.active.getSheetByName('GTM Workspace');
    const rows = env.values('GTM Workspace');
    rows.forEach((row, i) => {
      if (i >= 2 && row[3]) sheet.getRange(i + 1, 8).setValue(row[4] === 'GTM-OTHER01');
    });
    env.click('listGtmWorkspaces');
    assert.equal(env.rt.call('getWorkspaceContext').publicId, 'GTM-OTHER01');
    const row = env.row('Tags', 'Tag Name', 'Tag 1');
    env.set('Tags', row, {Paused: true});
    env.apply('Tags');
    assert.match(env.result('Tags', row), /listed from another container/);
  });

  test('API rate limits are handled with retries', () => {
    const env = new TestEnv({workspace: workspaceWithTags(3),
                             fake: {rateLimit: {max: 4, windowMs: 60000}}});
    env.list('Tags');
    pauseTags(env, 3);
    env.apply('Tags');
    for (let i = 1; i <= 3; i++) {
      assert.equal(env.entity('tag', String(i)).paused, true);
    }
    // Requests were rejected with 429 and retried successfully.
    assert.ok(env.fake.requests.some(r => r.status === 429));
  });

  test('the own rate limiter waits instead of exceeding the quota', () => {
    const env = new TestEnv({workspace: workspaceWithTags(2)});
    // 25 requests were made very recently.
    const now = env.rt.now();
    const times = Array.from({length: 25}, (_, i) => now - 1000 + i);
    env.rt.run(`CacheService.getUserCache().put('gtmApiCallTimes', ` +
        `${JSON.stringify(JSON.stringify(times))}, 600);`);
    env.list('Tags');
    assert.ok(env.rt.sleeps.totalMs >= 90000, String(env.rt.sleeps.totalMs));
  });

  test('work stops before the execution time limit', () => {
    const env = new TestEnv({workspace: workspaceWithTags(4)}).list('Tags');
    pauseTags(env, 4);
    env.rt = env.rt.nextExecution();
    env.rt.run('LIMITS.bulkUpdateChunkSize = 1;');
    // Every bulk update takes two minutes.
    const fake = env.fake;
    const rt = env.rt;
    rt.setFetchHandler(req => {
      if (/bulk_update/.test(req.url)) rt.advanceClock(120000);
      return fake.handle(req);
    });
    env.rt.spreadsheet.active.setActiveSheet(
        env.rt.spreadsheet.active.getSheetByName('Tags'));
    env.rt.call('applyActiveSheet');
    const results = [1, 2, 3, 4].map(i =>
      env.result('Tags', env.row('Tags', 'Tag Name', `Tag ${i}`)));
    assert.ok(results.some(r => /^✔/.test(r)));
    assert.ok(results.some(r => /time limit/.test(r)), JSON.stringify(results));
    assert.ok(results.every(r => /^✔|time limit/.test(r)));
  });

  test('API errors are shown clearly', () => {
    const env = new TestEnv({workspace: workspaceWithTags(1)});
    env.fake.injectError({method: 'GET', path: /\/tags$/, times: 1, code: 403,
                          message: 'The caller does not have permission'});
    env.click('listActiveSheet', 'Tags');
    assert.match(env.alertMessages()[0],
        /Tag Manager API error: Permission denied: The caller does not have permission/);
  });

  test('listing again asks before discarding unapplied changes', () => {
    const env = new TestEnv({workspace: workspaceWithTags(2)}).list('Tags');
    const row = env.row('Tags', 'Tag Name', 'Tag 1');
    env.set('Tags', row, {Paused: true});
    env.rt = env.rt.nextExecution();
    env.rt.ui.queueResponse('NO');
    env.rt.spreadsheet.active.setActiveSheet(
        env.rt.spreadsheet.active.getSheetByName('Tags'));
    env.rt.call('listActiveSheet');
    assert.equal(env.get('Tags', row, 'Paused'), true);
    assert.match(env.rt.ui.alerts.map(a => a.prompt).join(' '), /Cancelled/);
  });

  test('cancelling the confirmation applies nothing', () => {
    const env = new TestEnv({workspace: workspaceWithTags(1)}).list('Tags');
    env.set('Tags', env.row('Tags', 'Tag Name', 'Tag 1'), {Paused: true});
    env.rt = env.rt.nextExecution();
    env.rt.ui.queueResponse('NO');
    env.rt.spreadsheet.active.setActiveSheet(
        env.rt.spreadsheet.active.getSheetByName('Tags'));
    env.rt.call('applyActiveSheet');
    assert.equal(env.entity('tag', '1').paused, undefined);
  });

  test('an outdated sheet layout is rebuilt by List and blocks Apply', () => {
    const env = new TestEnv({workspace: workspaceWithTags(2)}).list('Tags');
    const sheet = env.rt.spreadsheet.active.getSheetByName('Tags');
    // Simulate a sheet from an older version (different columns).
    sheet.getRange(2, 5).setValue('Old column');
    env.set('Tags', env.row('Tags', 'Tag Name', 'Tag 1'), {Paused: true});
    env.apply('Tags');
    assert.match(env.alertMessages()[0], /do not match this version/);
    assert.equal(env.entity('tag', '1').paused, undefined);
    env.click('listActiveSheet', 'Tags');
    assert.equal(env.values('Tags')[1][4], 'Folder');
    assert.equal(env.values('Tags').length, 4);
  });

  test('Google tag containers: settings without workspaces', () => {
    const env = new TestEnv({
      preSeed: {containers: [{accountId: '6000001', containerId: '9000005',
        name: 'G-TAG container', publicId: 'GT-ABCDEFG',
        features: {supportTags: false, supportGtagConfigs: true,
                   supportBuiltInVariables: true}}]},
      workspace: {
        containerId: '9000005', workspaceId: '0',
        gtagConfigs: [{gtagConfigId: '1', type: 'ogt_session_timeout',
                       parameter: params({sessionMinutes: '30'})}]}});
    const ctx = env.rt.call('getWorkspaceContext');
    assert.equal(ctx.isGoogleTag, true);
    assert.equal(ctx.workspacePath, 'accounts/6000001/containers/9000005/workspaces/0');
    env.click('listActiveSheet', 'Tags');
    assert.match(env.alertMessages()[0], /works with web and server containers/);
    env.list('Google Tag Settings');
    const row = env.row('Google Tag Settings', 'Parameter Key', 'sessionMinutes');
    env.set('Google Tag Settings', row, {Value: '45'});
    env.apply('Google Tag Settings');
    assert.match(env.alertMessages().join(' '), /applied immediately/);
    assert.match(env.result('Google Tag Settings', row), /^✔ Updated/);
    assert.equal(paramValue(env.entity('gtagConfig', '1'), 'sessionMinutes'), '45');
    assert.equal(env.fake.requestsTo('POST', /bulk_update/).length, 0);
  });
});
