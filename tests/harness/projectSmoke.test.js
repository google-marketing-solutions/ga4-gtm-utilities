/**
 * @fileoverview Smoke test against the REAL project sources (features/,
 * shared/, utils/): every file must load into the emulated runtime and the
 * onOpen() menu must only reference functions that exist. This catches
 * typos in menu function names, load-order problems (a top-level const used
 * before the file declaring it has loaded) and calls to Apps Script APIs the
 * harness does not provide.
 *
 * Set HARNESS_SKIP_PROJECT=1 to skip (e.g. while the sources are mid-refactor).
 */
'use strict';

const assert = require('node:assert/strict');
const {describe, it} = require('node:test');

const {createRuntime} = require('./appsScriptRuntime');
const {createFakeGtmApi} = require('./fakeGtmApi');

const skip = process.env.HARNESS_SKIP_PROJECT === '1' ? 'HARNESS_SKIP_PROJECT=1' : false;

describe('project smoke test (real sources)', {skip}, () => {
  it('loads every source file without errors', () => {
    const rt = createRuntime({loadErrors: 'collect', fetchHandler: createFakeGtmApi()});
    const errors = rt.loadErrors.map((e) => `${e.file}: ${e.error && e.error.message}`);
    assert.deepEqual(errors, []);
    assert.ok(rt.loadedFiles.length > 0, 'expected at least one source file under features/, shared/, utils/');
    assert.deepEqual(rt.harnessErrors, []);
  });

  it('onOpen() builds menus whose items all point at defined global functions', (t) => {
    const rt = createRuntime({fetchHandler: createFakeGtmApi()});
    if (!rt.isFunction('onOpen')) {
      t.skip('the project defines no onOpen()');
      return;
    }
    // Simple trigger semantics: AuthMode.LIMITED, so UrlFetchApp is off-limits.
    rt.callTrigger('onOpen');
    assert.ok(rt.ui.menus.length > 0, 'onOpen() should add at least one menu');
    assert.deepEqual(rt.missingMenuFunctions(), []);
  });
});
