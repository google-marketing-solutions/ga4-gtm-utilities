/**
 * @fileoverview Convenience barrel: `const h = require('./harness');`
 */
'use strict';

const runtime = require('./appsScriptRuntime');
const fake = require('./fakeGtmApi');

module.exports = {
  createRuntime: runtime.createRuntime,
  resolveLiveToken: runtime.resolveLiveToken,
  collectSourceFiles: runtime.collectSourceFiles,
  createFakeGtmApi: fake.createFakeGtmApi,
  FakeGtmApi: fake.FakeGtmApi,
  GTM_API_BASE: fake.GTM_API_BASE,
  BUILT_IN_TRIGGER_IDS: fake.BUILT_IN_TRIGGER_IDS,
};
