/**
 * @fileoverview Self-tests for the Apps Script test harness: they prove the
 * mocks (SpreadsheetApp, Ui, services, UrlFetchApp), the fake GTM API and the
 * runtime loader behave as documented. Run with `node --test tests/` or
 * `node tests`.
 */
'use strict';

const assert = require('node:assert/strict');
const path = require('node:path');
const {spawnSync} = require('node:child_process');
const {describe, it} = require('node:test');

const {createRuntime, resolveLiveToken} = require('./appsScriptRuntime');
const {createFakeGtmApi, BUILT_IN_TRIGGER_IDS} = require('./fakeGtmApi');

const MINI_PROJECT = path.join(__dirname, 'fixtures', 'miniProject');
const API = 'https://tagmanager.googleapis.com/tagmanager/v2/';

/** Script-side helpers used by the fake GTM API tests. */
const GTM_HELPERS = `
  const API = '${API}';
  function api(method, path, body, extra) {
    const params = Object.assign({method: method, muteHttpExceptions: true,
      headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}}, extra || {});
    if (body !== undefined) {
      params.contentType = 'application/json';
      params.payload = JSON.stringify(body);
    }
    const res = UrlFetchApp.fetch(API + path, params);
    const text = res.getContentText();
    return {code: res.getResponseCode(), body: text ? JSON.parse(text) : null, text: text};
  }
  function listAll(path, key) {
    let items = [];
    let token;
    let pages = 0;
    do {
      const r = api('get', path + (token ? '?pageToken=' + encodeURIComponent(token) : ''));
      if (r.code !== 200) throw new Error('HTTP ' + r.code + ': ' + r.text);
      items = items.concat(r.body[key] || []);
      token = r.body.nextPageToken;
      pages++;
    } while (token);
    return {items: items, pages: pages};
  }
`;

/**
 * Runtime with no project files, optional extra sources.
 * @param {!Object=} options
 * @return {!Object}
 */
function bareRuntime(options = {}) {
  return createRuntime({load: false, manifest: null, ...options});
}

/**
 * Runtime + fake API with the script-side GTM helpers loaded.
 * @param {!Object=} fakeOptions
 * @param {!Object=} runtimeOptions
 * @return {{rt: !Object, fake: !Object}}
 */
function gtmRuntime(fakeOptions = {}, runtimeOptions = {}) {
  const fake = createFakeGtmApi(fakeOptions);
  const rt = createRuntime({
    load: false, manifest: null, fetchHandler: fake,
    sources: [{filename: 'gtmHelpers.js', code: GTM_HELPERS}],
    ...runtimeOptions,
  });
  return {rt, fake};
}

/**
 * Returns the thrown error (from host code calling into the runtime).
 * @param {function()} fn
 * @return {!Error}
 */
function catchError(fn) {
  try {
    fn();
  } catch (e) {
    return e;
  }
  assert.fail('expected an error to be thrown');
}

/** @return {boolean} Whether googleapis.com is reachable (for real HTTP tests). */
function isOnline() {
  const r = spawnSync('curl', ['-sS', '-o', '/dev/null', '-w', '%{http_code}', '--max-time', '10',
    'https://tagmanager.googleapis.com/$discovery/rest?version=v2'], {encoding: 'utf8'});
  return r.status === 0 && /^[1-5]\d\d$/.test(String(r.stdout).trim());
}

// ===========================================================================
describe('runtime: loading and the shared global scope', () => {
  it('loads features/, shared/, utils/ in path order with one shared scope', () => {
    const rt = createRuntime({rootDir: MINI_PROJECT});
    assert.deepEqual(rt.loadedFiles, ['features/featureA.js', 'shared/config.js', 'utils/helpers.js']);
    // Function declarations are globals (properties of the context).
    assert.equal(typeof rt.context.featureA, 'function');
    // Top-level const is visible to other files but is NOT a context property.
    assert.equal(rt.context.SHARED_CONFIG, undefined);
    assert.deepEqual(rt.get('SHARED_CONFIG'), {prefix: 'cfg:'});
    assert.equal(rt.get('DERIVED'), 'cfg:derived');
    assert.equal(rt.get('SHARED_CONFIG.prefix'), 'cfg:');
    assert.equal(rt.call('featureA'), 'cfg:[A]');
    assert.equal(rt.call('helperArrow', 'x'), '[x]', 'const arrow functions are callable');
    assert.deepEqual(rt.call('featureCounter'), [1, 2]);
    assert.deepEqual(rt.call('featureCounter'), [3, 4], 'top-level let state persists');
    assert.equal(rt.has('nope'), false);
    assert.equal(rt.isFunction('bump'), true);
    assert.deepEqual(rt.globalFunctionNames(), ['bump', 'featureA', 'featureCounter', 'helperC']);
  });

  it('honours options.files order and inline sources', () => {
    const rt = createRuntime({
      rootDir: MINI_PROJECT, manifest: null,
      files: ['shared/config.js'],
      sources: [{filename: 'extra.js', code: 'function extra() { return SHARED_CONFIG.prefix + "x"; }'}],
    });
    assert.deepEqual(rt.loadedFiles, ['shared/config.js', 'extra.js']);
    assert.equal(rt.call('extra'), 'cfg:x');
  });

  it('reports load errors like Apps Script (undefined refs, duplicate const)', () => {
    const rt = bareRuntime({
      loadErrors: 'collect',
      sources: [
        {filename: 'a.js', code: 'const X = Y + 1;'},
        {filename: 'b.js', code: 'const Y = 1;'},
      ],
    });
    assert.equal(rt.loadErrors.length, 1);
    assert.equal(rt.loadErrors[0].file, 'a.js');
    assert.match(rt.loadErrors[0].error.message, /Y is not defined/);
    const err = catchError(() => bareRuntime({
      sources: [{filename: 'one.js', code: 'const DUP = 1;'}, {filename: 'two.js', code: 'const DUP = 2;'}],
    }));
    assert.match(err.message, /Failed to load two\.js: .*already been declared/);
  });

  it('converts arguments and results across the realm boundary', () => {
    const rt = bareRuntime({sources: [{filename: 's.js', code: `
      function inspect(arr, obj, date) {
        return {isArray: arr instanceof Array, isObject: obj instanceof Object, isDate: date instanceof Date,
          sum: arr.reduce((a, b) => a + b, 0), when: date};
      }
      function makeRows() { return [[1, 'a'], [2, 'b']]; }
    `}]});
    const when = new Date('2024-05-06T07:08:09Z');
    const out = rt.call('inspect', [1, 2, 3], {a: 1}, when);
    assert.deepEqual(out, {isArray: true, isObject: true, isDate: true, sum: 6, when});
    assert.deepEqual(rt.call('makeRows'), [[1, 'a'], [2, 'b']]);
    const raw = rt.callRaw('makeRows');
    assert.ok(Array.isArray(raw));
    assert.notEqual(Object.getPrototypeOf(raw), Array.prototype, 'callRaw returns context-realm arrays');
  });

  it('throws service errors as context-realm Errors named Exception', () => {
    const rt = bareRuntime();
    const result = rt.run(`(function () {
      try { SpreadsheetApp.getActive().getActiveSheet().getRange(0, 1); }
      catch (e) { return [e instanceof Error, e.name, e.message, String(e)]; }
    })()`);
    assert.deepEqual(rt.toHost(result), [true, 'Exception', 'The starting row of the range is too small.',
      'Exception: The starting row of the range is too small.']);
  });

  it('installs the TagManager advanced service only when the manifest enables it', () => {
    const fake = createFakeGtmApi();
    const {workspacePath} = fake.seedWorkspace({tags: [{name: 'T1', type: 'html'}]});
    const rt = createRuntime({rootDir: MINI_PROJECT, fetchHandler: fake});
    assert.equal(rt.run('typeof TagManager'), 'object');
    const tags = rt.toHost(rt.run(`TagManager.Accounts.Containers.Workspaces.Tags.list('${workspacePath}')`));
    assert.deepEqual(tags.tag.map((t) => t.name), ['T1']);
    const err = rt.toHost(rt.run(`(function () {
      try { TagManager.Accounts.Containers.Workspaces.Tags.get('${workspacePath}/tags/999'); }
      catch (e) { return {name: e.name, message: e.message, code: e.details.code}; }
    })()`));
    assert.deepEqual(err, {name: 'GoogleJsonResponseException', code: 404,
      message: 'API call to tagmanager.accounts.containers.workspaces.tags.get failed with error: ' +
        'Not found or permission denied.'});
    assert.equal(rt.fetchLog.length, 2, 'advanced service calls go through UrlFetchApp transport');
    const off = createRuntime({rootDir: MINI_PROJECT, advancedServices: false});
    assert.equal(off.run('typeof TagManager'), 'undefined');
    assert.equal(rt.run('Session.getScriptTimeZone()'), 'Europe/Berlin', 'time zone from manifest');
  });

  it('keeps a virtual clock: sleep advances Date.now() without waiting', () => {
    const rt = bareRuntime();
    const started = Date.now();
    rt.run('var t0 = Date.now(); Utilities.sleep(120000); Utilities.sleep(30000); var t1 = Date.now();');
    assert.ok(Date.now() - started < 1000, 'did not really sleep');
    assert.ok(rt.get('t1') - rt.get('t0') >= 150000);
    assert.deepEqual(rt.sleeps, {totalMs: 150000, calls: [120000, 30000]});
    assert.equal(rt.run('new Date() instanceof Date && typeof Date() === "string"'), true);
    assert.equal(catchError(() => rt.run('Utilities.sleep(300001)')).name, 'Exception');
  });

  it('can guard against runaway scripts with timeoutMs', () => {
    const rt = bareRuntime({timeoutMs: 200, sources: [{filename: 'spin.js', code: 'function spin() { for (;;) {} }'}]});
    assert.match(catchError(() => rt.call('spin')).message, /timed out/);
  });

  it('runs simple triggers in AuthMode.LIMITED (no UrlFetchApp)', () => {
    const rt = bareRuntime({fetchHandler: () => ({code: 200, body: 'ok'}), sources: [{filename: 't.js', code: `
      function onOpen(e) { return [e.authMode, e.source.getName(), e.user.getEmail()]; }
      function onOpenThatFetches() { UrlFetchApp.fetch('https://e.test/'); }
    `}]});
    assert.deepEqual(rt.callTrigger('onOpen'), ['LIMITED', 'Test Spreadsheet', 'test.user@example.com']);
    assert.equal(catchError(() => rt.callTrigger('onOpenThatFetches')).message,
        'You do not have permission to call UrlFetchApp.fetch. Required permissions: ' +
        'https://www.googleapis.com/auth/script.external_request');
    assert.equal(rt.run('UrlFetchApp.fetch("https://e.test/").getContentText()'), 'ok', 'FULL again afterwards');
    assert.deepEqual(rt.callTrigger('onOpen', {}, {authMode: 'FULL'})[0], 'FULL');
  });

  it('nextExecution() re-runs top-level code but keeps persistent state', () => {
    const fake = createFakeGtmApi();
    const rt1 = bareRuntime({fetchHandler: fake, sources: [{filename: 's.js', code: `
      let runs = 0;
      function step() {
        runs++;
        PropertiesService.getDocumentProperties().setProperty('p', String(runs));
        CacheService.getUserCache().put('c', 'v', 600);
        SpreadsheetApp.getActive().getActiveSheet().getRange('A1').setValue('run ' + runs).setBackground('#00ff00');
        return runs;
      }
    `}]});
    assert.equal(rt1.call('step'), 1);
    rt1.run('Utilities.sleep(5000)');
    rt1.ui.queueResponse('NO');
    const rt2 = rt1.nextExecution();
    assert.deepEqual(rt2.spreadsheet.getSheetValues('Sheet1'), [['run 1']]);
    assert.equal(rt2.spreadsheet.getCell('Sheet1', 'A1').format.background, '#00ff00');
    assert.equal(rt2.properties.document.get('p'), '1');
    assert.equal(rt2.run('CacheService.getUserCache().get("c")'), 'v');
    assert.ok(rt2.now() >= rt1.now(), 'virtual clock carried over');
    assert.equal(rt2.run('SpreadsheetApp.getUi().alert("Q", SpreadsheetApp.getUi().ButtonSet.YES_NO)'), 'NO');
    assert.equal(rt2.call('step'), 1, 'globals are fresh in the next execution');
    assert.equal(rt2.options.fetchHandler, fake, 'same fake backend');
  });
});

// ===========================================================================
describe('SpreadsheetApp mock', () => {
  it('manages sheets like Apps Script', () => {
    const rt = bareRuntime({spreadsheet: {sheets: {Data: [['a']], Other: []}}});
    const ss = rt.spreadsheet.active;
    assert.equal(rt.run('SpreadsheetApp.getActive() === SpreadsheetApp.getActiveSpreadsheet()'), true);
    assert.deepEqual(rt.spreadsheet.sheetNames(), ['Data', 'Other']);
    const created = rt.run('SpreadsheetApp.getActive().insertSheet("New", 0)');
    assert.equal(created.getIndex(), 1);
    assert.equal(ss.getActiveSheet().getName(), 'New', 'inserted sheet becomes active');
    const dup = catchError(() => rt.run('SpreadsheetApp.getActive().insertSheet("data")'));
    assert.equal(dup.message, 'A sheet with the name "data" already exists. Please enter another name.');
    assert.equal(rt.run('SpreadsheetApp.getActive().getSheetByName("Missing")'), null);
    rt.run('SpreadsheetApp.getActive().deleteSheet(SpreadsheetApp.getActive().getSheetByName("Other"))');
    assert.deepEqual(rt.spreadsheet.sheetNames(), ['New', 'Data']);
    assert.equal(rt.run('SpreadsheetApp.getActive().getSheets().length'), 2);
    assert.match(ss.getUrl(), /^https:\/\/docs\.google\.com\/spreadsheets\/d\/test-spreadsheet-id\/edit$/);
  });

  it('resolves numeric and A1 ranges', () => {
    const rt = bareRuntime({spreadsheet: {sheets: {'My Sheet': [[1, 2, 3], [4, 5, 6]]}}});
    const sheet = rt.spreadsheet.getSheet('My Sheet');
    assert.equal(sheet.getRange(2, 2).getA1Notation(), 'B2');
    assert.equal(sheet.getRange(1, 1, 2, 3).getA1Notation(), 'A1:C2');
    assert.deepEqual(rt.toHost(sheet.getRange('B1:C2').getValues()), [[2, 3], [5, 6]]);
    const col = sheet.getRange('A:A');
    assert.equal(col.getA1Notation(), 'A:A');
    assert.equal(col.getNumRows(), 1000);
    assert.equal(sheet.getRange('A2:C').getNumRows(), 999);
    assert.deepEqual(rt.toHost(rt.run('SpreadsheetApp.getActive().getRange("\'My Sheet\'!C2").getValue()')), 6);
    assert.equal(catchError(() => sheet.getRange('nonsense!!')).message, 'Range not found');
    assert.equal(catchError(() => sheet.getRange(1, 1, 0, 1)).message,
        'The number of rows in the range must be at least 1.');
    assert.equal(catchError(() => rt.run('SpreadsheetApp.getActive().getSheets()[0].getRange(undefined, 1)')).message,
        'The parameters (null,number) don\'t match the method signature for SpreadsheetApp.Sheet.getRange.');
    assert.equal(sheet.getRange(1, 1).offset(1, 2).getA1Notation(), 'C2');
    assert.equal(sheet.getRange('A1:C2').getCell(2, 3).getValue(), 6);
  });

  it('validates setValues dimensions exactly like Apps Script', () => {
    const rt = bareRuntime();
    const run = (code) => catchError(() => rt.run(`SpreadsheetApp.getActive().getActiveSheet().${code}`)).message;
    assert.equal(run('getRange(1, 1, 3, 2).setValues([[1, 2], [3, 4]])'),
        'The number of rows in the data does not match the number of rows in the range. The data has 2 but the range has 3.');
    assert.equal(run('getRange(1, 1, 2, 3).setValues([[1, 2, 3], [4, 5]])'),
        'The number of columns in the data does not match the number of columns in the range. The data has 2 but the range has 3.');
    assert.equal(run('getRange(1, 1, 1, 2).setValues([1, 2])'),
        'The parameters (number[]) don\'t match the method signature for SpreadsheetApp.Range.setValues.');
    assert.equal(run('getRange(1, 1, 1, 2).setBackgrounds([["red"]])'),
        'The number of columns in the data does not match the number of columns in the range. The data has 1 but the range has 2.');
  });

  it('auto-types strings like the Sheets UI unless the cell is plain text', () => {
    const rt = bareRuntime();
    const when = new Date('2024-01-02T03:04:05Z');
    rt.run(`var sh = SpreadsheetApp.getActive().getActiveSheet();
      sh.getRange(2, 1, 1, 3).setNumberFormat('@');
      sh.getRange(1, 1, 1, 12).setValues([['123', '1.5', '1,234', '5%', 'TRUE', 'false', '2024-01-01',
        '00123', 'GTM-AB12', "'123", ' 42 ', '']]);
      sh.getRange(2, 1, 1, 3).setValues([['123', 'TRUE', '007']]);`);
    rt.globals.SpreadsheetApp.getActive().getActiveSheet().getRange(3, 1, 1, 4).setValues([[7, true, when, null]]);
    const values = rt.spreadsheet.getSheetValues('Sheet1');
    assert.equal(values[0].length, 11, 'the data range ends at the last non-empty column');
    assert.deepEqual(values[0], [123, 1.5, 1234, 0.05, true, false, '2024-01-01', 123, 'GTM-AB12', '123', 42]);
    assert.deepEqual(values[1].slice(0, 3), ['123', 'TRUE', '007'], 'plain text (@) cells keep strings');
    assert.deepEqual(values[2].slice(0, 4), [7, true, when, '']);
    assert.equal(rt.spreadsheet.getCell('Sheet1', 'D1').format.numberFormat, '0%');
    assert.equal(rt.run('SpreadsheetApp.getActive().getActiveSheet().getRange("C3").getValue() instanceof Date'), true);
  });

  it('tracks content extents and grows the grid on write', () => {
    const rt = bareRuntime();
    const sheet = rt.spreadsheet.getSheet('Sheet1');
    assert.equal(sheet.getLastRow(), 0);
    assert.equal(sheet.getDataRange().getA1Notation(), 'A1');
    sheet.getRange(1200, 30).setValue('far');
    assert.equal(sheet.getLastRow(), 1200);
    assert.equal(sheet.getLastColumn(), 30);
    assert.equal(sheet.getMaxRows(), 1200);
    assert.equal(sheet.getMaxColumns(), 30);
    sheet.getRange(5, 2).setBackground('red');
    assert.equal(sheet.getLastRow(), 1200, 'formats do not count as content');
    sheet.getRange(1200, 30).clearContent();
    assert.equal(sheet.getLastRow(), 0);
    const strict = bareRuntime({strictBounds: true});
    assert.equal(catchError(() => strict.spreadsheet.getSheet('Sheet1').getRange(1001, 1)).message,
        'The coordinates of the range are outside the dimensions of the sheet.');
  });

  it('computes display values', () => {
    const rt = bareRuntime({spreadsheet: {timeZone: 'UTC'}});
    const sheet = rt.spreadsheet.getSheet('Sheet1');
    sheet.getRange(1, 1, 1, 5).setValues([[true, 1.5, 'x', new Date(Date.UTC(2024, 0, 2, 13, 4, 5)), 1234.5]]);
    sheet.getRange('E1').setNumberFormat('#,##0.00');
    assert.deepEqual(rt.toHost(sheet.getRange('A1:E1').getDisplayValues()),
        [['TRUE', '1.5', 'x', '1/2/2024 13:04:05', '1,234.50']]);
  });

  it('stores formats, notes and clears them selectively', () => {
    const rt = bareRuntime();
    rt.run(`var r = SpreadsheetApp.getActive().getActiveSheet().getRange('A1:B2');
      r.setValues([[1, 2], [3, 4]]).setBackground('#FF0000').setFontWeight('bold').setNote('hi');
      r.setFontColors([['blue', null], ['#abc', 'black']]);`);
    const sheet = rt.spreadsheet.getSheet('Sheet1');
    assert.equal(sheet.getRange('A1').getBackground(), '#ff0000');
    assert.deepEqual(rt.toHost(sheet.getRange('A1:B2').getFontColors()), [['#0000ff', '#000000'], ['#aabbcc', '#000000']]);
    assert.equal(sheet.getRange('B2').getNote(), 'hi');
    sheet.getRange('A1:B2').clearContent();
    assert.equal(sheet.getRange('A1').getValue(), '');
    assert.equal(sheet.getRange('A1').getFontWeight(), 'bold', 'clearContent keeps formats');
    sheet.getRange('A1:B2').clearFormat();
    assert.equal(sheet.getRange('A1').getBackground(), '#ffffff');
    assert.equal(sheet.getRange('A1').getNote(), 'hi', 'clearFormat keeps notes');
    sheet.getRange('A1:B2').clearNote();
    assert.equal(sheet.getRange('A1').getNote(), '');
  });

  it('builds data validation rules', () => {
    const rt = bareRuntime();
    const rule = rt.run(`SpreadsheetApp.newDataValidation()
      .requireValueInList(['Create', 'Update', 'Delete'], true).setHelpText('Pick one').build()`);
    assert.equal(rule.getCriteriaType(), 'VALUE_IN_LIST');
    assert.equal(rule.getCriteriaType(), rt.run('SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST'));
    assert.deepEqual(rt.toHost(rule.getCriteriaValues()), [['Create', 'Update', 'Delete'], true]);
    assert.equal(rule.getAllowInvalid(), true, 'allowInvalid defaults to true');
    assert.equal(rule.getHelpText(), 'Pick one');
    const strict = rule.copy().setAllowInvalid(false).build();
    assert.equal(strict.getAllowInvalid(), false);
    assert.equal(rule.getAllowInvalid(), true, 'copy() does not mutate the original');
    const sheet = rt.spreadsheet.getSheet('Sheet1');
    sheet.getRange('A1:A3').setDataValidation(strict);
    assert.equal(sheet.getRange('A2').getDataValidation().getCriteriaType(), 'VALUE_IN_LIST');
    assert.equal(rt.toHost(sheet.getRange('A1:B1').getDataValidations())[0][1], null);
    sheet.getRange('A1').setValue('Update');
    assert.equal(catchError(() => sheet.getRange('A2').setValue('Nope')).message,
        'The data you entered in cell A2 violates the data validation rules set on this cell.');
    const range = rt.run(`(function () {
      const sh = SpreadsheetApp.getActive().getActiveSheet();
      sh.getRange('C1:C2').setValues([['x'], ['y']]);
      return SpreadsheetApp.newDataValidation().requireValueInRange(sh.getRange('C1:C2'), false).build();
    })()`);
    assert.equal(range.getCriteriaType(), 'VALUE_IN_RANGE');
    assert.equal(range.getCriteriaValues()[0].getA1Notation(), 'C1:C2');
    sheet.getRange('A1:A3').clearDataValidations();
    assert.equal(sheet.getRange('A2').getDataValidation(), null);
  });

  it('implements checkbox semantics', () => {
    const rt = bareRuntime();
    const sheet = rt.spreadsheet.getSheet('Sheet1');
    sheet.getRange('A1:A3').setValues([['TRUE'], ['x'], ['']]);
    sheet.getRange('A1:A3').insertCheckboxes();
    assert.deepEqual(rt.spreadsheet.getSheetValues('Sheet1'), [[false], [false], [false]],
        'insertCheckboxes sets every cell to false (documented behaviour)');
    assert.equal(sheet.getRange('A1').getDataValidation().getCriteriaType(), 'CHECKBOX');
    sheet.getRange('A1').setValue('true');
    sheet.getRange('A2').setValue(true);
    assert.equal(sheet.getRange('A1:A2').isChecked(), true);
    sheet.getRange('A1').setValue('');
    assert.equal(sheet.getRange('A1').getValue(), false, 'clearing a checkbox unchecks it');
    sheet.getRange('A1:A3').check();
    assert.deepEqual(rt.spreadsheet.getSheetValues('Sheet1'), [[true], [true], [true]]);
    sheet.getRange('A3').uncheck();
    assert.equal(sheet.getRange('A1:A3').isChecked(), null);
    sheet.getRange('A1:A3').clearContent();
    assert.deepEqual(rt.spreadsheet.getSheetValues('Sheet1'), [[false], [false], [false]]);
    sheet.getRange('A1:A3').removeCheckboxes();
    assert.deepEqual(rt.spreadsheet.getSheetValues('Sheet1'), []);
    assert.equal(sheet.getRange('A1').getDataValidation(), null);
    rt.run(`SpreadsheetApp.getActive().getActiveSheet().getRange('B1:B2')
      .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build())`);
    assert.deepEqual(rt.toHost(sheet.getRange('B1:B2').getValues()), [[false], [false]]);
    assert.equal(sheet.getRange('B1').getDisplayValue(), 'FALSE');
  });

  it('stores formulas and evaluates HYPERLINK', () => {
    const rt = bareRuntime();
    const sheet = rt.spreadsheet.getSheet('Sheet1');
    sheet.getRange('A1').setValue('=hyperlink("https://tagmanager.google.com/#/x","My Tag")');
    assert.equal(sheet.getRange('A1').getValue(), 'My Tag');
    assert.equal(sheet.getRange('A1').getFormula(), '=hyperlink("https://tagmanager.google.com/#/x","My Tag")');
    assert.equal(rt.spreadsheet.getCell('Sheet1', 'A1').link, 'https://tagmanager.google.com/#/x');
    sheet.getRange('B1').setFormula('=SUM(A2:A9)');
    assert.equal(sheet.getRange('B1').getValue(), '');
    assert.equal(sheet.getLastColumn(), 2, 'formula cells count as content');
    assert.equal(rt.warnings.at(-1).code, 'formula-not-evaluated');
    const custom = bareRuntime({formulaEvaluator: (f) => (f === '=1+1' ? 2 : undefined)});
    custom.spreadsheet.getSheet('Sheet1').getRange('A1').setValue('=1+1');
    assert.equal(custom.spreadsheet.getSheet('Sheet1').getRange('A1').getValue(), 2);
  });

  it('supports rich text values', () => {
    const rt = bareRuntime();
    rt.run(`var v = SpreadsheetApp.newRichTextValue().setText('Open tag').setLinkUrl('https://x.test/t').build();
      SpreadsheetApp.getActive().getActiveSheet().getRange('A1').setRichTextValue(v);`);
    const sheet = rt.spreadsheet.getSheet('Sheet1');
    assert.equal(sheet.getRange('A1').getValue(), 'Open tag');
    assert.equal(sheet.getRange('A1').getRichTextValue().getLinkUrl(), 'https://x.test/t');
    assert.equal(sheet.getRange('A1').getRichTextValue().getText(), 'Open tag');
    sheet.getRange('B1').setValue(5);
    assert.equal(sheet.getRange('B1').getRichTextValue(), null, 'non-text cells have no rich text');
  });

  it('inserts/deletes rows and columns, inheriting formats and validation', () => {
    const rt = bareRuntime();
    const sheet = rt.spreadsheet.getSheet('Sheet1');
    sheet.getRange('A1:B3').setValues([['h1', 'h2'], ['a', 'b'], ['c', 'd']]);
    sheet.getRange('B2').insertCheckboxes();
    sheet.getRange('A2').setBackground('#00ff00');
    sheet.insertRowsAfter(2, 2);
    assert.deepEqual(rt.spreadsheet.getSheetValues('Sheet1'), [['h1', 'h2'], ['a', false], ['', false], ['', false], ['c', 'd']]);
    assert.equal(sheet.getRange('A3').getBackground(), '#00ff00');
    assert.equal(sheet.getRange('B4').getDataValidation().getCriteriaType(), 'CHECKBOX');
    assert.equal(sheet.getMaxRows(), 1002);
    sheet.deleteRows(3, 2);
    assert.deepEqual(rt.spreadsheet.getSheetValues('Sheet1'), [['h1', 'h2'], ['a', false], ['c', 'd']]);
    sheet.insertColumnBefore(1);
    assert.deepEqual(rt.spreadsheet.getSheetValues('Sheet1')[0], ['', 'h1', 'h2']);
    sheet.deleteColumn(1);
    assert.equal(catchError(() => sheet.deleteRows(1, 1000)).message,
        'Sorry, it is not possible to delete all non-frozen rows.');
    assert.equal(catchError(() => sheet.deleteRows(999, 5)).message, 'Those rows are out of bounds.');
  });

  it('records layout: hidden columns, frozen rows, widths, merges, tabs', () => {
    const rt = bareRuntime();
    const sheet = rt.spreadsheet.getSheet('Sheet1');
    sheet.hideColumns(2, 2);
    assert.equal(sheet.isColumnHiddenByUser(3), true);
    sheet.showColumns(3);
    assert.equal(sheet.isColumnHiddenByUser(3), false);
    assert.equal(catchError(() => sheet.hideColumns(26, 2)).message, 'Those columns are out of bounds.');
    sheet.setFrozenRows(2);
    sheet.setFrozenColumns(1);
    assert.equal(sheet.getFrozenRows(), 2);
    sheet.setColumnWidths(1, 3, 250);
    assert.equal(sheet.getColumnWidth(2), 250);
    assert.equal(sheet.getColumnWidth(9), 100);
    sheet.setRowHeightsForced(1, 2, 40);
    assert.equal(sheet.getRowHeight(2), 40);
    sheet.getRange('A1:B1').setValues([['keep', 'drop']]).merge();
    assert.equal(sheet.getRange('A1').isPartOfMerge(), true);
    assert.equal(sheet.getRange('B1').getValue(), '', 'merging keeps only the top-left value');
    assert.equal(catchError(() => sheet.getRange('B1:C1').merge()).message,
        'You must select all cells in a merged range to merge or unmerge them.');
    sheet.getRange('A1:B1').breakApart();
    assert.equal(sheet.getRange('A1').isPartOfMerge(), false);
    sheet.setTabColor('#FF9900');
    assert.equal(sheet.getTabColor(), '#ff9900');
    rt.run('SpreadsheetApp.getActive().insertSheet("Two")');
    sheet.hideSheet();
    assert.equal(sheet.isSheetHidden(), true);
  });

  it('appends and sorts rows', () => {
    const rt = bareRuntime();
    const sheet = rt.spreadsheet.getSheet('Sheet1');
    sheet.appendRow(['name', 'n']);
    sheet.appendRow(['b', '2']);
    sheet.appendRow(['a', '10']);
    sheet.appendRow(['c', '']);
    sheet.setFrozenRows(1);
    sheet.sort(2, false);
    assert.deepEqual(rt.spreadsheet.getSheetValues('Sheet1'), [['name', 'n'], ['a', 10], ['b', 2], ['c', '']]);
    sheet.getRange('A2:B4').sort({column: 1, ascending: true});
    assert.deepEqual(rt.spreadsheet.getSheetValues('Sheet1').map((r) => r[0]), ['name', 'a', 'b', 'c']);
  });

  it('provides fixture helpers: addSheet, dumpSheet, toJSON/loadFixture', () => {
    const rt = bareRuntime({sources: [{filename: 'top.js', code: `
      const ss = SpreadsheetApp.getActive();
      function sheetNamesFromScript() { return ss.getSheets().map(s => s.getName()); }
    `}]});
    rt.spreadsheet.addSheet('Raw', [['123', true], ['x']]);
    assert.deepEqual(rt.spreadsheet.getSheetValues('Raw'), [['123', true], ['x', '']], 'raw by default');
    rt.spreadsheet.addSheet('Typed', [['123', 'TRUE']], {autoType: true, frozenRows: 1});
    assert.deepEqual(rt.spreadsheet.getSheetValues('Typed'), [[123, true]]);
    const sheet = rt.spreadsheet.getSheet('Typed');
    sheet.getRange('C2:C5').insertCheckboxes();
    sheet.getRange('A1:B1').setBackground('#cccccc').setFontWeight('bold');
    sheet.getRange('A1').setNote('header');
    const dump = rt.spreadsheet.dumpSheet('Typed');
    assert.deepEqual(dump.validations, {'C2:C5': {criteria: 'CHECKBOX', values: [], allowInvalid: true, helpText: ''}});
    assert.deepEqual(dump.formats, {'A1:B1': {background: '#cccccc', fontWeight: 'bold'}});
    assert.deepEqual(dump.notes, {A1: 'header'});
    assert.equal(dump.frozenRows, 1);
    assert.match(rt.spreadsheet.printSheet('Typed'), /\| A\s+\| B\s+\| C/);

    const json = JSON.parse(JSON.stringify(rt.spreadsheet.toJSON()));
    const other = bareRuntime({spreadsheet: {fixture: json}});
    assert.deepEqual(other.spreadsheet.toJSON(), json, 'toJSON/loadFixture round-trip');

    rt.spreadsheet.loadFixture({sheets: [{name: 'Only', values: [[{$date: '2024-01-01T00:00:00.000Z'}]]}]});
    assert.deepEqual(rt.call('sheetNamesFromScript'), ['Only'], 'captured `ss` stays valid after loadFixture');
    assert.ok(rt.spreadsheet.getSheetValues('Only')[0][0] instanceof Date);
  });
});

// ===========================================================================
describe('Ui mock', () => {
  it('returns default responses per button set and records alerts', () => {
    const rt = bareRuntime();
    const out = rt.toHost(rt.run(`(function () {
      const ui = SpreadsheetApp.getUi();
      return [ui.alert('Hello'), ui.alert('T', 'Sure?', ui.ButtonSet.YES_NO),
        ui.alert('Go?', ui.ButtonSet.OK_CANCEL), ui.alert('T', 'Q', ui.ButtonSet.YES_NO_CANCEL)];
    })()`));
    assert.deepEqual(out, ['OK', 'YES', 'OK', 'YES']);
    assert.deepEqual(rt.ui.alerts.map((a) => [a.title, a.prompt, a.buttons]),
        [[null, 'Hello', 'OK'], ['T', 'Sure?', 'YES_NO'], [null, 'Go?', 'OK_CANCEL'], ['T', 'Q', 'YES_NO_CANCEL']]);
  });

  it('consumes scripted responses for alert() and prompt()', () => {
    const rt = bareRuntime({uiResponses: ['NO']});
    rt.ui.queueResponse({text: 'GTM-ABC123'}, {text: 'ignored', button: 'CANCEL'}, (dialog) =>
      (dialog.type === 'alert' ? 'CLOSE' : 'x'));
    const out = rt.toHost(rt.run(`(function () {
      const ui = SpreadsheetApp.getUi();
      const a = ui.alert('Delete?', ui.ButtonSet.YES_NO);
      const p1 = ui.prompt('Container ID?');
      const p2 = ui.prompt('Title', 'Name?', ui.ButtonSet.OK_CANCEL);
      const b = ui.alert('Closed?', ui.ButtonSet.OK);
      return [a == ui.Button.NO, p1.getResponseText(), p1.getSelectedButton(), p2.getSelectedButton(), b];
    })()`));
    assert.deepEqual(out, [true, 'GTM-ABC123', 'OK', 'CANCEL', 'CLOSE']);
    assert.equal(rt.ui.prompts[0].text, 'GTM-ABC123');
    rt.ui.queueResponse('YES');
    assert.match(catchError(() => rt.run('SpreadsheetApp.getUi().alert("x")')).message,
        /not valid for ButtonSet\.OK/);
  });

  it('records menus and finds menu items without a backing function', () => {
    const rt = bareRuntime({sources: [{filename: 'menu.js', code: `
      function onOpen() {
        const ui = SpreadsheetApp.getUi();
        ui.createMenu('Tools')
          .addItem('Run', 'runThing')
          .addSeparator()
          .addSubMenu(ui.createMenu('More').addItem('Missing', 'notDefinedAnywhere'))
          .addToUi();
      }
      function runThing() { SpreadsheetApp.getActive().toast('done', 'Tools', 3); }
    `}]});
    rt.call('onOpen');
    assert.deepEqual(rt.ui.menuItems(), [
      {path: 'Tools > Run', caption: 'Run', functionName: 'runThing'},
      {path: 'Tools > More > Missing', caption: 'Missing', functionName: 'notDefinedAnywhere'},
    ]);
    assert.deepEqual(rt.missingMenuFunctions(), ['notDefinedAnywhere']);
    rt.call(rt.ui.findMenuItem('Run').functionName);
    assert.deepEqual(rt.ui.toasts, [{type: 'toast', message: 'done', title: 'Tools', timeoutSeconds: 3}]);
  });
});

// ===========================================================================
describe('Apps Script services', () => {
  it('Utilities: digests, base64, uuid, formatString', () => {
    const rt = bareRuntime();
    const hex = (bytes) => rt.toHost(bytes).map((b) => (b & 0xff).toString(16).padStart(2, '0')).join('');
    const md5 = rt.run('Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, "")');
    assert.ok(rt.toHost(md5).every((b) => b >= -128 && b <= 127), 'signed bytes');
    assert.equal(rt.toHost(md5)[0], -44);
    assert.equal(hex(md5), 'd41d8cd98f00b204e9800998ecf8427e');
    assert.equal(hex(rt.run('Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, "abc")')),
        'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad');
    assert.equal(rt.run('Utilities.base64Encode("hello")'), 'aGVsbG8=');
    assert.equal(rt.run('Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, ""))'),
        '1B2M2Y8AsgTpgAmY7PhCfg==');
    assert.equal(rt.run('Utilities.newBlob(Utilities.base64Decode("aGVsbG8=")).getDataAsString()'), 'hello');
    assert.equal(hex(rt.run('Utilities.computeHmacSha256Signature("msg", "key")')),
        '2d93cbc1be167bcb1637a4a23cbff01a7878f0c50ee833954ea5221bb1b8c628');
    assert.match(rt.run('Utilities.getUuid()'), /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/);
    assert.equal(rt.run('Utilities.formatString("%s has %d tags (%.1f%%)", "ws", 12, 12.34)'), 'ws has 12 tags (12.3%)');
  });

  it('Utilities.formatDate supports common Java patterns and time zones', () => {
    const rt = bareRuntime();
    rt.run('var d = new Date(Date.UTC(2024, 6, 4, 18, 5, 9, 7));');
    const fmt = (tz, p) => rt.run(`Utilities.formatDate(d, ${JSON.stringify(tz)}, ${JSON.stringify(p)})`);
    assert.equal(fmt('UTC', 'yyyy-MM-dd HH:mm:ss'), '2024-07-04 18:05:09');
    assert.equal(fmt('America/New_York', 'yyyy-MM-dd HH:mm:ss'), '2024-07-04 14:05:09');
    assert.equal(fmt('Europe/Berlin', "yyyy-MM-dd'T'HH:mm:ss.SSSXXX"), '2024-07-04T20:05:09.007+02:00');
    assert.equal(fmt('GMT', 'EEE, d MMM yyyy h:mm a'), 'Thu, 4 Jul 2024 6:05 PM');
    assert.equal(fmt('America/New_York', 'Z'), '-0400');
    assert.equal(fmt('GMT+05:30', 'HH:mm'), '23:35');
    assert.equal(fmt('UTC', "MMMM dd, yy 'at' H'h'"), 'July 04, 24 at 18h');
  });

  it('PropertiesService stores strings per scope and enforces quotas', () => {
    const rt = bareRuntime({properties: {script: {seeded: 'yes'}}});
    rt.run(`PropertiesService.getDocumentProperties().setProperty('a', 1).setProperties({b: 'two', c: true});
      PropertiesService.getUserProperties().setProperty('u', 'x');`);
    assert.deepEqual(rt.properties.document.all(), {a: '1', b: 'two', c: 'true'});
    assert.deepEqual(rt.toHost(rt.run('PropertiesService.getDocumentProperties().getProperties()')),
        {a: '1', b: 'two', c: 'true'});
    assert.equal(rt.run('PropertiesService.getScriptProperties().getProperty("seeded")'), 'yes');
    assert.equal(rt.run('PropertiesService.getUserProperties().getProperty("missing")'), null);
    rt.run('PropertiesService.getDocumentProperties().deleteProperty("a").setProperties({z: "1"}, true)');
    assert.deepEqual(rt.properties.document.all(), {z: '1'});
    assert.equal(catchError(() => rt.run('PropertiesService.getUserProperties().setProperty("big", "x".repeat(10000))')).message,
        'Argument too large: value');
  });

  it('CacheService expires entries on the virtual clock', () => {
    const rt = bareRuntime();
    rt.run(`const c = CacheService.getUserCache(); c.put('k', 'v', 60); c.putAll({a: '1', b: '2'});`);
    assert.equal(rt.run('CacheService.getUserCache().get("k")'), 'v');
    assert.deepEqual(rt.toHost(rt.run('CacheService.getUserCache().getAll(["a", "b", "zz"])')), {a: '1', b: '2'});
    rt.run('Utilities.sleep(61000)');
    assert.equal(rt.run('CacheService.getUserCache().get("k")'), null);
    assert.equal(rt.run('CacheService.getUserCache().get("a")'), '1', 'default TTL is 600 s');
    assert.equal(rt.run('CacheService.getScriptCache().get("a")'), null, 'scopes are separate');
    rt.advanceClock(600000);
    assert.deepEqual(rt.cache.user.all(), {});
  });

  it('LockService locks and simulated contention', () => {
    const rt = bareRuntime();
    assert.deepEqual(rt.toHost(rt.run(`(function () {
      const lock = LockService.getDocumentLock();
      const got = lock.tryLock(1000);
      const has = lock.hasLock();
      lock.releaseLock();
      return [got, has, lock.hasLock()];
    })()`)), [true, true, false]);
    rt.locks.simulateContention(true, 'document');
    assert.equal(rt.run('LockService.getDocumentLock().tryLock(5000)'), false);
    assert.equal(catchError(() => rt.run('LockService.getDocumentLock().waitLock(5000)')).message,
        'Lock timeout: another process was holding the lock for too long.');
    assert.equal(rt.run('LockService.getScriptLock().tryLock(1)'), true);
  });

  it('Session, ScriptApp, Logger and console', () => {
    const rt = bareRuntime({userEmail: 'me@example.com', timeZone: 'Asia/Tokyo',
      oauthToken: () => 'token-from-function-123'});
    assert.equal(rt.run('Session.getActiveUser().getEmail()'), 'me@example.com');
    assert.equal(rt.run('Session.getScriptTimeZone()'), 'Asia/Tokyo');
    assert.equal(rt.run('ScriptApp.getOAuthToken()'), 'token-from-function-123');
    assert.equal(bareRuntime().run('ScriptApp.getOAuthToken()'), 'test-oauth-token');
    rt.run(`Logger.log('Found %s tags', 3); console.log('obj', {a: 1}); console.error(new Error('boom'));
      console.info('%s of %d done', 'two', 5, 'extra');`);
    assert.deepEqual(rt.logs.slice(0, 2).map((l) => [l.source, l.level, l.message]),
        [['Logger', 'log', 'Found 3 tags'], ['console', 'log', 'obj {"a":1}']]);
    assert.match(rt.logs[2].message, /^Error: boom/);
    assert.equal(rt.logs[3].message, 'two of 5 done extra');
    assert.match(rt.logText(), /Found 3 tags/);
    rt.run('ScriptApp.newTrigger("nightly").timeBased().everyDays(1).atHour(3).create()');
    assert.equal(rt.triggers[0].spec.functionName, 'nightly');
  });
});

// ===========================================================================
describe('UrlFetchApp mock', () => {
  it('normalizes requests like Apps Script and exposes HTTPResponse', () => {
    const seen = [];
    const rt = bareRuntime({fetchHandler: (req) => {
      seen.push(req);
      return {code: 201, body: {ok: true}, headers: {'X-Test': ['a', 'b']}};
    }});
    const out = rt.toHost(rt.run(`(function () {
      const r1 = UrlFetchApp.fetch('https://example.test/a');
      const r2 = UrlFetchApp.fetch('https://example.test/b', {method: 'post', payload: 'x=1'});
      const r3 = UrlFetchApp.fetch('https://example.test/c', {method: 'put', contentType: 'application/json',
        payload: JSON.stringify({a: 1}), headers: {'X-Custom': 'yes'}});
      return [r1.getResponseCode(), JSON.parse(r1.getContentText()), r1.getHeaders()['X-Test'],
        r1.getAllHeaders()['X-Test'], r3.getBlob().getDataAsString(), r1.getBlob().getContentType()];
    })()`));
    assert.deepEqual(out, [201, {ok: true}, 'b', ['a', 'b'], '{\n  "ok": true\n}', 'application/json']);
    assert.equal(seen[0].method, 'GET');
    assert.equal(seen[0].payload, null);
    assert.equal(seen[1].method, 'POST');
    assert.equal(seen[1].contentType, 'application/x-www-form-urlencoded', 'Apps Script default content type');
    assert.deepEqual(seen[2].json(), {a: 1});
    assert.equal(seen[2].header('x-custom'), 'yes');
  });

  it('form-encodes object payloads and warns when JSON was intended', () => {
    const seen = [];
    const rt = bareRuntime({fetchHandler: (req) => (seen.push(req), {code: 200, body: ''})});
    rt.run(`UrlFetchApp.fetch('https://example.test/x', {method: 'post', contentType: 'application/json',
      payload: {name: 'a b', n: 1}})`);
    assert.equal(seen[0].payload, 'name=a%20b&n=1');
    assert.equal(rt.warnings[0].code, 'object-payload-with-json-content-type');
  });

  it('throws Apps Script-style errors unless muteHttpExceptions', () => {
    const body = JSON.stringify({error: {code: 404, message: 'nope'.repeat(100)}}, null, 2);
    const rt = bareRuntime({fetchHandler: () => ({code: 404, body})});
    const err = catchError(() => rt.run('UrlFetchApp.fetch("https://api.example.test/v1/things?x=1")'));
    assert.equal(err.name, 'Exception');
    assert.match(err.message, /^Request failed for https:\/\/api\.example\.test returned code 404\. Truncated server response: \{ "error": \{ "code": 404, "message": "nopenope/);
    assert.match(err.message, /\.\.\. \(use muteHttpExceptions option to examine full response\)$/);
    assert.equal(rt.run('UrlFetchApp.fetch("https://api.example.test/", {muteHttpExceptions: true}).getResponseCode()'), 404);
  });

  it('fetchAll keeps order; getRequest describes the request', () => {
    const rt = bareRuntime({fetchHandler: (req) => ({code: 200, body: req.url.slice(-1)})});
    assert.deepEqual(rt.toHost(rt.run(`UrlFetchApp.fetchAll(['https://e.test/1', {url: 'https://e.test/2', method: 'post', payload: 'p'}])
      .map(r => r.getContentText())`)), ['1', '2']);
    const req = rt.toHost(rt.run('UrlFetchApp.getRequest("https://e.test/x", {method: "post", payload: "a=b"})'));
    assert.equal(req.method, 'post');
    assert.equal(req.contentType, 'application/x-www-form-urlencoded');
    assert.equal(req.payload, 'a=b');
  });

  it('never logs credentials', () => {
    // Deliberately fake and not shaped like a real OAuth token (keeps git-secrets quiet).
    const token = 'fake-oauth-token-for-redaction-test';
    const rt = bareRuntime({oauthToken: token, fetchHandler: () => ({code: 200, body: 'ok'})});
    rt.run(`UrlFetchApp.fetch('https://e.test/?echo=' + ScriptApp.getOAuthToken(),
      {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}, method: 'post', payload: 'token=' + ScriptApp.getOAuthToken()})`);
    const logged = JSON.stringify(rt.fetchLog);
    assert.ok(!logged.includes(token), logged);
    assert.equal(rt.fetchLog[0].requestHeaders.Authorization, 'Bearer <redacted>');
    assert.match(rt.fetchLog[0].url, /echo=<redacted>/);
    assert.equal(rt.fetchLog[0].status, 200);
    assert.equal(typeof rt.fetchLog[0].durationMs, 'number');
  });

  it('surfaces harness problems loudly', () => {
    const rt = bareRuntime();
    assert.match(catchError(() => rt.run('UrlFetchApp.fetch("https://e.test/")')).message,
        /HarnessError: UrlFetchApp mock: no fetchHandler response/);
    assert.equal(rt.harnessErrors.length, 1);
    const throwing = bareRuntime({fetchHandler: () => {
      throw new TypeError('bug in fake');
    }});
    assert.match(catchError(() => throwing.run('UrlFetchApp.fetch("https://e.test/")')).message,
        /HarnessError: fetchHandler threw .*bug in fake/s);
    const offline = bareRuntime({fetchHandler: () => ({networkError: 'Address unavailable'})});
    assert.equal(catchError(() => offline.run('UrlFetchApp.fetch("https://e.test/x")')).message,
        'Address unavailable: https://e.test/x');
  });
});

// ===========================================================================
describe('fake GTM API', () => {
  it('lists and gets accounts and containers; looks up containers', () => {
    const {rt, fake} = gtmRuntime();
    fake.seed({
      accounts: [{accountId: '111', name: 'Acme'}],
      containers: [{accountId: '111', containerId: '222', name: 'Web', publicId: 'GTM-WEB1'},
        {accountId: '111', containerId: '333', name: 'Server', publicId: 'GTM-SRV1', usageContext: ['server']}],
      destinations: [{accountId: '111', containerId: '222', destinationId: 'G-ABC'}],
    });
    const accounts = rt.call('api', 'get', 'accounts');
    assert.deepEqual(accounts.body.account.map((a) => [a.accountId, a.name, a.path]), [['111', 'Acme', 'accounts/111']]);
    const containers = rt.call('api', 'get', 'accounts/111/containers').body.container;
    assert.deepEqual(containers.map((c) => [c.containerId, c.publicId, c.usageContext]),
        [['222', 'GTM-WEB1', ['web']], ['333', 'GTM-SRV1', ['server']]]);
    assert.equal(rt.call('api', 'get', 'accounts/containers:lookup?tagId=GTM-SRV1').body.containerId, '333');
    assert.equal(rt.call('api', 'get', 'accounts/containers:lookup?destinationId=G-ABC').body.containerId, '222');
    const missing = rt.call('api', 'get', 'accounts/containers:lookup?tagId=GTM-NOPE');
    assert.deepEqual(missing.body, {error: {code: 404, message: 'Not found or permission denied.',
      errors: [{message: 'Not found or permission denied.', domain: 'global', reason: 'notFound'}], status: 'NOT_FOUND'}});
    assert.equal(rt.call('api', 'get', 'accounts/999').code, 404);
  });

  it('implements tag CRUD with server-managed fields', () => {
    const {rt, fake} = gtmRuntime();
    const {workspacePath} = fake.seedWorkspace({accountId: '10', containerId: '20', workspaceId: '3',
      triggers: [{triggerId: '7', name: 'Click', type: 'click'}]});
    const created = rt.call('api', 'post', `${workspacePath}/tags`, {name: 'GA4 - purchase', type: 'gaawe',
      paused: false, notes: '', firingTriggerId: [7, '2147479553'], parameter: [{type: 'template', key: 'eventName', value: 'purchase'}]});
    assert.equal(created.code, 200);
    const tag = created.body;
    assert.equal(tag.tagId, '8', 'ids come from a per-container counter shared by all kinds');
    assert.equal(tag.path, `${workspacePath}/tags/8`);
    assert.deepEqual([tag.accountId, tag.containerId, tag.workspaceId], ['10', '20', '3']);
    assert.match(tag.fingerprint, /^\d{13,}$/);
    assert.equal(tag.tagManagerUrl, `https://tagmanager.google.com/#/container/${workspacePath}/tags/8?apiLink=tag`);
    assert.deepEqual(tag.firingTriggerId, ['7', '2147479553']);
    assert.ok(!('paused' in tag) && !('notes' in tag), 'false/empty fields are omitted like the real API');

    assert.deepEqual(rt.call('api', 'get', tag.path).body, tag);
    const stale = rt.call('api', 'put', `${tag.path}?fingerprint=123`, {...tag, name: 'Renamed'});
    assert.equal(stale.code, 409);
    assert.equal(stale.body.error.status, 'ABORTED');
    const updated = rt.call('api', 'put', `${tag.path}?fingerprint=${tag.fingerprint}`, {name: 'Renamed', type: 'gaawe'});
    assert.equal(updated.code, 200);
    assert.equal(updated.body.name, 'Renamed');
    assert.notEqual(updated.body.fingerprint, tag.fingerprint);
    assert.equal(updated.body.parameter, undefined, 'PUT replaces the whole entity');

    const deleted = rt.call('api', 'delete', tag.path);
    assert.deepEqual([deleted.code, deleted.text], [204, '']);
    assert.equal(rt.call('api', 'get', tag.path).code, 404);
    assert.equal(rt.call('api', 'delete', tag.path).code, 404);
  });

  it('validates names, types and references', () => {
    const {rt, fake} = gtmRuntime();
    const {workspacePath} = fake.seedWorkspace({tags: [{name: 'Existing', type: 'html'}]});
    const err = (body) => rt.call('api', 'post', `${workspacePath}/tags`, body).body.error;
    assert.deepEqual(Object.keys(err({type: 'html'})), ['code', 'message', 'errors', 'status']);
    assert.equal(err({type: 'html'}).code, 400);
    assert.match(err({type: 'html'}).message, /name field is required/);
    assert.match(err({name: 'Existing', type: 'html'}).message, /duplicate name/);
    assert.match(err({name: 'No type'}).message, /type field is required/);
    assert.match(err({name: 'Bad ref', type: 'html', firingTriggerId: ['999']}).message, /Unknown trigger ID "999"/);
    assert.match(err({name: 'Bad folder', type: 'html', parentFolderId: '55'}).message, /Unknown folder ID "55"/);
    assert.equal(err({name: 'Bad folder', type: 'html', parentFolderId: '55'}).errors[0].reason, 'badRequest');
    assert.equal(rt.call('api', 'post', `${workspacePath}/tags`, {name: 'Ok', type: 'html',
      firingTriggerId: [BUILT_IN_TRIGGER_IDS[0]]}).code, 200);
    // Same name is fine for a different kind.
    assert.equal(rt.call('api', 'post', `${workspacePath}/triggers`, {name: 'Existing', type: 'pageview'}).code, 200);
    // Seeds are checked too, so broken test data fails fast.
    assert.throws(() => createFakeGtmApi().seedWorkspace({tags: [{name: 'x', type: 'html', firingTriggerId: ['42']}]}),
        /FakeGtmApi\.seed: tag "x" .*Unknown trigger ID "42"/);
    const ids = createFakeGtmApi().seedWorkspace({tags: [{name: 't', type: 'html'}], folders: [{name: 'f'}]});
    assert.equal(ids.workspacePath, 'accounts/6000001/containers/9000001/workspaces/1');
  });

  it('rejects unknown fields and invalid enum values like the API JSON parser', () => {
    const {rt, fake} = gtmRuntime();
    const {workspacePath} = fake.seedWorkspace();
    const err = (body) => rt.call('api', 'post', `${workspacePath}/tags`, body).body.error.message;
    assert.equal(err({name: 'T', type: 'html', firingTriggerIds: ['1']}),
        'Invalid JSON payload received. Unknown name "firingTriggerIds" at \'tag\': Cannot find field.');
    assert.equal(err({name: 'T', type: 'html', parameter: [{type: 'template', key: 'k', valeu: 'x'}]}),
        'Invalid JSON payload received. Unknown name "valeu" at \'tag.parameter[0]\': Cannot find field.');
    assert.equal(err({name: 'T', type: 'html', tagFiringOption: 'everyTime'}),
        'Invalid value at \'tag.tag_firing_option\' (TYPE_ENUM), "everyTime"');
    assert.match(err({name: 'T', type: 'html', firingTriggerId: '2147479553'}),
        /^Invalid value at 'tag\.firing_trigger_id' \(repeated field/);
    assert.match(err({name: 'T', type: 'html', consentSettings: {consentStatus: 'maybe'}}),
        /^Invalid value at 'tag\.consent_settings\.consent_status' \(TYPE_ENUM\)/);
    assert.match(rt.call('api', 'post', `${workspacePath}/triggers`, {name: 'Tr', type: 'pageView'}).body.error.message,
        /^Invalid value at 'trigger\.type' \(TYPE_ENUM\), "pageView"/);
    const bulk = rt.call('api', 'post', `${workspacePath}/bulk_update`,
        {changes: [{template: {templateId: 'new_1', name: 'x'}, changeStatus: 'added'}]});
    assert.equal(bulk.body.error.message,
        'Invalid JSON payload received. Unknown name "template" at \'changes[0]\': Cannot find field.');
    assert.throws(() => fake.seedWorkspace({tags: [{name: 'Typo', type: 'html', pased: true}]}),
        /FakeGtmApi\.seed: invalid tag .*Unknown name "pased"/);
    const lenient = gtmRuntime({validateSchema: false});
    const ws = lenient.fake.seedWorkspace().workspacePath;
    assert.equal(lenient.rt.call('api', 'post', `${ws}/tags`, {name: 'T', type: 'html', extra: 1}).code, 200);
  });

  it('paginates list responses and omits empty lists', () => {
    const tags = Array.from({length: 450}, (_, i) => ({name: `Tag ${i + 1}`, type: 'html'}));
    const {rt, fake} = gtmRuntime({pageSize: 200});
    const {workspacePath} = fake.seedWorkspace({tags});
    const all = rt.call('listAll', `${workspacePath}/tags`, 'tag');
    assert.equal(all.pages, 3);
    assert.equal(all.items.length, 450);
    assert.deepEqual(all.items.slice(0, 2).map((t) => t.tagId), ['1', '2']);
    const first = rt.call('api', 'get', `${workspacePath}/tags`).body;
    assert.equal(first.tag.length, 200);
    assert.equal(typeof first.nextPageToken, 'string');
    assert.deepEqual(rt.call('api', 'get', `${workspacePath}/variables`).body, {}, 'empty list -> {}');
    assert.equal(rt.call('api', 'get', `${workspacePath}/tags?pageToken=garbage`).code, 400);
  });

  it('supports every workspace entity kind', () => {
    const {rt, fake} = gtmRuntime();
    const {workspacePath} = fake.seedWorkspace({usageContext: ['server']});
    const kinds = [
      ['triggers', 'trigger', 'triggerId', {name: 'tr', type: 'always'}],
      ['variables', 'variable', 'variableId', {name: 'v', type: 'c', parameter: [{type: 'template', key: 'value', value: '1'}]}],
      ['folders', 'folder', 'folderId', {name: 'f'}],
      ['clients', 'client', 'clientId', {name: 'GA4', type: 'gaaw_client', priority: 0}],
      ['transformations', 'transformation', 'transformationId', {name: 'x', type: 'tf_allow_params'}],
      ['templates', 'template', 'templateId', {name: 'tpl', templateData: '___INFO___'}],
      ['zones', 'zone', 'zoneId', {name: 'z'}],
      ['gtag_config', 'gtagConfig', 'gtagConfigId', {type: 'google', parameter: []}],
    ];
    for (const [collection, listKey, idField, body] of kinds) {
      const created = rt.call('api', 'post', `${workspacePath}/${collection}`, body);
      assert.equal(created.code, 200, `${collection}: ${created.text}`);
      assert.match(created.body[idField], /^\d+$/);
      assert.equal(created.body.path, `${workspacePath}/${collection}/${created.body[idField]}`);
      const listed = rt.call('api', 'get', `${workspacePath}/${collection}`).body[listKey];
      assert.equal(listed.length, 1, collection);
    }
    assert.equal(fake.list(workspacePath, 'gtag_config').length, 1);
    assert.equal(fake.find(workspacePath, 'clients', 'GA4').priority, 0);
  });

  it('enables, lists and disables built-in variables', () => {
    const {rt, fake} = gtmRuntime();
    const {workspacePath} = fake.seedWorkspace({builtInVariables: ['pageUrl']});
    const created = rt.call('api', 'post', `${workspacePath}/built_in_variables?type=clickText&type=event`);
    assert.deepEqual(created.body.builtInVariable.map((b) => [b.type, b.name]), [['clickText', 'Click Text'], ['event', 'Event']]);
    const listed = rt.call('api', 'get', `${workspacePath}/built_in_variables`).body.builtInVariable;
    assert.deepEqual(listed.map((b) => b.type), ['pageUrl', 'clickText', 'event']);
    assert.equal(rt.call('api', 'delete', `${workspacePath}/built_in_variables?type=pageUrl`).code, 204);
    assert.equal(rt.call('api', 'post', `${workspacePath}/built_in_variables?type=notAType`).code, 400);
    assert.deepEqual(fake.list(workspacePath, 'builtInVariable').map((b) => b.type), ['clickText', 'event']);
  });

  it('manages workspaces and reports workspace status', () => {
    const {rt, fake} = gtmRuntime();
    const {containerPath, workspacePath} = fake.seedWorkspace({
      tags: [{tagId: '1', name: 'Keep', type: 'html'}, {tagId: '2', name: 'Change', type: 'html'},
        {tagId: '3', name: 'Remove', type: 'html'}]});
    assert.deepEqual(rt.call('api', 'get', `${workspacePath}/status`).body, {}, 'no changes -> {}');
    rt.call('api', 'put', `${workspacePath}/tags/2`, {name: 'Changed', type: 'html'});
    rt.call('api', 'delete', `${workspacePath}/tags/3`);
    rt.call('api', 'post', `${workspacePath}/tags`, {name: 'New', type: 'html'});
    const status = rt.call('api', 'get', `${workspacePath}/status`).body.workspaceChange;
    assert.deepEqual(status.map((c) => [c.tag.tagId, c.tag.name, c.changeStatus]),
        [['2', 'Changed', 'updated'], ['3', 'Remove', 'deleted'], ['4', 'New', 'added']]);
    const reverted = rt.call('api', 'post', `${workspacePath}/tags/2:revert`);
    assert.equal(reverted.body.tag.name, 'Change');

    const ws2 = rt.call('api', 'post', `${containerPath}/workspaces`, {name: 'Feature'});
    assert.equal(ws2.body.workspaceId, '2');
    assert.deepEqual(rt.call('listAll', `${ws2.body.path}/tags`, 'tag').items.map((t) => t.name),
        ['Keep', 'Change', 'Remove'], 'new workspaces start from the container version');
    assert.equal(rt.call('api', 'post', `${containerPath}/workspaces`, {name: 'Third'}).code, 200);
    const limit = rt.call('api', 'post', `${containerPath}/workspaces`, {name: 'Fourth'});
    assert.equal(limit.code, 400);
    assert.match(limit.body.error.message, /maximum number of workspaces/);
    assert.equal(rt.call('api', 'delete', ws2.body.path).code, 204);
    assert.equal(rt.call('api', 'get', `${containerPath}/workspaces`).body.workspace.length, 2);
  });

  it('bulk_update remaps new_N ids, including forward references', () => {
    const {rt, fake} = gtmRuntime();
    const {workspacePath} = fake.seedWorkspace({tags: [{tagId: '5', name: 'Old', type: 'html'},
      {tagId: '6', name: 'Gone', type: 'html'}]});
    const res = rt.call('api', 'post', `${workspacePath}/bulk_update`, {changes: [
      {tag: {tagId: 'new_3', name: 'New tag', type: 'html', parentFolderId: 'new_1', firingTriggerId: ['new_2']},
        changeStatus: 'added'},
      {folder: {folderId: 'new_1', name: 'New folder'}, changeStatus: 'added'},
      {trigger: {triggerId: 'new_2', name: 'New trigger', type: 'pageview', parentFolderId: 'new_1'},
        changeStatus: 'added'},
      {tag: {tagId: '5', name: 'Old (renamed)', type: 'html'}, changeStatus: 'updated'},
      {tag: {tagId: '6'}, changeStatus: 'deleted'},
    ]});
    assert.equal(res.code, 200, res.text);
    const changes = res.body.changes;
    assert.deepEqual(changes.map((c) => [Object.keys(c).find((k) => k !== 'changeStatus'), c.changeStatus]),
        [['tag', 'added'], ['folder', 'added'], ['trigger', 'added'], ['tag', 'updated']],
        'request order; deleted entities are not returned');
    const [newTag, folder, trigger] = [changes[0].tag, changes[1].folder, changes[2].trigger];
    assert.deepEqual([newTag.tagId, folder.folderId, trigger.triggerId], ['7', '8', '9'],
        'real ids are allocated in request order');
    assert.equal(newTag.parentFolderId, folder.folderId);
    assert.deepEqual(newTag.firingTriggerId, [trigger.triggerId]);
    assert.equal(trigger.parentFolderId, folder.folderId);
    assert.equal(fake.get(`${workspacePath}/tags/6`), null);
    assert.equal(fake.find(workspacePath, 'tag', 'Old (renamed)').tagId, '5');
  });

  it('bulk_update is atomic and names the failing change', () => {
    const {rt, fake} = gtmRuntime();
    const {workspacePath} = fake.seedWorkspace({tags: [{tagId: '1', name: 'Taken', type: 'html'}]});
    const before = JSON.stringify(fake.state);
    const res = rt.call('api', 'post', `${workspacePath}/bulk_update`, {changes: [
      {folder: {folderId: 'new_1', name: 'F'}, changeStatus: 'added'},
      {tag: {tagId: '1', name: 'Taken', type: 'html', parentFolderId: 'new_1'}, changeStatus: 'updated'},
      {tag: {tagId: 'new_2', name: 'Taken', type: 'html'}, changeStatus: 'added'},
    ]});
    assert.equal(res.code, 400);
    assert.match(res.body.error.message, /^changes\[2\] \(added tag new_2\): Found entity with duplicate name/);
    assert.equal(JSON.stringify(fake.state), before, 'nothing was applied');
    const next = rt.call('api', 'post', `${workspacePath}/folders`, {name: 'After'});
    assert.equal(next.body.folderId, '2', 'id counter rolled back too');

    const badId = rt.call('api', 'post', `${workspacePath}/bulk_update`,
        {changes: [{tag: {tagId: '17', name: 'X', type: 'html'}, changeStatus: 'added'}]});
    assert.match(badId.body.error.message, /^changes\[0\]: .*"new_" followed by a number/);
    const wrongKind = rt.call('api', 'post', `${workspacePath}/bulk_update`, {changes: [
      {tag: {tagId: 'new_1', name: 'A', type: 'html'}, changeStatus: 'added'},
      {tag: {tagId: 'new_2', name: 'B', type: 'html', firingTriggerId: ['new_1']}, changeStatus: 'added'},
    ]});
    assert.match(wrongKind.body.error.message, /^changes\[1\].*which is a tag, not a trigger/);
    const missing = rt.call('api', 'post', `${workspacePath}/bulk_update`,
        {changes: [{tag: {tagId: '404', name: 'X', type: 'html'}, changeStatus: 'updated'}]});
    assert.equal(missing.code, 404);
  });

  it('enforces auth and JSON content type', () => {
    const {rt, fake} = gtmRuntime({acceptedTokens: ['good-token-12345']}, {oauthToken: 'good-token-12345'});
    const {workspacePath} = fake.seedWorkspace();
    const noAuth = rt.toHost(rt.run(`(function () {
      const r = UrlFetchApp.fetch(API + '${workspacePath}/tags', {muteHttpExceptions: true});
      return [r.getResponseCode(), JSON.parse(r.getContentText()).error.status];
    })()`));
    assert.deepEqual(noAuth, [401, 'UNAUTHENTICATED']);
    const badToken = rt.toHost(rt.run(`UrlFetchApp.fetch(API + '${workspacePath}/tags',
      {muteHttpExceptions: true, headers: {Authorization: 'Bearer wrong-token-999'}}).getResponseCode()`));
    assert.equal(badToken, 401);
    const formEncoded = rt.call('api', 'post', `${workspacePath}/tags`, undefined,
        {payload: JSON.stringify({name: 'x', type: 'html'})});
    assert.equal(formEncoded.code, 400);
    assert.match(formEncoded.body.error.message, /^Invalid JSON payload received\. Unknown name "\{"name":"x"/);
    assert.equal(rt.call('api', 'get', `${workspacePath}/tags`).code, 200);
  });

  it('simulates rate limits on the virtual clock', () => {
    const {rt, fake} = gtmRuntime({rateLimit: {max: 2, windowMs: 60000}});
    const {workspacePath} = fake.seedWorkspace();
    const codes = [1, 2, 3].map(() => rt.call('api', 'get', `${workspacePath}/tags`).code);
    assert.deepEqual(codes, [200, 200, 429]);
    const last = fake.requests.at(-1).response.error;
    assert.equal(last.status, 'RESOURCE_EXHAUSTED');
    assert.equal(last.errors[0].reason, 'rateLimitExceeded');
    rt.run('Utilities.sleep(60000)');
    assert.equal(rt.call('api', 'get', `${workspacePath}/tags`).code, 200, 'window passed (virtual time)');
  });

  it('injects errors for resilience tests', () => {
    const {rt, fake} = gtmRuntime();
    const {workspacePath} = fake.seedWorkspace();
    fake.injectError({method: 'GET', path: /\/tags$/, code: 503, message: 'The service is currently unavailable.'});
    assert.equal(rt.call('api', 'get', `${workspacePath}/tags`).code, 503);
    assert.equal(rt.call('api', 'get', `${workspacePath}/tags`).code, 200, 'consumed after `times` uses');
    fake.injectError({path: `${workspacePath}/folders`, networkError: 'Timeout'});
    assert.equal(catchError(() => rt.call('api', 'get', `${workspacePath}/folders`)).message,
        `Timeout: ${API}${workspacePath}/folders`);
  });

  it('links destinations', () => {
    const {rt, fake} = gtmRuntime();
    fake.seed({containers: [{accountId: '1', containerId: '2'}, {accountId: '1', containerId: '3'}]});
    const linked = rt.call('api', 'post', 'accounts/1/containers/2/destinations:link?destinationId=G-XYZ');
    assert.equal(linked.code, 200);
    assert.equal(linked.body.destinationId, 'G-XYZ');
    assert.equal(rt.call('api', 'get', linked.body.path).body.destinationLinkId, linked.body.destinationLinkId);
    const moved = rt.call('api', 'post', 'accounts/1/containers/3/destinations:link?destinationId=G-XYZ');
    assert.equal(moved.code, 200);
    assert.deepEqual(rt.call('api', 'get', 'accounts/1/containers/2/destinations').body, {});
    assert.equal(rt.call('api', 'get', 'accounts/1/containers/3/destinations').body.destination.length, 1);
    assert.equal(rt.call('api', 'get', 'accounts/containers:lookup?destinationId=G-XYZ').body.containerId, '3');
  });

  it('offers state, snapshots and request logs for assertions', () => {
    const {rt, fake} = gtmRuntime();
    const {workspacePath} = fake.seedWorkspace({folders: [{name: 'F'}]});
    const snap = fake.snapshot();
    rt.call('api', 'post', `${workspacePath}/tags`, {name: 'T', type: 'html', parentFolderId: '1'});
    assert.equal(fake.state.entities[workspacePath].tag.length, 1);
    assert.deepEqual(fake.workspaceChanges(workspacePath).map((c) => [c.kind, c.changeStatus]), [['tag', 'added']]);
    const posts = fake.requestsTo('POST', /\/tags$/);
    assert.equal(posts.length, 1);
    assert.deepEqual(posts[0].body, {name: 'T', type: 'html', parentFolderId: '1'});
    assert.equal(posts[0].status, 200);
    fake.restore(snap);
    assert.equal(fake.state.entities[workspacePath].tag.length, 0);
    assert.equal(fake.handle({url: 'https://example.test/', method: 'GET'}), undefined, 'non-GTM URLs pass through');
    assert.equal(rt.call('api', 'get', `${workspacePath}:quick_preview`).code, 501, 'unimplemented routes say so');
  });
});

// ===========================================================================
describe('realHttp mode (curl)', () => {
  const online = isOnline();

  it('GETs the public GTM discovery document', {skip: online ? false : 'offline'}, () => {
    const rt = bareRuntime({realHttp: true});
    const out = rt.toHost(rt.run(`(function () {
      const r = UrlFetchApp.fetch('https://tagmanager.googleapis.com/$discovery/rest?version=v2');
      const doc = JSON.parse(r.getContentText());
      return [r.getResponseCode(), doc.name, doc.version, Object.keys(doc.resources)];
    })()`));
    assert.deepEqual(out, [200, 'tagmanager', 'v2', ['accounts']]);
    assert.equal(rt.fetchLog[0].mode, 'realHttp');
    assert.equal(rt.fetchLog[0].status, 200);
  });

  it('reports HTTP errors like Apps Script without leaking the token', {skip: online ? false : 'offline'}, () => {
    const token = 'harness-bogus-token-0123456789';
    const rt = bareRuntime({realHttp: true, oauthToken: token});
    const muted = rt.toHost(rt.run(`(function () {
      const r = UrlFetchApp.fetch('${API}accounts', {muteHttpExceptions: true,
        headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}});
      return [r.getResponseCode(), JSON.parse(r.getContentText()).error.status];
    })()`));
    assert.deepEqual(muted, [401, 'UNAUTHENTICATED']);
    const err = catchError(() => rt.run(`UrlFetchApp.fetch('${API}accounts',
      {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}})`));
    assert.match(err.message, /^Request failed for https:\/\/tagmanager\.googleapis\.com returned code 401\./);
    assert.ok(!JSON.stringify(rt.fetchLog).includes(token));
    assert.ok(!err.message.includes(token));
  });

  const liveToken = (() => {
    try {
      return resolveLiveToken();
    } catch (e) {
      return null;
    }
  })();
  it('lists real GTM accounts (live; needs GTM_OAUTH_TOKEN or GTM_OAUTH_TOKEN_CMD)',
      {skip: online && liveToken ? false : 'no live token configured'}, () => {
        const rt = bareRuntime({realHttp: true, sources: [{filename: 'gtmHelpers.js', code: GTM_HELPERS}]});
        const accounts = rt.call('listAll', 'accounts', 'account');
        assert.ok(Array.isArray(accounts.items));
        assert.ok(!JSON.stringify(rt.fetchLog).includes(liveToken));
      });
});
