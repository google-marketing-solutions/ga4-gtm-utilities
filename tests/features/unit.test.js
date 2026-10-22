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
 * @fileoverview Unit tests of the helpers (text formats, parameters, trigger
 * conditions), run inside the Apps Script runtime emulator.
 */
'use strict';

const {test, describe} = require('node:test');
const assert = require('node:assert/strict');
const {createRuntime} = require('../harness/appsScriptRuntime');

const rt = createRuntime({fetchHandler: () => ({code: 500, body: '{}'})});

describe('text formats', () => {
  test('parseList splits by comma or by line', () => {
    assert.deepEqual(rt.call('parseList', 'a, b ,c'), ['a', 'b', 'c']);
    assert.deepEqual(rt.call('parseList', 'a, b\nc'), ['a, b', 'c']);
    assert.deepEqual(rt.call('parseList', ''), []);
    assert.deepEqual(rt.call('parseList', '  \n '), []);
  });

  test('formatList uses lines when an item has a comma', () => {
    assert.equal(rt.call('formatList', ['a', 'b']), 'a, b');
    assert.equal(rt.call('formatList', ['a, b', 'c']), 'a, b\nc');
  });

  test('parseTable and formatTable round-trip', () => {
    const text = 'page_type = {{Page Type}}\nurl = https://x.com/?a=b\nempty =';
    const rows = rt.call('parseTable', text);
    assert.deepEqual(rows, [
      {name: 'page_type', value: '{{Page Type}}'},
      {name: 'url', value: 'https://x.com/?a=b'},
      {name: 'empty', value: ''}
    ]);
    assert.equal(rt.call('formatTable', rows),
        'page_type = {{Page Type}}\nurl = https://x.com/?a=b\nempty = ');
    assert.deepEqual(rt.call('parseTable', 'a=b=c'), [{name: 'a', value: 'b=c'}]);
    assert.deepEqual(rt.call('parseTable', '^/p\\?x=1$ = yes'),
        [{name: '^/p\\?x=1$', value: 'yes'}]);
  });

  test('toText normalises line breaks and trailing spaces', () => {
    assert.equal(rt.call('toText', ' a \r\nb  \n'), 'a\nb');
    assert.equal(rt.call('toText', 12), '12');
    assert.equal(rt.call('toText', null), '');
  });

  test('toBool', () => {
    assert.equal(rt.call('toBool', 'TRUE'), true);
    assert.equal(rt.call('toBool', true), true);
    assert.equal(rt.call('toBool', 'false'), false);
    assert.equal(rt.call('toBool', ''), false);
  });

  test('hashString is stable', () => {
    assert.equal(rt.call('hashString', 'abc'), rt.call('hashString', 'abc'));
    assert.notEqual(rt.call('hashString', 'abc'), rt.call('hashString', 'abd'));
  });

  test('cellSafeValue truncates very long values', () => {
    const long = 'x'.repeat(60000);
    const safe = rt.call('cellSafeValue', long);
    assert.ok(safe.length < 50000);
    assert.match(safe, /truncated/);
    assert.equal(rt.call('cellSafeValue', 'short'), 'short');
  });
});

describe('parameters', () => {
  test('setParamValue keeps compatible types and removes empty values', () => {
    const entity = JSON.parse(rt.run(`(() => {
      const e = {parameter: [{type: 'integer', key: 'n', value: '1'},
                             {type: 'boolean', key: 'b', value: 'false'},
                             {type: 'template', key: 'gone', value: 'x'}]};
      setParamValue(e, 'n', '5');
      setParamValue(e, 'b', 'TRUE');
      setParamValue(e, 'gone', '');
      setParamValue(e, 'new', ' hello ');
      setParamValue(e, 'rdp', '{{RDP}}', {type: 'boolean'});
      return JSON.stringify(e);
    })()`));
    assert.deepEqual(entity.parameter, [
      {type: 'integer', key: 'n', value: '5'},
      {type: 'boolean', key: 'b', value: 'true'},
      {type: 'template', key: 'new', value: 'hello'},
      {type: 'template', key: 'rdp', value: '{{RDP}}'}
    ]);
  });

  test('table rows keep extra fields of existing rows', () => {
    const out = rt.run(`(() => {
      const e = {parameter: [{type: 'list', key: 't', list: [
        {type: 'map', map: [{type: 'template', key: 'k', value: 'a'},
                            {type: 'template', key: 'v', value: '1'},
                            {type: 'boolean', key: 'extra', value: 'true'}]}]}]};
      setTableRows(e, 't', ['k', 'v'], [{name: 'a', value: '2'},
                                         {name: 'b', value: '3'}]);
      return JSON.stringify(e);
    })()`);
    const entity = JSON.parse(out);
    const list = entity.parameter[0].list;
    assert.equal(list.length, 2);
    assert.deepEqual(list[0].map.map(e => [e.key, e.value]),
        [['k', 'a'], ['v', '2'], ['extra', 'true']]);
    assert.deepEqual(list[1].map.map(e => [e.key, e.value]),
        [['k', 'b'], ['v', '3']]);
  });

  test('textToParam converts JSON and keeps the original types', () => {
    const out = rt.run(`JSON.stringify(textToParam('rows',
        '[{"name":"x","on":"false"}]',
        {type: 'list', key: 'rows', list: [{type: 'map', map: [
          {type: 'template', key: 'name', value: 'a'},
          {type: 'boolean', key: 'on', value: 'true'}]}]}))`);
    assert.deepEqual(JSON.parse(out), {type: 'list', key: 'rows', list: [
      {type: 'map', map: [{type: 'template', value: 'x', key: 'name'},
                          {type: 'boolean', value: 'false', key: 'on'}]}]});
    const scalar = JSON.parse(rt.run(
        `JSON.stringify(textToParam('n', '42', {type: 'integer', key: 'n', value: '1'}))`));
    assert.deepEqual(scalar, {type: 'integer', value: '42', key: 'n'});
    assert.throws(() => rt.run(`textToParam('x', '[oops', undefined)`),
        /valid JSON/);
  });

  test('paramToText shows lists and maps as JSON', () => {
    const text = rt.run(`paramToText({type: 'list', key: 'l', list: [
        {type: 'template', value: 'a'}, {type: 'template', value: 'b'}]})`);
    assert.equal(text, '["a","b"]');
  });
});

describe('trigger conditions', () => {
  const roundTrip = line => rt.run(
      `formatCondition_(parseCondition_(${JSON.stringify(line)}))`);

  test('parse and format all operators', () => {
    [
      '{{Page Path}} equals /',
      '{{Page Path}} does not equal /',
      '{{Page URL}} contains checkout',
      '{{Page URL}} does not contain test',
      '{{Page Path}} starts with /shop',
      '{{Page Path}} does not start with /admin',
      '{{Page Path}} ends with .pdf',
      '{{Page Path}} does not end with .html',
      '{{Page URL}} matches RegEx ^https://',
      '{{Page URL}} does not match RegEx (ignore case) staging',
      '{{Click Element}} matches CSS selector .btn',
      '{{Click Element}} does not match CSS selector #nav a',
      '{{value}} less than 10',
      '{{value}} less than or equal to 10',
      '{{value}} greater than 0',
      '{{value}} greater than or equal to 1'
    ].forEach(line => assert.equal(roundTrip(line), line));
  });

  test('operators are case-insensitive and values may be empty', () => {
    assert.equal(roundTrip('{{Click Text}} EQUALS'), '{{Click Text}} equals');
    assert.equal(roundTrip('{{Event}} Matches RegEx (Ignore Case) ^a'),
        '{{Event}} matches RegEx (ignore case) ^a');
  });

  test('parse builds the API structure', () => {
    const condition = JSON.parse(rt.run(`JSON.stringify(parseCondition_(
        '{{Page Path}} does not contain /checkout'))`));
    assert.deepEqual(condition, {type: 'contains', parameter: [
      {type: 'template', key: 'arg0', value: '{{Page Path}}'},
      {type: 'template', key: 'arg1', value: '/checkout'},
      {type: 'boolean', key: 'negate', value: 'true'}]});
  });

  test('invalid conditions give clear errors', () => {
    assert.throws(() => rt.run(`parseCondition_('Page Path contains x')`),
        /must start with a variable/);
    assert.throws(() => rt.run(`parseCondition_('{{Page Path}} is like x')`),
        /Unknown operator/);
  });
});

describe('type names', () => {
  test('resolveTypeLabel accepts names and codes', () => {
    assert.equal(rt.run(`resolveTypeLabel(TAG_TYPE_NAMES, 'ga4 event')`),
        'gaawe');
    assert.equal(rt.run(`resolveTypeLabel(TAG_TYPE_NAMES, 'gaawe')`), 'gaawe');
    assert.equal(rt.run(
        `resolveTypeLabel(VARIABLE_TYPE_NAMES, 'Data Layer Variable')`), 'v');
  });
});
