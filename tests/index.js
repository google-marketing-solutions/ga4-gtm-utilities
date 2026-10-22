#!/usr/bin/env node
/**
 * @fileoverview Entry point that runs every `*.test.js` file below tests/.
 *
 * Why this file exists:
 *  - `node --test tests/` treats the directory argument as a module path
 *    (Node 22 resolves it to tests/index.js), so this file makes that command
 *    run the whole suite.
 *  - `node tests` (no --test flag) works on Node builds whose `--test` glob
 *    support is broken (e.g. Debian's libnode packaging, which fails with
 *    "Missing internal module 'internal/deps/brace-expansion'").
 *  - `node --test` without arguments (from the repo root or tests/) uses Node's
 *    default patterns and finds the *.test.js files directly; this file does
 *    not match those patterns, so nothing runs twice.
 *
 * Test files are loaded in-process; node:test collects their tests into one run.
 */
'use strict';

const fs = require('node:fs');
const path = require('node:path');

/**
 * Recursively finds *.test.js files, skipping node_modules and fixtures.
 * @param {string} dir
 * @return {!Array<string>}
 */
function findTestFiles(dir) {
  const out = [];
  for (const entry of fs.readdirSync(dir, {withFileTypes: true})) {
    if (entry.name === 'node_modules' || entry.name === 'fixtures' || entry.name.startsWith('.')) continue;
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) out.push(...findTestFiles(full));
    else if (entry.isFile() && /\.test\.(c?js)$/.test(entry.name)) out.push(full);
  }
  return out.sort();
}

const filter = process.env.HARNESS_TEST_FILTER ? new RegExp(process.env.HARNESS_TEST_FILTER) : null;
for (const file of findTestFiles(__dirname)) {
  if (!filter || filter.test(file)) require(file);
}
