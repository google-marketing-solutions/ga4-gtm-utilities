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
 * @fileoverview Read-only probe of real GTM containers, used during
 * development to inspect entity types and parameter keys.
 *
 * Usage: node tests/live/probe.js OUTPUT_DIR GTM-XXXXXXX [GTM-YYYYYYY ...]
 * Writes one JSON file per container with all entities of its first
 * (default) workspace. Nothing is modified.
 */

const {execFileSync} = require('node:child_process');
const fs = require('node:fs');
const path = require('node:path');
const {getLiveAccessToken, getLiveExtraHeaders} = require('./liveAuth');

const BASE = 'https://tagmanager.googleapis.com/tagmanager/v2/';

/**
 * Makes a GET request.
 * @param {string} apiPath
 * @return {{code: number, body: *}}
 */
function get(apiPath) {
  const headers = Object.assign({Authorization: 'Bearer ' +
      getLiveAccessToken()}, getLiveExtraHeaders());
  const args = ['-sS', '-w', '\n%{http_code}', BASE + apiPath];
  Object.keys(headers).forEach(key => args.push('-H', `${key}: ${headers[key]}`));
  const output = execFileSync('curl', args, {encoding: 'utf8'});
  const index = output.lastIndexOf('\n');
  const code = Number(output.slice(index + 1));
  let body = output.slice(0, index);
  try {
    body = JSON.parse(body);
  } catch (e) {
    // Keep the text.
  }
  return {code, body};
}

/**
 * Lists all items of a collection.
 * @param {string} apiPath
 * @param {string} key
 * @return {!Array<!Object>|{error: *}}
 */
function listAll(apiPath, key) {
  let items = [];
  let token = '';
  do {
    const response = get(apiPath + (token ? '?pageToken=' +
        encodeURIComponent(token) : ''));
    if (response.code !== 200) return {error: response.body, code: response.code};
    items = items.concat(response.body[key] || []);
    token = response.body.nextPageToken;
  } while (token);
  return items;
}

function main() {
  const [outDir, ...ids] = process.argv.slice(2);
  if (!outDir || !ids.length) {
    console.error('Usage: node probe.js OUTPUT_DIR GTM-XXXX [...]');
    process.exit(1);
  }
  fs.mkdirSync(outDir, {recursive: true});
  for (const id of ids) {
    const lookup = get('accounts/containers:lookup?tagId=' + id);
    if (lookup.code !== 200) {
      console.log(id, 'lookup failed', lookup.code, JSON.stringify(lookup.body));
      continue;
    }
    const container = lookup.body;
    const result = {container};
    result.workspaces = listAll(`${container.path}/workspaces`, 'workspace');
    result.destinations = listAll(`${container.path}/destinations`,
        'destination');
    const workspace = Array.isArray(result.workspaces) && result.workspaces[0];
    if (workspace) {
      for (const [collection, key] of [['tags', 'tag'], ['triggers', 'trigger'],
           ['variables', 'variable'], ['folders', 'folder'],
           ['clients', 'client'], ['transformations', 'transformation'],
           ['templates', 'template'], ['gtag_config', 'gtagConfig'],
           ['built_in_variables', 'builtInVariable'], ['zones', 'zone']]) {
        result[collection] = listAll(`${workspace.path}/${collection}`, key);
      }
      result.status = get(`${workspace.path}/status`).body;
    }
    const file = path.join(outDir, `${id}.json`);
    fs.writeFileSync(file, JSON.stringify(result, null, 2));
    const count = value => Array.isArray(value) ? value.length :
        `error ${value && value.code}`;
    console.log(`${id}: ${container.name} (${container.usageContext}) ` +
        `features=${JSON.stringify(container.features)} tagIds=` +
        `${container.tagIds} workspaces=${count(result.workspaces)} ` +
        `tags=${count(result.tags)} triggers=${count(result.triggers)} ` +
        `variables=${count(result.variables)} clients=${count(result.clients)} ` +
        `gtag_config=${count(result.gtag_config)} destinations=` +
        `${count(result.destinations)} -> ${file}`);
  }
}

main();
