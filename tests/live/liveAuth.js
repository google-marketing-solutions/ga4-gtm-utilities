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
 * @fileoverview OAuth access token for live tests against the real GTM API.
 *
 * Sources, in order:
 * 1. GTM_OAUTH_TOKEN: an access token with Tag Manager scopes.
 * 2. A named clasp login (default name "gtm-test", override with
 *    GTM_CLASP_USER), created with:
 *      clasp login --user gtm-test --extra-scopes \
 *        https://www.googleapis.com/auth/tagmanager.readonly,\
 *        https://www.googleapis.com/auth/tagmanager.edit.containers,\
 *        https://www.googleapis.com/auth/tagmanager.delete.containers
 *
 * GTM_QUOTA_PROJECT (optional) sets the x-goog-user-project header so that
 * the API calls are attributed to a Google Cloud project with the Tag Manager
 * API enabled.
 */

const {execFileSync} = require('node:child_process');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

let cached = null;

/**
 * Returns an access token (synchronously), or null if none is available.
 * @return {?string}
 */
function getLiveAccessToken() {
  if (process.env.GTM_OAUTH_TOKEN) return process.env.GTM_OAUTH_TOKEN;
  if (cached && cached.expiresAt > Date.now() + 60000) return cached.token;
  const user = process.env.GTM_CLASP_USER || 'gtm-test';
  const rcPath = path.join(os.homedir(), '.clasprc.json');
  if (!fs.existsSync(rcPath)) return null;
  const rc = JSON.parse(fs.readFileSync(rcPath, 'utf8'));
  const creds = rc.tokens && rc.tokens[user];
  if (!creds || !creds.refresh_token) return null;
  const body = new URLSearchParams({
    client_id: creds.client_id,
    client_secret: creds.client_secret,
    refresh_token: creds.refresh_token,
    grant_type: 'refresh_token'
  }).toString();
  const output = execFileSync('curl', ['-sS', '-X', 'POST',
    'https://oauth2.googleapis.com/token',
    '-H', 'Content-Type: application/x-www-form-urlencoded',
    '--data-binary', '@-'], {input: body, encoding: 'utf8'});
  const response = JSON.parse(output);
  if (!response.access_token) {
    throw new Error('Could not refresh the access token: ' +
        (response.error_description || response.error || 'unknown error'));
  }
  cached = {token: response.access_token,
            expiresAt: Date.now() + (response.expires_in || 3600) * 1000};
  return cached.token;
}

/**
 * Returns extra headers for live API requests.
 * @return {!Object<string, string>}
 */
function getLiveExtraHeaders() {
  return process.env.GTM_QUOTA_PROJECT ?
      {'x-goog-user-project': process.env.GTM_QUOTA_PROJECT} : {};
}

module.exports = {getLiveAccessToken, getLiveExtraHeaders};
