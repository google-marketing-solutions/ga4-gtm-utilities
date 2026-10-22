# Tests: zero-dependency Apps Script harness

This directory holds a Node.js (v22) test harness for the Apps Script sources
in `features/`, `shared/` and `utils/`. It needs no npm packages: only Node's
built-ins (`node:test`, `node:assert`, `node:vm`, `node:crypto`,
`node:child_process`, `node:fs`, `node:path`) and the `curl` binary for the
optional real-HTTP mode.

```
tests/
  index.js                    # runs every *.test.js (target of `node --test tests/` and `node tests`)
  package.json                # {"private": true, "scripts": {"test": "node --test", ...}}
  harness/
    appsScriptRuntime.js      # createRuntime(): vm context + Apps Script globals + source loader
    fakeGtmApi.js             # createFakeGtmApi(): in-memory GTM REST API v2 (fetchHandler)
    spreadsheetMock.js        # SpreadsheetApp / Spreadsheet / Sheet / Range
    spreadsheetBuilders.js    # enums, DataValidation, RichText, TextStyle, ConditionalFormat, ...
    spreadsheetHarness.js     # runtime.spreadsheet fixture helpers
    uiMock.js                 # Ui (alerts, prompts, menus, dialogs) + recorder
    servicesMock.js           # Utilities, Session, Properties, Cache, Lock, Logger, console, ScriptApp, HtmlService, ...
    urlFetchMock.js           # UrlFetchApp (fetchHandler or curl transport)
    tagManagerService.js      # `TagManager` advanced service on top of UrlFetchApp
    gtmSchema.js              # GENERATED field/enum table from the GTM API v2 discovery doc
    tools/generateGtmSchema.js  # regenerates gtmSchema.js (curl or a local discovery JSON)
    realm.js, a1.js, dateFormat.js, index.js
    selftest.test.js          # proves the mocks work
    projectSmoke.test.js      # loads the REAL sources, checks onOpen() menu wiring
    fixtures/miniProject/     # tiny project used by the loader tests
```

## Test suites

| Directory | What it tests |
| --- | --- |
| `harness/` | The mocks themselves (`selftest.test.js`) and the menu wiring of the real sources (`projectSmoke.test.js`). |
| `features/` | The tool, end to end, against the fake GTM API: `unit.test.js` (text formats, parameters, trigger conditions), `web.test.js` and `server.test.js` (every feature sheet: list, edit, preview, create, copy, update, delete, validation), `robustness.test.js` (partial bulk failures, fallback to single requests, stale rows, rate limits, the execution time budget, API errors, Google tag containers). `features/helpers.js` simulates user actions (menu clicks, cell edits, copying rows). |
| `live/` | The tool against the **real** Tag Manager API (`live.test.js`), in a temporary workspace of each container that is deleted afterwards. Skipped unless `GTM_LIVE_CONTAINERS` and a token are available (see `live/liveAuth.js`). `live/probe.js` is a read-only tool that saves all entities of containers to JSON files. |

Live tests, for example:

```
clasp login --user gtm-test --extra-scopes https://www.googleapis.com/auth/tagmanager.readonly,https://www.googleapis.com/auth/tagmanager.edit.containers,https://www.googleapis.com/auth/tagmanager.delete.containers
GTM_LIVE_CONTAINERS=GTM-AAAAAAA,GTM-BBBBBBB HARNESS_TEST_FILTER=live node tests
```

## Running

| Command | Notes |
| --- | --- |
| `node --test tests/` | Official Node 22 builds. Node resolves the directory to `tests/index.js`, which loads every `*.test.js`. |
| `node tests` | Works on **any** Node 22, including Debian/Ubuntu's `libnode` packaging, where every `node --test ...` invocation crashes with `Missing internal module 'internal/deps/brace-expansion'` (a packaging bug, not ours). |
| `node --test` (repo root or `tests/`) / `cd tests && npm test` | Node's default patterns find `**/*.test.js`. |
| `node tests/harness/selftest.test.js` | A single file. |

Environment variables:

| Variable | Effect |
| --- | --- |
| `HARNESS_TEST_FILTER=<regex>` | `node tests` only loads matching test files. |
| `HARNESS_ECHO_LOGS=1` | Echo the script's `console.*` / `Logger.log` output (default: captured in `runtime.logs` only). |
| `HARNESS_SKIP_PROJECT=1` | Skip `projectSmoke.test.js` (e.g. while sources are mid-refactor). |
| `GTM_OAUTH_TOKEN` | Live mode: OAuth access token that `ScriptApp.getOAuthToken()` returns when `realHttp: true`. |
| `GTM_OAUTH_TOKEN_CMD` | Live mode: command that prints a token (cached 45 min), e.g. `gcloud auth application-default print-access-token`. |
| `GTM_QUOTA_PROJECT` | Live mode: adds `X-Goog-User-Project` to real requests to `*.googleapis.com` (needed for gcloud user credentials). |

### Live mode (real GTM API)

```js
const rt = createRuntime({realHttp: true});   // token from GTM_OAUTH_TOKEN / GTM_OAUTH_TOKEN_CMD
rt.spreadsheet.addSheet(...);
rt.call('listGtmAccounts');                    // real UrlFetchApp traffic through curl
```

The token needs the Tag Manager scopes the script uses (for example
`https://www.googleapis.com/auth/tagmanager.edit.containers`). The simplest
source is the Apps Script project itself: run
`function t() { Logger.log(ScriptApp.getOAuthToken()); }` in the editor and
`export GTM_OAUTH_TOKEN=...` (valid for about 1 hour). The alternative is gcloud
Application Default Credentials:

```sh
gcloud auth application-default login \
  --scopes=https://www.googleapis.com/auth/tagmanager.edit.containers,https://www.googleapis.com/auth/cloud-platform
export GTM_OAUTH_TOKEN_CMD="gcloud auth application-default print-access-token"
export GTM_QUOTA_PROJECT=<gcp-project-with-the-tag-manager-api-enabled>
```

The selftest's live case (`lists real GTM accounts`) runs only when one of the
token variables is set. Live tests mutate real containers, so point them at a
throw-away container or workspace.

If you use the clasp-based helper in `tests/live/liveAuth.js`, pass it in
explicitly:
`createRuntime({realHttp: true, oauthToken: () => getLiveAccessToken(), realHttpHeaders: getLiveExtraHeaders()})`.

Security: curl gets its whole configuration (URL, headers including
`Authorization`, and payload) on **stdin** (`curl -K -`), so tokens never show
up on the command line, in `ps` output or in error messages. `runtime.fetchLog`
redacts credential headers and any token value that `ScriptApp.getOAuthToken()`
returned.

## Writing tests

```js
const assert = require('node:assert/strict');
const {test} = require('node:test');
const {createRuntime, createFakeGtmApi} = require('./harness');

test('lists tags into the sheet', () => {
  const fake = createFakeGtmApi();
  const {workspacePath} = fake.seedWorkspace({
    accountId: '6000001', containerId: '9000001', workspaceId: '1',
    tags: [{name: 'GA4 - purchase', type: 'gaawe', parameter: []}],
  });
  const rt = createRuntime({fetchHandler: fake});          // loads features/, shared/, utils/
  rt.spreadsheet.addSheet('GTM Workspace', [['Name', 'Path', 'Select'], ['WS', workspacePath, true]]);
  rt.ui.queueResponse('YES');                               // next alert() answers YES

  rt.call('someMenuFunction');

  assert.deepEqual(rt.ui.alerts.map((a) => a.prompt), [/* ... */]);
  assert.deepEqual(rt.spreadsheet.getSheetValues('Tags')[1], [/* ... */]);
  assert.equal(fake.requestsTo('PUT').length, 0);
  assert.deepEqual(rt.warnings, []);                         // no suspicious API usage
});
```

Tips:

- `rt.call(fn, ...args)` converts arguments into the script's realm and converts
  the result back to plain host data, so `assert.deepStrictEqual` works.
  `rt.callRaw` skips the conversions.
- Read top-level `const`/`let` values with `rt.get('NAME')` (they are not
  properties of `rt.context`).
- Spreadsheet values written by the script are auto-typed like the Sheets UI:
  `'123'` becomes `123`. Account and tag ids written as strings come back as
  numbers, which is real Sheets behavior and a common source of bugs.
- `Utilities.sleep()` is virtual: it advances a clock that `Date.now()` and
  `new Date()` inside the script also follow, so retry and backoff code runs
  instantly. You can also move time forward from the test with
  `rt.advanceClock(ms)`.
- In Apps Script, every menu click is a **new execution**: top-level code runs
  again and globals start fresh. One runtime keeps its globals across
  `rt.call()`s, so use `rt = rt.nextExecution()` between simulated clicks. The
  new runtime keeps the spreadsheet, properties, cache, clock, queued UI
  responses and the same `fetchHandler` (so the fake GTM state survives).
- Call `onOpen`/`onEdit` with `rt.callTrigger('onOpen', event)` to get
  simple-trigger semantics. The function runs in `AuthMode.LIMITED`, where
  `UrlFetchApp` (and therefore the API) throws "You do not have permission to
  call UrlFetchApp.fetch ...", as it does in Apps Script.

## `createRuntime(options)` API

`const {createRuntime} = require('./harness/appsScriptRuntime');`

| Option | Default | Description |
| --- | --- | --- |
| `rootDir` | repo root | Where `dirs`, `files` and `appsscript.json` are resolved. |
| `dirs` | `['features','shared','utils']` | Directories scanned recursively for `.js`/`.gs`, loaded in relative-path (code-point) order. |
| `files` | – | Explicit list of files (absolute or relative to `rootDir`) in load order; replaces `dirs`. |
| `sources` | `[]` | Extra `{filename, code}` scripts loaded after the files (always, even with `load:false`). |
| `load` | `true` | `false` skips the project files on disk. |
| `loadErrors` | `'throw'` | `'collect'` records errors in `runtime.loadErrors` and keeps going. |
| `manifest` | `rootDir/appsscript.json` | Manifest object (`null` = none). Supplies `timeZone` and enabled advanced services. |
| `advancedServices` | `'auto'` | `'auto'` installs the `TagManager` mock if the manifest enables `tagmanager`. Can also be `true`, `false` or `{TagManager: 'Symbol'}`. |
| `fetchHandler` | – | `(req) => ({code, body, headers}) \| undefined`, or an object with `handle(req)` such as a `FakeGtmApi`. Returning `undefined` falls through to real HTTP if that is on; otherwise you get a harness error. Return `{networkError: 'Timeout'}` to simulate a transport failure. |
| `realHttp` | `false` | Send requests the handler doesn't answer to the network through `curl`. |
| `curlPath`, `httpTimeoutSeconds` | `'curl'`, `120` | curl binary and the `--max-time` value. |
| `realHttpHeaders` | – | Extra transport headers for real requests (the script doesn't see them). |
| `oauthToken` | `'test-oauth-token'` | String or function used by `ScriptApp.getOAuthToken()`. With `realHttp` and no option, the token comes from the env variables above. |
| `userEmail`, `effectiveUserEmail` | `test.user@example.com` | `Session` users. |
| `timeZone` | manifest, else `America/New_York` | `Session.getScriptTimeZone()`. |
| `spreadsheet` | – | `{id, name, timeZone, locale, sheets, fixture}`. `sheets` is either `{name: rows2D}` or `[{name, values, frozenRows, ...}]`; `fixture` is a `toJSON()`-shaped object. |
| `uiResponses`, `uiHandler` | `[]`, – | Initial `Ui` response queue, and a fallback responder `(dialog) => response`. |
| `properties` | – | Initial `{document, user, script}` property maps. |
| `realSleep` | `false` | Make `Utilities.sleep` really block (with `Atomics.wait`). |
| `virtualTime` | `true` | Make the script's `Date` follow the virtual clock. |
| `echoLogs` | `HARNESS_ECHO_LOGS==='1'` | Echo script logs to stdout and stderr. |
| `strictBounds` | `false` | Throw "The coordinates of the range are outside the dimensions of the sheet." instead of growing the sheet. |
| `enforceDataValidation` | `true` | Reject writes that violate `setAllowInvalid(false)` rules. |
| `enforceQuotas` | `true` | Enforce the PropertiesService value (9 KB) and store (500 KB) limits. |
| `formulaEvaluator` | – | `(formula, {sheet,row,column}) => value \| undefined`, for formulas the harness can't evaluate. |
| `timeoutMs` | – | `vm` timeout for `call()` and `run()`, to stop runaway loops. |
| `globals` | `{}` | Extra or overriding globals, for example a `DriveApp` stub. |
| `fetchErrorTruncate`, `fetchLogPayloadLimit` | `300`, `2000` | Truncation lengths for the "Truncated server response" text and for log payloads. |
| `htmlFiles`, `scriptId`, `userLocale` | – | Inputs for `HtmlService.createHtmlOutputFromFile`, `ScriptApp.getScriptId` and `Session.getActiveUserLocale`. |

The returned runtime object has these members:

| Member | Description |
| --- | --- |
| `context` | The `vm` context, which is the script's global object. |
| `run(code, filename?)` | Evaluate code in the context. Returns a raw value. |
| `call(fnName, ...args)` / `callRaw(...)` | Call a global function (dotted paths are allowed), with or without realm conversion. |
| `callTrigger(fnName, event?, {authMode = 'LIMITED'}?)` | Simple-trigger call: the default event `{authMode, source, user, triggerUid}` is merged with `event`, and services that need authorization throw while it runs. |
| `nextExecution(overrides?)` | Start a new runtime that emulates the next execution: fresh globals, with persistent state carried over (see Tips). |
| `get(name)` / `getRaw(name)` / `has(name)` / `isFunction(name)` | Read globals, including top-level `const`. |
| `loadFile(path)` | Load one more source file. |
| `spreadsheet` | Fixture helpers (below). `spreadsheet.active` is the mock `Spreadsheet` and `spreadsheet.app` is `SpreadsheetApp`. |
| `ui` | Recorder and controller: `alerts`, `prompts`, `toasts`, `menus`, `dialogs`, `history`, `queueResponse(...)`, `menuItems()`, `menuFunctionNames()`, `findMenuItem(caption)`, `reset()`, `instance` (the `Ui`). |
| `fetchLog` | `[{index, method, url, status, durationMs, mode, requestHeaders, contentType, payload, payloadBytes, responsePreview, responseBytes, error}]` with secrets redacted. |
| `logs` / `logText()` | `[{source: 'console'\|'Logger', level, message, time}]`. |
| `warnings` | Suspicious usage the mocks noticed: `non-primitive-cell-value`, `formula-not-evaluated`, `object-payload-with-json-content-type`, `payload-ignored`, `non-string-property`. |
| `harnessErrors` | Problems in the harness itself, such as a throwing `fetchHandler` or a request with no handler. |
| `sleeps` | `{totalMs, calls}` for `Utilities.sleep`. |
| `properties` | `{document, user, script}` stores with `.data` (a Map), `get`, `set`, `all()` and `clear()`. |
| `cache` | `{document, user, script}` stores with `.entries`, `get` and `all()`. |
| `locks` | `simulateContention(bool, scope?)` and `events`. |
| `triggers` | Triggers created with `ScriptApp.newTrigger`. |
| `globals` | The installed mock services. |
| `loadedFiles`, `loadErrors` | What was loaded and what failed. |
| `now()`, `advanceClock(ms)` | The virtual clock. |
| `setFetchHandler(h)` | Swap the transport at runtime. |
| `menuFunctionNames()`, `missingMenuFunctions()` | Menu items whose function is undefined. |
| `globalFunctionNames()` | Function declarations from the project code. |
| `resetRecordings()` | Clear the logs, fetch log, warnings and UI records. |
| `toHost(v)`, `toContext(v)`, `realm` | Realm helpers. |

`runtime.spreadsheet` fixture helpers return plain host data:

- `addSheet(name, rows2D, {autoType, index, frozenRows, frozenColumns, maxRows, maxColumns, activate})` creates the sheet or replaces its content. Values are stored raw unless `autoType` is set.
- `getSheet(name)`, `hasSheet(name)`, `removeSheet(name)`, `sheetNames()`.
- `getSheetValues(name)`, `getSheetDisplayValues(name)`, `getRowValues(name, row)` read the data range.
- `getCell(name, 'B3')` returns `{value, display, formula, note, validation, format, link}`.
- `dumpSheet(name, {compress, display})` returns values plus formulas, notes, validations and non-default formats. Identical cells are compressed into A1 ranges, for example `{'C2:C9': {criteria: 'CHECKBOX', ...}}`. It also returns hidden rows and columns, widths, merges, borders, filter, conditional formats, protections and frozen panes.
- `printSheet(name)` returns an ASCII table for debugging.
- `toJSON()` and `loadFixture(jsonOrObject)` round-trip the whole spreadsheet. Dates are encoded as `{$date: iso}` and `VALUE_IN_RANGE` sources as `{$range: 'Sheet!A1:A9'}`. `loadFixture` mutates in place, so a top-level `const ss = SpreadsheetApp.getActive()` in the script stays valid.

## `createFakeGtmApi(options)` API

`const {createFakeGtmApi} = require('./harness/fakeGtmApi');` (the module also
exports `FakeGtmApi`, `ApiError`, `GTM_API_BASE`, `ENTITY_KINDS`,
`BUILT_IN_TRIGGER_IDS`, `BUILT_IN_VARIABLE_TYPES` and `normalizeKind`)

| Option | Default | Description |
| --- | --- | --- |
| `pageSize` | `200` | Items per list page. Client `pageSize` values are ignored, like the real API. |
| `rateLimit` | – | `{max, windowMs, code = 429}`: sliding window, measured on the runtime's virtual clock (`req.timestamp`). Excess requests get `RESOURCE_EXHAUSTED` / `rateLimitExceeded`. |
| `requireAuth` | `true` | Requests without `Authorization: Bearer ...` get 401 `UNAUTHENTICATED`. |
| `acceptedTokens` | – | If set, other tokens get 401. |
| `strictContentType` | `true` | A POST/PUT body without `application/json` gets 400 `Invalid JSON payload received. Unknown name "..."`. This is what happens when a script forgets `contentType`. |
| `validateReferences` | `true` | `parentFolderId` and `firing`/`blocking`/`enabling`/`disablingTriggerId` must exist. The built-in trigger ids `2147479553`, `2147479572` and `2147479573` are allowed. Seeded data is checked too, so broken fixtures fail fast. |
| `validateSchema` | `true` | Reject request fields and enum values the API doesn't know, using a table generated from the official discovery document (`gtmSchema.js`). Examples: `Invalid JSON payload received. Unknown name "firingTriggerIds" at 'tag': Cannot find field.` and `Invalid value at 'tag.tag_firing_option' (TYPE_ENUM), "everyTime"`. This includes unknown keys in `bulk_update` changes (templates use `customTemplate`) and non-array values for repeated fields. Seeds are checked as well. |
| `omitEmptyFields` | `true` | Strip top-level `false`, `''`, `null` and `[]` fields from stored entities, as the API omits them. |
| `maxWorkspacesPerContainer` | `3` | Workspace limit for a standard container. |
| `fingerprintMismatchCode` | `409` | Status for a `?fingerprint=` mismatch (`ABORTED`). |
| `builtInTriggerIds`, `now` | see above, `Date.now` | |

Endpoints (base `https://tagmanager.googleapis.com/tagmanager/v2/`):

- `accounts` (list), `accounts/{a}` (get, PUT).
- `accounts/containers:lookup?tagId=|destinationId=`.
- `accounts/{a}/containers` (list, create; new containers get a Default Workspace) and `accounts/{a}/containers/{c}` (get, PUT, DELETE).
- `.../destinations` (list), `.../destinations/{id}` (get) and `POST .../destinations:link?destinationId=` (moves the destination from any other container).
- `.../workspaces` (list, create) and `.../workspaces/{w}` (get, PUT, DELETE). A new workspace starts from the container's base version.
- `GET .../workspaces/{w}/status`, which diffs against the seeded base and returns `{}` when nothing changed.
- `POST .../workspaces/{w}/bulk_update`.
- `.../workspaces/{w}/{tags|triggers|variables|folders|clients|transformations|templates|zones|gtag_config}`: list, create, `/{id}` get/PUT/DELETE, and `/{id}:revert`.
- `.../folders/{id}:entities` and `.../folders/{id}:move_entities_to_folder`.
- `.../workspaces/{w}/built_in_variables`: list, `POST ?type=a&type=b`, `DELETE ?type=`, and `:revert?type=`.

Other routes return 501, or 405 for a known path with the wrong method. Non-GTM URLs return `undefined`, so they fall through to the next transport.

The fake follows these semantics:

- **List responses** use the real keys: `tag`, `trigger`, `variable`, `folder`, `client`, `transformation`, `template`, `zone`, `gtagConfig`, `builtInVariable`, `destination`, `workspace`, `container` and `account`, plus `nextPageToken`. The key is omitted when the list is empty.
- **Creates** assign server fields: ids from one per-container counter shared by all kinds, `accountId`, `containerId`, `workspaceId`, `path`, `fingerprint` (a strictly increasing `Date.now()` string) and `tagManagerUrl`. Reference ids are coerced to strings.
- **Updates** (PUT) replace the whole entity and check `?fingerprint=`.
- **Deletes** return 204 with an empty body.
- **Validation**: `name` is required except for `gtagConfig`, and `type` is required for tags, triggers, variables, clients, transformations and gtagConfig. Names must be unique per kind and workspace. Unknown ids return 404 `Not found or permission denied.`.
- **Errors** use the Google envelope: `{"error":{"code":400,"message":"...","errors":[{"message":"...","domain":"global","reason":"badRequest"}],"status":"INVALID_ARGUMENT"}}`, pretty-printed like the real API.
- **bulk_update** accepts `{changes:[{tag|trigger|variable|folder|client|transformation|customTemplate|zone|gtagConfig|builtInVariable: {...}, changeStatus:'added'|'updated'|'deleted'}]}`.
  - Added entities must use `new_N` ids. Other changes can reference those ids, even ones defined later in the list.
  - Real ids are allocated in request order.
  - The update is atomic: everything is validated against a copy first. On any error nothing is applied (the id counter included), and a 400 (or 404) message starts with `changes[i] (...)`.
  - The response is `{changes: [{tag: {...}, changeStatus}, ...]}` with added and updated entities in request order; deleted entities are not included.

Methods for seeding and assertions:

- `seed({accounts, containers, workspaces, entities: {[wsPath]: {tag|tags|...: [...], builtInVariable: ['pageUrl']}}, destinations})` creates missing parents automatically. Seeded entities form the workspace's base version.
- `seedWorkspace({accountId='6000001', containerId='9000001', workspaceId='1', accountName, containerName, workspaceName, publicId, usageContext=['web'], folders, triggers, variables, builtInVariables, tags, clients, transformations, templates, zones, gtagConfigs})` returns `{accountPath, containerPath, workspacePath, accountId, containerId, workspaceId, publicId}`. Entities without ids are numbered in that order, so folders come first.
- The `state` getter returns a deep-copied snapshot: `{accounts, containers, workspaces, entities: {[wsPath]: {tag: [...], ...}}, destinations}`.
- `list(wsPath, kind)`, `find(wsPath, kind, nameOrPredicate)` and `get(resourcePath)` return copies. Kind names are flexible: `'tag'`, `'tags'`, `'gtag_config'` and so on all work.
- `workspaceChanges(wsPath)` returns `[{kind, id, name, changeStatus, entity}]`.
- `requests` is `[{method, url, path, query, body, rawBody, status, response, timestamp}]`, and `requestsTo(method?, pathStringOrRegExp?)` filters it.
- `injectError({method?, path?, match?, code=500, message, status, reason, times=1, networkError?})` and `clearInjectedErrors()` make requests fail.
- `snapshot()` / `restore(snap)` and `reset()`.
- `handle(req)` and the bound `handler` let you call the fake directly.

## Fidelity notes and known limitations

- **insertCheckboxes()** sets **every** cell in the range to `false`, which is what the Apps Script documentation specifies. It doesn't only fill empty cells. Clearing a checkbox cell (`clearContent()` or `setValue('')`) leaves it unchecked (`false`), and `removeCheckboxes()` clears the checked/unchecked values.
- **Auto-typing** follows the en_US Sheets UI for plain numbers, thousands separators, percentages, a leading `'` and `TRUE`/`FALSE`. Dates and currency strings are *not* parsed. Strings starting with `=` become formulas, but only `HYPERLINK(...)`, string, number and boolean literals are evaluated. Other formulas read as `''`, produce a `formula-not-evaluated` warning, and can be handled with `formulaEvaluator`.
- **Display values and number formats** cover a subset: automatic, `@`, `0`, `0.00`, `#,##0(.00)`, percentages, and simple date patterns.
- **The grid grows when you write past `getMaxRows()`/`getMaxColumns()`**. Reads past the edge return `''`. Use `strictBounds` for the stricter behavior.
- **Object identity**: the mocks return the same `Sheet` object every time, while Apps Script returns new wrapper objects. Ranges keep fixed coordinates, as in Apps Script.
- **Sheet names**: `getSheetByName` is case-sensitive, but `insertSheet`/`setName` reject names that differ only by case.
- **Recorded, not enforced**: filters (rows are never hidden), conditional formats, banding, borders and protections. Data validation is enforced only for `setAllowInvalid(false)` rules whose criteria can be evaluated (lists, ranges, checkboxes, numbers, text). Notes are kept by `clear()`, as in Sheets.
- **UrlFetchApp**:
  - `fetchAll` runs requests sequentially.
  - Real-HTTP mode doesn't support binary payloads, and HTTP/2 header names arrive in lower case.
  - The "Truncated server response" cut-off (300 chars by default) is an approximation.
  - URL Fetch quotas aren't modeled.
- **Utilities**: `zip`, `unzip`, `parseDate` and RSA signatures aren't implemented. In `formatDate`, `z` produces Intl short zone names (for example `GMT+2` rather than `CEST`).
- **Not modeled**: the 6-minute execution limit (use `timeoutMs` to guard against loops) and the Apps Script OAuth consent flow.
- **Other globals**: `setTimeout` and the like don't exist, as in Apps Script. Services that aren't mocked (DriveApp, GmailApp, ...) raise a `ReferenceError` unless you provide them through `globals`.
- **Fake GTM API**:
  - Validation and error messages approximate the real API's wording. Status codes and envelopes are faithful.
  - Tag, trigger and variable parameters aren't validated against templates, and there's no compiler or versioning.
  - `create_version`, `sync`, `quick_preview`, environments, versions and user permissions return 501.
  - The fingerprint mismatch status (409) and the 3-workspace limit are configurable.
- **Debian's packaged Node 22** can't run `node --test` at all (see Running). Use `node tests` there.
