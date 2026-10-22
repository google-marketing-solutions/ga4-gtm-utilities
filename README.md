## GA4 Google Tag Manager Utilities

This is not an officially supported Google product.

GA4 GTM Utilities is a Google Sheets solution (Apps Script + spreadsheet) for
**bulk changes to Google Tag Manager containers**, both **client-side (web)
and server-side**. List tags, variables, triggers, clients and other items in
a sheet, edit them like any spreadsheet, then preview and apply the changes to
a GTM workspace in one go.

### What you can do

| Sheet | Containers | What you can do |
|---|---|---|
| **GTM Workspace** | all | Select the account, container and workspace (or find a container by its ID). Create a new workspace for your bulk changes. |
| **Tags** | web, server | Settings shared by **all tags** (Google and non-Google): name, folder, paused, firing and exception triggers, consent checks, firing option, priority, tag sequencing, schedule, notes. Copy rows to duplicate tags, delete tags. |
| **GA4 Event Tags** | web | Event name, measurement ID, event parameters, user properties, event settings variable, ecommerce. Create, update and delete. |
| **Event Parameters & User Properties** | web, server | **One row per parameter** of GA4 event tags, Google tags and destination tags, Google Ads / Floodlight tags, Google tag settings variables and server GA4 tags. Add, change, move and delete parameters across many tags at once. |
| **Google Tags** | web | Google tag and the **destination tags** of the unified Google tag / GTM experience (Google Analytics, Google Ads, Floodlight and Merchant Center destinations): tag/destination ID, configuration parameters, shared event parameters, settings variables. |
| **Google Ads & Floodlight Tags** | web | Google Ads conversion, remarketing, user-provided data and call conversion tags, Conversion Linker, Floodlight counter and sales tags. |
| **Server GA4 Tags** | server | Measurement ID, event name, parameters and user properties to include, exclude, add or edit. |
| **Server Ads & Floodlight Tags** | server | Google Ads conversion, remarketing and user-provided data tags, Conversion Linker and Floodlight tags. |
| **Clients** | server | All clients: name, type, priority, GA4 client cookie settings, GTM web container client settings, region-specific settings. |
| **Transformations** | server | Allow, exclude and augment transformations: parameters, affected tags, priority. |
| **Variables** | web, server | All variables with their main setting (e.g. data layer key, constant value, JavaScript code, lookup table), default value, case conversion and **where each variable is used**. Variables that are still used can't be deleted by accident. |
| **Triggers** | web, server | All triggers with their event name and conditions written as text (e.g. `{{Page Path}} contains /checkout`) and the tags that use them. |
| **Advanced Settings** | web, server | **One row per parameter of any tag, variable, client or transformation**, including non-Google tags and custom templates. |
| **Google Tag Settings** | web | Container-wide **Google tag settings** (the Settings tab of Tag Manager): automatic event detection, domains, user-provided data collection, internal traffic, unwanted referrals, session timeout, cookie settings, consent defaults. The linked destinations are listed in the title row. |
| **Changelog** | all | Every change made with the tool: when, who, what, and a link to the item in Tag Manager. |

### How it works

Every feature sheet works the same way:

1. **Select a workspace** in the *GTM Workspace* sheet: menu *GA4 GTM
   Utilities › Select workspace* (list accounts › containers › workspaces and
   tick *Select*, or use *Find container by ID*). Tip: use *Create new
   workspace* to keep bulk changes separate.
2. Open a feature sheet and use **List from GTM**.
3. **Edit** the cells. Edited rows are detected automatically.
   * To **create** an item, add a row and set the *Action* to `Create`. To
     **duplicate** an item with all its settings, copy its row (including the
     ID), change what you need (at least the name) and set the *Action* to
     `Create`.
   * To **delete** an item, set the *Action* to `Delete`.
4. **Preview changes** shows in the *Result* column exactly what will change.
5. **Apply changes to GTM**. The *Result* column and the *Changelog* sheet show
   what happened.
6. Review the changes in the Tag Manager workspace and **publish** them there.

Hover over the column headers to see the expected values.

**Formats**

* Lists (e.g. triggers, consent types): values separated by commas, or one
  per line if a value contains a comma.
* Tables (e.g. event parameters, lookup tables): one row per line in the
  format `name = value`.
* Trigger conditions: one condition per line, e.g.
  `{{Page URL}} matches RegEx (ignore case) ^https://shop\.`.
* Advanced Settings: lists and tables of other parameters are shown as JSON.

### Safety

* All changes are made in the selected **workspace**. Nothing is published.
* Updates are **diff-based**: only the cells you changed are sent to GTM;
  every other setting of the item (including settings that are not shown in
  the sheet) is kept unchanged.
* Rows listed from another container can't update or delete items, and
  several rows can't change the same item.
* Variables and triggers that are still used can't be deleted.
* Every change is recorded in the *Changelog* sheet.

### Unified Google tag and Tag Manager

The tool supports the latest Google tag / Tag Manager features:

* Google **destination tags** (`dest_ga`, `dest_aw`, `dest_dc`, `dest_mc`) in
  the *Google Tags* sheet, and their parameters in the *Event Parameters &
  User Properties* sheet.
* Container-wide **Google tag settings** (`gtag_config`) in the *Google Tag
  Settings* sheet, and the **destinations** linked to the container.
* Google tag containers (type *Google tag*) can be opened with *Find
  container by ID* (their G- or GT- ID) to view and edit their Google tag
  settings. Changes to Google tag containers take effect immediately (they
  have no workspaces).

### APIs, quotas and limits

The script calls the [Google Tag Manager API v2](https://developers.google.com/tag-platform/tag-manager/api/v2)
with the following OAuth scopes: read and edit Tag Manager containers, edit
the current spreadsheet, show dialogs, connect to external services and see
your email address (for the changelog). It can't publish containers or manage
users.

* Changes are applied with `workspaces.bulk_update` (many changes in one API
  call), so bulk changes are fast and stay within the
  [API quota](https://developers.google.com/tag-platform/tag-manager/api/v2/limits-quotas).
  If a change is rejected, only that change fails; the other changes are
  applied.
* Requests are rate limited automatically and retried when the API is busy.
* Apps Script runs are limited to 6 minutes. If a run takes too long, the
  remaining rows are marked as not applied; use *Apply changes* again to
  continue.
* Values longer than 50,000 characters (the limit of a cell) are shown
  truncated and can't be edited in the sheet.

### How to access the spreadsheet

It is strongly recommended to use the template spreadsheet:

1. Join [this group](https://groups.google.com/g/ga4-gtm-utilities-users) to
   gain access to the spreadsheet.
2. Create a copy of
   [this spreadsheet](https://docs.google.com/spreadsheets/d/1yZVgAlMJbCpZDyimdj0h4QHQbogLJZv83oXqRwT-A40/edit?resourcekey=0-aZ7q7_HASW9TZpFbquEZsw#gid=515466659).

If you use your own spreadsheet instead, add the code with
[clasp](https://github.com/google/clasp) (`clasp push`), reload the
spreadsheet and run *GA4 GTM Utilities › Tools › Set up / repair all sheets*.

### Development

* `features/`: one file per feature sheet (a declarative definition of its
  columns).
* `shared/`: the engine (listing, change detection, preview and apply), the
  executor (bulk updates), the sheet layout, the menu and the changelog.
* `utils/`: the GTM API client, workspace selection and helpers.
* `tests/`: Node.js tests that run the Apps Script code with mocks of the
  Apps Script services and a fake GTM API (`node --test tests/`), plus live
  tests against the real GTM API (see `tests/README.md`).

### Feedback

If you have feedback or notice bugs, please raise an issue in GitHub, post in
the group, or submit a response to
[this survey](https://docs.google.com/forms/d/e/1FAIpQLScHrZbNU2RZGMtcWTVVEsxe5ZzARFvjqFQziixNPUPCsNcUUQ/viewform).
Please note that this is not an officially supported Google product, so
support is not guaranteed.
