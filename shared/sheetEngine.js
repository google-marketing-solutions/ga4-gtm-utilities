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
 * @fileoverview Generic engine behind every feature sheet.
 *
 * A feature sheet is described by a definition (see features/*.js):
 * - mode 'entity': every row is one GTM entity (tag, variable, trigger, ...).
 * - mode 'item': every row is one item inside an entity (a parameter or a
 *   table row, e.g. a GA4 event parameter).
 *
 * Every feature sheet has the same structure:
 *   Action | Result | <definition columns> | ID | _sync (hidden)
 * The hidden _sync column stores a hash of the row as it was listed, which is
 * used to detect edited rows, plus metadata that identifies list items.
 *
 * Updates are diff-based: only cells whose value differs from the current
 * value in GTM are applied, everything else in the entity is kept unchanged.
 */

/** Column keys of the system columns. */
const SYSTEM_COLUMNS = {
  action: {key: '_action', header: 'Action', width: 80},
  result: {key: '_result', header: 'Result', width: 260},
  id: {key: '_id', header: 'ID', width: 60},
  sync: {key: '_sync', header: '_sync (do not edit)', width: 60}
};

/**
 * Returns all the columns of a sheet in order (system columns included).
 * @param {!Object} def
 * @return {!Array<!Object>}
 */
function getAllColumns_(def) {
  return [SYSTEM_COLUMNS.action, SYSTEM_COLUMNS.result]
      .concat(def.columns)
      .concat([Object.assign({}, SYSTEM_COLUMNS.id,
                             {header: def.idHeader || 'ID'}),
               SYSTEM_COLUMNS.sync]);
}

/**
 * Returns the 1-based column index of a column key.
 * @param {!Object} def
 * @param {string} key
 * @return {number}
 */
function getColumnIndex_(def, key) {
  const index = getAllColumns_(def).findIndex(col => col.key === key);
  if (index === -1) throw new Error(`Unknown column ${key} in ${def.key}`);
  return index + 1;
}

/**
 * Returns the canonical text of a cell value for comparisons.
 * @param {!Object} col
 * @param {*} value
 * @return {string}
 */
function canonicalValue_(col, value) {
  switch (col.kind) {
    case 'checkbox':
      return toBool(value) ? 'true' : 'false';
    case 'list':
      return formatList(parseList(value));
    case 'table':
      return formatTable(parseTable(value));
    default:
      return toText(value);
  }
}

/**
 * Returns true if a column can be edited.
 * @param {!Object} col
 * @return {boolean}
 */
function isEditableColumn_(col) {
  return !col.readOnly && typeof col.set === 'function';
}

/**
 * Computes the hash of a row from its editable values and its ID.
 * @param {!Object} def
 * @param {!Object<string, *>} valuesByKey
 * @param {string} id
 * @return {string}
 */
function computeRowHash_(def, valuesByKey, id) {
  const parts = def.columns.filter(col => !col.readOnly)
      .map(col => canonicalValue_(col, valuesByKey[col.key]));
  parts.push(toText(id));
  return hashString(parts.join('\u0001'));
}

/**
 * Parses the content of the hidden _sync cell.
 * @param {*} value
 * @return {{h: string, m: (?Object|undefined)}}
 */
function parseSyncCell_(value) {
  const text = toText(value);
  if (!text) return {h: ''};
  try {
    const parsed = JSON.parse(text);
    return {h: parsed.h || '', m: parsed.m};
  } catch (e) {
    return {h: ''};
  }
}

/**
 * Builds the content of the hidden _sync cell.
 * @param {string} hash
 * @param {?Object=} meta
 * @return {string}
 */
function buildSyncCell_(hash, meta) {
  const sync = {h: hash};
  if (meta) sync.m = meta;
  return JSON.stringify(sync);
}

/**
 * Normalises the text typed in the Action column.
 * @param {*} value
 * @return {string} 'create', 'update', 'delete', '' or 'invalid'.
 */
function parseAction_(value) {
  const text = toText(value).toLowerCase();
  if (!text) return '';
  if (text === 'create' || text === 'new' || text === 'add') return 'create';
  if (text === 'update' || text === 'edit' || text === 'modify') {
    return 'update';
  }
  if (text === 'delete' || text === 'remove') return 'delete';
  return 'invalid';
}

/**
 * Throws if a sheet cannot be used with the selected container.
 * @param {!Object} def
 * @param {!Object} ctx
 */
function assertSheetSupportsContainer_(def, ctx) {
  const contextName = ctx.isGoogleTag ? 'googleTag' :
      (ctx.isServer ? 'server' : 'web');
  if (def.contexts.indexOf(contextName) === -1) {
    const names = {web: 'web', server: 'server', googleTag: 'Google tag'};
    throw new UserError(`The "${def.sheetName}" sheet works with ` +
        `${def.contexts.map(c => names[c]).join(' and ')} containers, but ` +
        `the selected container "${ctx.containerName}" is a ` +
        `${names[contextName]} container.`);
  }
}

/**
 * Returns the feature sheet for a definition, creating it if needed.
 * @param {!Object} def
 * @return {!SpreadsheetApp.Sheet}
 */
function getFeatureSheet_(def) {
  let sheet = SpreadsheetApp.getActive().getSheetByName(def.sheetName);
  if (!sheet) {
    setupFeatureSheet_(def);
    sheet = SpreadsheetApp.getActive().getSheetByName(def.sheetName);
  }
  return sheet;
}

/**
 * Returns true if the header row of a feature sheet matches the columns of
 * its definition (it may not, e.g. after an update of the tool).
 * @param {!Object} def
 * @param {!SpreadsheetApp.Sheet} sheet
 * @return {boolean}
 */
function hasCurrentLayout_(def, sheet) {
  const columns = getAllColumns_(def);
  if (sheet.getMaxColumns() < columns.length) return false;
  const headers = sheet.getRange(LAYOUT.HEADER_ROW, 1, 1, columns.length)
      .getValues()[0].map(toText);
  return columns.every((col, index) => headers[index] === col.header);
}

// ---------------------------------------------------------------------------
// Listing
// ---------------------------------------------------------------------------

/**
 * Builds the rows of a sheet from the workspace entities.
 * @param {!Object} def
 * @param {!WorkspaceData} data
 * @param {!Object} ctx
 * @return {!Array<{values: !Object<string, *>, notApplicable: !Object,
 *     id: string, meta: (?Object|undefined)}>}
 */
function buildSheetRows_(def, data, ctx) {
  const rows = [];
  if (def.mode === 'item') {
    def.kinds.forEach(kind => {
      if (!workspaceSupportsKind(ctx, kind)) return;
      const idField = GTM_ENTITY_KINDS[kind].idField;
      data.list(kind)
          .filter(entity => def.isEligible(entity, kind, data, ctx))
          .sort((a, b) => compareText(a.name || a.type, b.name || b.type))
          .forEach(entity => {
            def.explode(entity, kind, data, ctx).forEach(item => {
              rows.push({values: item.values, notApplicable: {},
                         id: String(entity[idField]), meta: item.meta});
            });
          });
    });
    return rows;
  }
  const idField = GTM_ENTITY_KINDS[def.kind].idField;
  data.list(def.kind)
      .filter(entity => !def.filter || def.filter(entity, data, ctx))
      .sort(def.sort || ((a, b) => compareText(a.name, b.name)))
      .forEach(entity => {
        const values = {};
        const notApplicable = {};
        def.columns.forEach(col => {
          if (col.appliesTo && !col.appliesTo(entity, data, ctx)) {
            values[col.key] = '';
            notApplicable[col.key] = true;
          } else {
            const value = col.get(entity, data, ctx);
            values[col.key] = value === undefined || value === null ? '' :
                cellSafeValue(value);
          }
        });
        rows.push({values: values, notApplicable: notApplicable,
                   id: String(entity[idField] || '')});
      });
  return rows;
}

/**
 * Lists the entities of the selected workspace in a feature sheet.
 * @param {!Object} def
 * @return {string} Message for the user.
 */
function listFeatureSheet_(def) {
  const ctx = getWorkspaceContext();
  assertSheetSupportsContainer_(def, ctx);
  const sheet = getFeatureSheet_(def);
  if (hasCurrentLayout_(def, sheet)) {
    confirmDiscardPendingChanges_(def, sheet);
  } else {
    // The columns have changed (e.g. after an update of the tool): rebuild
    // the layout before listing.
    setupFeatureSheet_(def);
  }
  const data = new WorkspaceData(ctx);
  const rows = buildSheetRows_(def, data, ctx);
  writeSheetRows_(def, sheet, rows);
  setListedWorkspace_(def, ctx, rows.length, data);
  refreshDropdownLists_(data, ctx);
  sheet.activate();
  const label = def.itemLabelPlural || 'items';
  return rows.length ?
      `Listed ${rows.length} ${label} from "${ctx.containerName}" ` +
      `(${ctx.workspaceName}). Edit the rows, then use "Apply changes".` :
      `No ${label} were found in "${ctx.containerName}" ` +
      `(${ctx.workspaceName}). You can add new rows with Action "Create".`;
}

/**
 * Asks for confirmation before listing if the sheet has unapplied changes.
 * @param {!Object} def
 * @param {!SpreadsheetApp.Sheet} sheet
 */
function confirmDiscardPendingChanges_(def, sheet) {
  const rows = readSheetRows_(def, sheet);
  const pending = rows.filter(row => {
    const action = parseAction_(row.action);
    if (action === 'create' || action === 'update' || action === 'delete') {
      return true;
    }
    return row.sync.h && row.sync.h !==
        computeRowHash_(def, row.valuesByKey, row.id);
  });
  if (!pending.length) return;
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert('Discard changes?',
      `${pending.length} row(s) in "${def.sheetName}" have changes that ` +
      'were not applied to GTM yet. Listing again will replace them.\n\n' +
      'Do you want to continue?', ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) {
    throw new UserError('Cancelled. Your changes were kept.');
  }
}

/**
 * Writes the listed rows to a sheet, replacing the previous rows.
 * @param {!Object} def
 * @param {!SpreadsheetApp.Sheet} sheet
 * @param {!Array<!Object>} rows
 */
function writeSheetRows_(def, sheet, rows) {
  const columns = getAllColumns_(def);
  const width = columns.length;
  const firstRow = LAYOUT.FIRST_DATA_ROW;
  const oldLastRow = Math.max(sheet.getLastRow(), firstRow);
  const oldCount = oldLastRow - firstRow + 1;
  const clearRange = sheet.getRange(firstRow, 1, oldCount, width);
  clearRange.clearContent();
  clearRange.setBackground(null);
  clearRange.setFontColor(null);
  clearRange.setFontLine('none');
  columns.forEach((col, index) => {
    if (col.kind === 'checkbox') {
      sheet.getRange(firstRow, index + 1, oldCount, 1).clearDataValidations();
    }
  });
  if (!rows.length) return;
  const neededRows = firstRow + rows.length - 1;
  if (neededRows > sheet.getMaxRows()) {
    const oldMaxRows = sheet.getMaxRows();
    ensureSheetRows_(sheet, neededRows + 100);
    applyDataRowFormats_(def, sheet, oldMaxRows + 1,
        sheet.getMaxRows() - oldMaxRows);
  }

  // Checkbox columns: a checkbox only where the column applies.
  columns.forEach((col, index) => {
    if (col.kind !== 'checkbox') return;
    const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    const rules = rows.map(row => [row.notApplicable[col.key] ? null : rule]);
    sheet.getRange(firstRow, index + 1, rows.length, 1)
        .setDataValidations(rules);
  });

  const values = rows.map(row => columns.map(col => {
    if (col.key === SYSTEM_COLUMNS.action.key ||
        col.key === SYSTEM_COLUMNS.result.key ||
        col.key === SYSTEM_COLUMNS.sync.key) {
      return '';
    }
    if (col.key === SYSTEM_COLUMNS.id.key) return row.id;
    if (col.kind === 'checkbox') {
      return row.notApplicable[col.key] ? '' : toBool(row.values[col.key]);
    }
    return row.values[col.key];
  }));
  sheet.getRange(firstRow, 1, rows.length, width).setValues(values);

  // Grey background for cells that do not apply to the row's type.
  const backgrounds = rows.map(row => columns.map(col =>
    row.notApplicable[col.key] ? COLORS.notApplicable : null));
  if (rows.some(row => Object.keys(row.notApplicable).length)) {
    sheet.getRange(firstRow, 1, rows.length, width)
        .setBackgrounds(backgrounds);
  }

  // Hashes are computed from the values as stored in the sheet.
  const stored = readSheetRows_(def, sheet);
  const syncColumn = getColumnIndex_(def, SYSTEM_COLUMNS.sync.key);
  const syncValues = stored.map((row, index) => [buildSyncCell_(
      computeRowHash_(def, row.valuesByKey, row.id),
      rows[index] ? rows[index].meta : undefined)]);
  if (syncValues.length) {
    sheet.getRange(firstRow, syncColumn, syncValues.length, 1)
        .setValues(syncValues);
  }
  sheet.setRowHeightsForced(firstRow, rows.length, 21);
}

/**
 * Stores which workspace a sheet was listed from and shows it in the sheet.
 * @param {!Object} def
 * @param {!Object} ctx
 * @param {number} count Number of listed rows.
 * @param {!WorkspaceData} data
 */
function setListedWorkspace_(def, ctx, count, data) {
  const info = {
    workspacePath: ctx.workspacePath,
    containerPath: ctx.containerPath,
    time: new Date().toISOString()
  };
  PropertiesService.getDocumentProperties().setProperty(
      PROPERTY_KEYS.listedWorkspace + def.key, JSON.stringify(info));
  const sheet = getFeatureSheet_(def);
  const time = Utilities.formatDate(new Date(), ctx.timeZone,
      'yyyy-MM-dd HH:mm');
  const extra = def.listStatusExtra ? def.listStatusExtra(data, ctx) : '';
  sheet.getRange(LAYOUT.GROUP_ROW, 1).setValue(
      `${count} listed from ${ctx.publicId || ctx.containerName} › ` +
      `${ctx.workspaceName} (${time})${extra}`);
}

/**
 * Returns which workspace a sheet was listed from.
 * @param {!Object} def
 * @return {?{workspacePath: string, containerPath: string}}
 */
function getListedWorkspace_(def) {
  const text = PropertiesService.getDocumentProperties().getProperty(
      PROPERTY_KEYS.listedWorkspace + def.key);
  if (!text) return null;
  try {
    return JSON.parse(text);
  } catch (e) {
    return null;
  }
}

// ---------------------------------------------------------------------------
// Reading rows
// ---------------------------------------------------------------------------

/**
 * Reads the data rows of a feature sheet. Empty rows are skipped.
 * @param {!Object} def
 * @param {!SpreadsheetApp.Sheet} sheet
 * @return {!Array<!Object>}
 */
function readSheetRows_(def, sheet) {
  const columns = getAllColumns_(def);
  const firstRow = LAYOUT.FIRST_DATA_ROW;
  const lastRow = sheet.getLastRow();
  if (lastRow < firstRow) return [];
  const values = sheet.getRange(firstRow, 1, lastRow - firstRow + 1,
      columns.length).getValues();
  const rows = [];
  values.forEach((rowValues, index) => {
    const valuesByKey = {};
    columns.forEach((col, colIndex) => {
      valuesByKey[col.key] = rowValues[colIndex];
    });
    const hasData = def.columns.some(col => {
      const value = valuesByKey[col.key];
      return col.kind === 'checkbox' ? false : !isBlank(value);
    }) || !isBlank(valuesByKey[SYSTEM_COLUMNS.id.key]);
    const action = valuesByKey[SYSTEM_COLUMNS.action.key];
    if (!hasData && isBlank(action)) return;
    rows.push({
      rowNumber: firstRow + index,
      action: action,
      valuesByKey: valuesByKey,
      id: toText(valuesByKey[SYSTEM_COLUMNS.id.key]),
      sync: parseSyncCell_(valuesByKey[SYSTEM_COLUMNS.sync.key]),
      hasData: hasData
    });
  });
  return rows;
}

// ---------------------------------------------------------------------------
// Applying changes
// ---------------------------------------------------------------------------

/**
 * Determines the operation of every row.
 * @param {!Object} def
 * @param {!Array<!Object>} rows
 * @param {!Object} ctx
 * @return {!Array<!Object>} Operations: {row, type, auto, result}.
 */
function classifyRows_(def, rows, ctx) {
  const listed = getListedWorkspace_(def);
  const sameContainer = Boolean(listed) &&
      listed.containerPath === ctx.containerPath;
  const ops = [];
  rows.forEach(row => {
    const op = {row: row, type: '', auto: false, result: null,
                sameContainer: sameContainer};
    const action = parseAction_(row.action);
    const changed = Boolean(row.sync.h) && row.sync.h !==
        computeRowHash_(def, row.valuesByKey, row.id);
    if (action === 'invalid') {
      op.result = {status: 'error', message: 'Unknown action. Use Create, ' +
          'Update or Delete.'};
    } else if (action === '') {
      if (changed) {
        op.type = 'update';
        op.auto = true;
      } else if (row.hasData && !row.sync.h) {
        op.result = {status: 'warning', message: 'Not processed: set the ' +
            'Action to "Create" to add this row to GTM.'};
      }
    } else if (!row.hasData) {
      op.result = {status: 'error', message: 'The row is empty.'};
    } else {
      op.type = action;
    }
    if ((op.type === 'update' || op.type === 'delete') && !op.result) {
      const listedItem = def.mode === 'item' ? Boolean(row.sync.m) :
          Boolean(row.id);
      if (!listedItem) {
        op.result = {status: 'error', message: `This row was not listed ` +
            `from GTM, so it can't be ${op.type}d. Use "Create" for new ` +
            'rows.'};
      } else if (!sameContainer) {
        op.result = {status: 'error', message: 'This row was listed from ' +
            'another container. List the sheet again before updating or ' +
            'deleting.'};
      }
    }
    if (op.type || op.result) ops.push(op);
  });

  // Several rows must not update or delete the same item. All rows are
  // checked (not only edited ones): a copied row that was edited without
  // choosing "Create" would otherwise change the original item.
  const itemKey = row => def.mode === 'item' ?
      row.id + '|' + JSON.stringify(row.sync.m) : row.id;
  const rowsByItem = {};
  rows.forEach(row => {
    if (!row.id || parseAction_(row.action) === 'create') return;
    if (def.mode === 'item' && (!row.sync.m || row.sync.m.placeholder)) {
      return;
    }
    (rowsByItem[itemKey(row)] = rowsByItem[itemKey(row)] || []).push(row);
  });
  ops.forEach(op => {
    if (op.result || (op.type !== 'update' && op.type !== 'delete')) return;
    const sameItem = rowsByItem[itemKey(op.row)] || [];
    if (sameItem.length < 2) return;
    const rowNumbers = sameItem.map(row => row.rowNumber).join(', ');
    op.result = {status: 'error', message: `Rows ${rowNumbers} refer to ` +
        'the same item. To copy an item, set the Action of the copy to ' +
        '"Create"; otherwise remove the duplicate rows.'};
  });
  return ops;
}

/**
 * Applies the cell values of a row to an entity draft.
 * @param {!Object} def
 * @param {!Object} draft The entity to modify.
 * @param {?Object} reference The entity to compare with (the current entity
 *     for updates, the copied entity for copies, null for new entities).
 * @param {!Object} row
 * @param {!WorkspaceData} data
 * @param {!Object} ctx
 * @param {string} mode 'create' or 'update'.
 * @return {{diffs: !Array<!Object>, warnings: !Array<string>}}
 */
function applyRowToEntity_(def, draft, reference, row, data, ctx, mode) {
  const diffs = [];
  const warnings = [];
  const errors = [];
  def.columns.forEach(col => {
    if (!isEditableColumn_(col)) return;
    const cell = row.valuesByKey[col.key];
    const desired = canonicalValue_(col, cell);
    if (col.appliesTo && !col.appliesTo(draft, data, ctx)) {
      if (desired !== '' && desired !== 'false') {
        warnings.push(`"${col.header}" does not apply and was ignored.`);
      }
      return;
    }
    const fullValue = reference ? col.get(reference, data, ctx) : '';
    const shownValue = cellSafeValue(fullValue);
    const current = reference ? canonicalValue_(col, shownValue) : '';
    if (!reference && (desired === '' ||
        (col.kind === 'checkbox' && desired === 'false'))) {
      return;
    }
    if (desired === current) return;
    if (shownValue !== fullValue) {
      errors.push(`"${col.header}" is too long to be edited in a cell. Edit ` +
          'it in Tag Manager.');
      return;
    }
    if (mode === 'update' && col.createOnly) {
      errors.push(`"${col.header}" can't be changed for an existing item.`);
      return;
    }
    try {
      col.set(draft, cell, data, ctx);
      diffs.push({header: col.header, from: current, to: desired});
    } catch (e) {
      if (e instanceof UserError) {
        errors.push(e.message);
      } else {
        throw e;
      }
    }
  });
  if (errors.length) throw new UserError(errors.join(' '));
  return {diffs: diffs, warnings: warnings};
}

/**
 * Checks the required values of a new entity.
 * @param {!Object} def
 * @param {!Object} draft
 * @param {!Object} row
 * @param {!WorkspaceData} data
 * @param {!Object} ctx
 */
function checkRequiredColumns_(def, draft, row, data, ctx) {
  const missing = def.columns.filter(col => {
    if (!col.required) return false;
    if (col.appliesTo && !col.appliesTo(draft, data, ctx)) return false;
    return canonicalValue_(col, col.get(draft, data, ctx)) === '';
  }).map(col => col.header);
  if (missing.length) {
    throw new UserError(`Missing required value(s): ${missing.join(', ')}.`);
  }
}

/**
 * Builds the changes for an entity-mode sheet.
 * @param {!Object} def
 * @param {!Array<!Object>} ops
 * @param {!WorkspaceData} data
 * @param {!Object} ctx
 * @return {!Array<!Object>} Changes.
 */
function buildEntityChanges_(def, ops, data, ctx) {
  const kind = def.kind;
  const idField = GTM_ENTITY_KINDS[kind].idField;
  const changes = [];
  const newNames = {};
  ops.forEach(op => {
    if (op.result || !op.type) return;
    const row = op.row;
    try {
      if (op.type === 'delete') {
        const current = data.getById(kind, row.id);
        if (!current) {
          op.result = {status: 'warning', message: 'Already deleted (not ' +
              'found in the workspace).'};
          return;
        }
        if (def.validateDelete) {
          const errors = def.validateDelete(current, data, ctx) || [];
          if (errors.length) throw new UserError(errors.join(' '));
        }
        changes.push({ops: [op], kind: kind, changeStatus: 'deleted',
                      entity: current, original: current, diffs: []});
        return;
      }
      if (op.type === 'update') {
        const current = data.getById(kind, row.id);
        if (!current) {
          throw new UserError('This item no longer exists in the workspace. ' +
              'List the sheet again.');
        }
        const draft = deepCopy(current);
        const result = applyRowToEntity_(def, draft, current, row, data, ctx,
            'update');
        if (!result.diffs.length) {
          op.result = {status: 'skipped', message: 'No changes.'};
          return;
        }
        if (draft.name !== current.name) {
          checkUniqueName_(kind, draft.name, row.id, data, newNames);
        }
        validateDraft_(def, draft, data, ctx);
        op.warnings = result.warnings;
        changes.push({ops: [op], kind: kind, changeStatus: 'updated',
                      entity: draft, original: current, diffs: result.diffs});
        return;
      }
      // Create: a copy of an existing entity, or a new entity.
      let draft;
      let source = null;
      if (row.id && op.sameContainer) {
        source = data.getById(kind, row.id);
        if (!source) {
          throw new UserError(`The ${def.itemLabel || 'item'} with ID ` +
              `${row.id} to copy was not found. Clear the ID to create a ` +
              'new item.');
        }
        draft = stripEntityIdentity(source, idField);
      } else if (def.newEntity) {
        draft = def.newEntity(row.valuesByKey, data, ctx);
      } else {
        throw new UserError(`To create a new ${def.itemLabel || 'item'} in ` +
            'this sheet, copy the row of an existing one (including its ID), ' +
            'change the values and set the Action to "Create".');
      }
      const result = applyRowToEntity_(def, draft, source, row, data, ctx,
          'create');
      if (!toText(draft.name)) throw new UserError('The name is required.');
      checkUniqueName_(kind, draft.name, '', data, newNames);
      checkRequiredColumns_(def, draft, row, data, ctx);
      validateDraft_(def, draft, data, ctx);
      newNames[draft.name] = true;
      op.warnings = result.warnings;
      changes.push({ops: [op], kind: kind, changeStatus: 'added',
                    entity: draft, original: null, source: source,
                    diffs: result.diffs});
    } catch (e) {
      if (!(e instanceof UserError)) throw e;
      op.result = {status: 'error', message: e.message};
    }
  });
  return changes;
}

/**
 * Throws if the name is already used by another entity of the same kind.
 * @param {string} kind
 * @param {string} name
 * @param {string} ownId ID of the entity being renamed ('' for new ones).
 * @param {!WorkspaceData} data
 * @param {!Object<string, boolean>} newNames Names of entities created in this
 *     run.
 */
function checkUniqueName_(kind, name, ownId, data, newNames) {
  const idField = GTM_ENTITY_KINDS[kind].idField;
  const existing = data.getByName(kind, name);
  if ((existing && String(existing[idField]) !== String(ownId)) ||
      newNames[name]) {
    throw new UserError(`The name "${name}" is already used by another ` +
        `${GTM_ENTITY_KINDS[kind].label.toLowerCase()}. Names must be ` +
        'unique.');
  }
}

/**
 * Runs the validation function of a definition.
 * @param {!Object} def
 * @param {!Object} draft
 * @param {!WorkspaceData} data
 * @param {!Object} ctx
 */
function validateDraft_(def, draft, data, ctx) {
  if (!def.validate) return;
  const errors = def.validate(draft, data, ctx) || [];
  if (errors.length) throw new UserError(errors.join(' '));
}

/**
 * Builds the changes for an item-mode sheet. Rows are grouped by entity so
 * that each entity is updated once.
 * @param {!Object} def
 * @param {!Array<!Object>} ops
 * @param {!WorkspaceData} data
 * @param {!Object} ctx
 * @return {!Array<!Object>} Changes.
 */
function buildItemChanges_(def, ops, data, ctx) {
  const groups = {};
  const order = [];
  ops.forEach(op => {
    if (op.result || !op.type) return;
    let target;
    try {
      target = def.findEntity(op.row, data, ctx, op);
    } catch (e) {
      if (!(e instanceof UserError)) throw e;
      op.result = {status: 'error', message: e.message};
      return;
    }
    const key = target.kind + '|' + (target.entity ?
        target.entity[GTM_ENTITY_KINDS[target.kind].idField] :
        'new:' + target.newKey);
    if (!groups[key]) {
      groups[key] = {target: target, ops: []};
      order.push(key);
    }
    groups[key].ops.push(op);
  });

  const changes = [];
  order.forEach(key => {
    const group = groups[key];
    const target = group.target;
    const draft = target.entity ? deepCopy(target.entity) :
        deepCopy(target.newEntity);
    const appliedOps = [];
    let diffs = [];
    group.ops.forEach(op => {
      try {
        let opDiffs;
        if (op.type === 'update') {
          opDiffs = def.updateItem(draft, op.row, data, ctx, target.kind);
        } else if (op.type === 'create') {
          opDiffs = def.createItem(draft, op.row, data, ctx, target.kind);
        } else {
          opDiffs = def.deleteItem(draft, op.row, data, ctx, target.kind);
        }
        if (!opDiffs.length) {
          op.result = {status: 'skipped', message: 'No changes.'};
          return;
        }
        if (op.row.newMeta !== undefined) op.newMeta = op.row.newMeta;
        op.diffs = opDiffs;
        diffs = diffs.concat(opDiffs);
        appliedOps.push(op);
      } catch (e) {
        if (!(e instanceof UserError)) throw e;
        op.result = {status: 'error', message: e.message};
      }
    });
    if (!appliedOps.length) return;
    try {
      if (def.validateEntity) {
        const errors = def.validateEntity(draft, target.kind, data, ctx) || [];
        if (errors.length) throw new UserError(errors.join(' '));
      }
    } catch (e) {
      if (!(e instanceof UserError)) throw e;
      appliedOps.forEach(op => {
        op.result = {status: 'error', message: e.message};
      });
      return;
    }
    changes.push({
      ops: appliedOps,
      kind: target.kind,
      changeStatus: target.entity ? 'updated' : 'added',
      entity: draft,
      original: target.entity || null,
      diffs: diffs
    });
  });
  return changes;
}

/**
 * Formats a list of differences.
 * @param {!Array<{header: string, from: string, to: string}>} diffs
 * @param {boolean} withValues Whether to include old and new values.
 * @return {string}
 */
function formatDiffs_(diffs, withValues) {
  return diffs.map(diff => {
    if (!withValues) return diff.header;
    const from = truncateText(diff.from.replace(/\n/g, '; ') || '(empty)', 60);
    const to = truncateText(diff.to.replace(/\n/g, '; ') || '(empty)', 60);
    return `${diff.header}: ${from} → ${to}`;
  }).join(withValues ? ' | ' : ', ');
}

/**
 * Previews or applies the changes made in a feature sheet.
 * @param {!Object} def
 * @param {{preview: boolean}} options
 * @return {string} Message for the user.
 */
function applyFeatureSheet_(def, options) {
  const preview = Boolean(options && options.preview);
  const ctx = getWorkspaceContext();
  assertSheetSupportsContainer_(def, ctx);
  const sheet = getFeatureSheet_(def);
  if (!hasCurrentLayout_(def, sheet)) {
    throw new UserError(`The columns of the "${def.sheetName}" sheet do not ` +
        'match this version of the tool. Use "List from GTM" to rebuild the ' +
        'sheet, then make your changes again.');
  }
  const rows = readSheetRows_(def, sheet);
  const ops = classifyRows_(def, rows, ctx);
  if (!ops.length) {
    return 'No changes found. Edit cells of listed rows, add rows with ' +
        'Action "Create" or set the Action to "Delete".';
  }
  const data = new WorkspaceData(ctx);
  const changes = def.mode === 'item' ?
      buildItemChanges_(def, ops, data, ctx) :
      buildEntityChanges_(def, ops, data, ctx);

  if (preview) {
    changes.forEach(change => {
      change.ops.forEach(op => {
        const diffs = op.diffs || change.diffs;
        const details = diffs.length ? ': ' + formatDiffs_(diffs, true) : '.';
        let message;
        if (op.type === 'delete') {
          message = def.mode === 'item' ? 'Will be removed' + details :
              'Will be deleted.';
        } else if (op.type === 'create') {
          message = def.mode === 'item' ? 'Will be added' + details :
              'Will be created' + (change.source ?
                  ` as a copy of "${change.source.name}"` : '') + details;
        } else {
          message = (op.auto ? 'Edited, will be updated' :
              'Will be updated') + details;
        }
        if (op.warnings && op.warnings.length) {
          message += ' ' + op.warnings.join(' ');
        }
        op.result = {status: 'preview', message: message};
      });
    });
    writeOperationResults_(def, sheet, ops, true);
    return summarizeOperations_(ops, true, ctx);
  }

  if (changes.length && !confirmApply_(def, changes, ops, ctx)) {
    return 'Cancelled. Nothing was changed.';
  }
  if (changes.length) {
    executeChanges_(changes, data, ctx);
  }
  writeOperationResults_(def, sheet, ops, false);
  logChanges_(changes, ctx);
  return summarizeOperations_(ops, false, ctx);
}

/**
 * Asks the user to confirm the changes.
 * @param {!Object} def
 * @param {!Array<!Object>} changes
 * @param {!Array<!Object>} ops
 * @param {!Object} ctx
 * @return {boolean}
 */
function confirmApply_(def, changes, ops, ctx) {
  const count = status => changes.filter(c => c.changeStatus === status)
      .reduce((total, c) => total + c.ops.length, 0);
  const errors = ops.filter(op => op.result && op.result.status === 'error')
      .length;
  const ui = SpreadsheetApp.getUi();
  let message = `Container: ${ctx.containerName} (${ctx.publicId})\n` +
      `Workspace: ${ctx.workspaceName}\n\n` +
      `Rows to create: ${count('added')}\n` +
      `Rows to update: ${count('updated')}\n` +
      `Rows to delete: ${count('deleted')}\n`;
  if (errors) message += `Rows with errors (skipped): ${errors}\n`;
  message += ctx.isGoogleTag ?
      '\nGoogle tag settings are applied immediately (there is no ' +
      'workspace or publishing step). Continue?' :
      '\nThe changes are made in the workspace only; nothing is published. ' +
      'Continue?';
  return ui.alert('Apply changes to GTM?', message, ui.ButtonSet.YES_NO) ===
      ui.Button.YES;
}

/**
 * Writes the results of the operations to the sheet.
 * @param {!Object} def
 * @param {!SpreadsheetApp.Sheet} sheet
 * @param {!Array<!Object>} ops
 * @param {boolean} preview
 */
function writeOperationResults_(def, sheet, ops, preview) {
  if (!ops.length) return;
  const firstRow = LAYOUT.FIRST_DATA_ROW;
  const lastRow = Math.max(...ops.map(op => op.row.rowNumber));
  const count = lastRow - firstRow + 1;
  const actionColumn = getColumnIndex_(def, SYSTEM_COLUMNS.action.key);
  const resultColumn = getColumnIndex_(def, SYSTEM_COLUMNS.result.key);
  const idColumn = getColumnIndex_(def, SYSTEM_COLUMNS.id.key);
  const syncColumn = getColumnIndex_(def, SYSTEM_COLUMNS.sync.key);

  const actionRange = sheet.getRange(firstRow, actionColumn, count, 2);
  const actionValues = actionRange.getValues();
  const resultColors = actionValues.map(() => [null]);
  const idRange = sheet.getRange(firstRow, idColumn, count, 2);
  const idValues = idRange.getValues();
  const deletedRows = [];
  const colors = {success: COLORS.success, error: COLORS.error,
                  warning: COLORS.error, preview: COLORS.preview,
                  skipped: COLORS.readOnlyText};

  ops.forEach(op => {
    const index = op.row.rowNumber - firstRow;
    const result = op.result || {status: 'skipped', message: ''};
    actionValues[index][1] = result.message;
    resultColors[index][0] = colors[result.status] || null;
    if (preview || result.status !== 'success') return;
    actionValues[index][0] = '';
    if (op.type === 'delete') {
      idValues[index][1] = '';
      deletedRows.push(op.row.rowNumber);
      return;
    }
    if (result.id) idValues[index][0] = result.id;
    const meta = result.meta !== undefined ? result.meta : op.row.sync.m;
    idValues[index][1] = buildSyncCell_(computeRowHash_(def,
        op.row.valuesByKey, idValues[index][0]), meta);
  });
  actionRange.setValues(actionValues);
  sheet.getRange(firstRow, resultColumn, count, 1).setFontColors(resultColors);
  idRange.setValues(idValues);
  deletedRows.forEach(rowNumber => {
    sheet.getRange(rowNumber, resultColumn + 1, 1,
        getAllColumns_(def).length - resultColumn)
        .setFontColor(COLORS.deleted)
        .setFontLine('line-through');
  });
}

/**
 * Builds the final message.
 * @param {!Array<!Object>} ops
 * @param {boolean} preview
 * @param {!Object} ctx
 * @return {string}
 */
function summarizeOperations_(ops, preview, ctx) {
  const count = status => ops.filter(op => op.result &&
      op.result.status === status).length;
  if (preview) {
    return `Preview: ${count('preview')} row(s) will change, ` +
        `${count('error')} error(s), ${count('skipped')} without changes. ` +
        'See the Result column. Nothing was changed in GTM.';
  }
  const parts = [`${count('success')} row(s) applied`];
  if (count('error')) parts.push(`${count('error')} error(s)`);
  if (count('warning')) parts.push(`${count('warning')} warning(s)`);
  if (count('skipped')) parts.push(`${count('skipped')} without changes`);
  return parts.join(', ') + ` in "${ctx.workspaceName}". See the Result ` +
      'column and the Changelog sheet.';
}
