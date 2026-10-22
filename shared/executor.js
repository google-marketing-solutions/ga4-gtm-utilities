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
 * @fileoverview Applies changes to a workspace.
 *
 * Changes are sent with workspaces.bulk_update, which applies many changes in
 * one API call (fast, and within the API rate limit). Because bulk_update is
 * atomic, a chunk that fails is split in two and retried, so that only the
 * changes with errors are rejected. If bulk_update is not available, the
 * changes are applied one by one.
 */

/**
 * Applies changes and sets the result of their operations.
 * @param {!Array<!Object>} changes Changes built by the sheet engine.
 * @param {!WorkspaceData} data
 * @param {!Object} ctx
 */
function executeChanges_(changes, data, ctx) {
  const ready = createPendingFolders_(changes, data, ctx);
  if (!ready.length) return;
  if (ctx.isGoogleTag) {
    runSingleChanges_(ready, ctx);
    return;
  }
  const remaining = runBulkChanges_(ready, ctx);
  if (remaining.length) runSingleChanges_(remaining, ctx);
}

/**
 * Creates the folders that are referenced by changes but do not exist yet.
 * @param {!Array<!Object>} changes
 * @param {!WorkspaceData} data
 * @param {!Object} ctx
 * @return {!Array<!Object>} The changes that can be applied.
 */
function createPendingFolders_(changes, data, ctx) {
  const names = {};
  changes.forEach(change => {
    const folderId = change.entity.parentFolderId;
    if (folderId && String(folderId).indexOf(NEW_FOLDER_PREFIX) === 0) {
      names[String(folderId).slice(NEW_FOLDER_PREFIX.length)] = true;
    }
  });
  const created = {};
  const failed = {};
  Object.keys(names).forEach(name => {
    try {
      const folder = gtmCreateEntity(ctx.workspacePath, 'folder',
          {name: name});
      created[NEW_FOLDER_PREFIX + name] = folder.folderId;
    } catch (e) {
      if (!(e instanceof GtmApiError)) throw e;
      failed[NEW_FOLDER_PREFIX + name] = e.message;
    }
  });
  if (Object.keys(created).length) data.invalidate('folder');
  return changes.filter(change => {
    const folderId = change.entity.parentFolderId;
    if (!folderId) return true;
    if (created[folderId]) {
      change.entity.parentFolderId = created[folderId];
      return true;
    }
    if (failed[folderId]) {
      setChangeError_(change, 'The folder could not be created: ' +
          failed[folderId]);
      return false;
    }
    return true;
  });
}

/**
 * Applies changes with workspaces.bulk_update.
 * @param {!Array<!Object>} changes
 * @param {!Object} ctx
 * @return {!Array<!Object>} Changes that still need to be applied one by one
 *     (because bulk_update is not supported).
 */
function runBulkChanges_(changes, ctx) {
  const chunks = chunkArray(changes, LIMITS.bulkUpdateChunkSize);
  for (let i = 0; i < chunks.length; i++) {
    if (runBulkChunk_(chunks[i], ctx) === 'unsupported') {
      return changes.filter(change => !change.done);
    }
  }
  return [];
}

/**
 * Applies one chunk of changes with bulk_update. On error, the chunk is split
 * in two halves that are retried separately.
 * @param {!Array<!Object>} chunk
 * @param {!Object} ctx
 * @return {string} 'ok' or 'unsupported'.
 */
function runBulkChunk_(chunk, ctx) {
  if (!chunk.length) return 'ok';
  if (getRemainingExecutionMs() < 20000) {
    chunk.forEach(setChangeTimeLimit_);
    return 'ok';
  }
  const payload = chunk.map((change, index) => {
    const info = GTM_ENTITY_KINDS[change.kind];
    const entity = deepCopy(change.entity);
    if (change.changeStatus === 'added') {
      entity[info.idField] = 'new_' + (index + 1);
    }
    const item = {changeStatus: change.changeStatus};
    item[info.bulkKey] = entity;
    return item;
  });
  try {
    const results = gtmBulkUpdate(ctx.workspacePath, payload);
    assignBulkResults_(chunk, results);
    return 'ok';
  } catch (e) {
    if (!(e instanceof GtmApiError)) throw e;
    if (isBulkUpdateUnsupported_(e)) return 'unsupported';
    if (isGlobalApiError_(e) || chunk.length === 1) {
      chunk.forEach(change => setChangeError_(change, e.message));
      return 'ok';
    }
    const middle = Math.ceil(chunk.length / 2);
    if (runBulkChunk_(chunk.slice(0, middle), ctx) === 'unsupported') {
      return 'unsupported';
    }
    return runBulkChunk_(chunk.slice(middle), ctx);
  }
}

/**
 * Returns true if an error means that bulk_update can't be used.
 * @param {!GtmApiError} e
 * @return {boolean}
 */
function isBulkUpdateUnsupported_(e) {
  return e.code === 501 ||
      /do not support bulk update|method not found|not implemented/i
          .test(e.message);
}

/**
 * Returns true if an error affects all changes (not a specific change).
 * @param {!GtmApiError} e
 * @return {boolean}
 */
function isGlobalApiError_(e) {
  return e.code === 401 || e.code === 429 || e.code >= 500 ||
      e.status === 'RESOURCE_EXHAUSTED' ||
      (e.code === 403 && /permission|insufficient|not enabled/i
          .test(e.message));
}

/**
 * Matches the entities returned by bulk_update with the changes.
 * @param {!Array<!Object>} chunk
 * @param {!Array<!Object>} results
 */
function assignBulkResults_(chunk, results) {
  const returned = results.map(getBulkChangeEntity).filter(Boolean);
  chunk.forEach(change => {
    const info = GTM_ENTITY_KINDS[change.kind];
    let entity = null;
    if (change.changeStatus === 'updated') {
      const id = String(change.entity[info.idField]);
      entity = (returned.find(r => r.kind === change.kind &&
          String(r.entity[info.idField]) === id) || {}).entity || null;
    } else if (change.changeStatus === 'added') {
      const match = returned.find(r => r.kind === change.kind && !r.used &&
          (change.entity.name ? r.entity.name === change.entity.name :
           r.entity.type === change.entity.type));
      if (match) {
        match.used = true;
        entity = match.entity;
      }
    }
    setChangeSuccess_(change, entity);
  });
}

/**
 * Applies changes one by one with the standard API methods.
 * @param {!Array<!Object>} changes
 * @param {!Object} ctx
 */
function runSingleChanges_(changes, ctx) {
  let stopMessage = '';
  changes.forEach(change => {
    if (change.done) return;
    if (stopMessage) {
      setChangeError_(change, stopMessage);
      return;
    }
    if (getRemainingExecutionMs() < 15000) {
      setChangeTimeLimit_(change);
      return;
    }
    try {
      if (change.changeStatus === 'added') {
        setChangeSuccess_(change, gtmCreateEntity(ctx.workspacePath,
            change.kind, change.entity));
      } else if (change.changeStatus === 'updated') {
        setChangeSuccess_(change, gtmUpdateEntity(change.entity.path,
            change.entity));
      } else {
        gtmDeleteEntity(change.entity.path);
        setChangeSuccess_(change, null);
      }
    } catch (e) {
      if (!(e instanceof GtmApiError)) throw e;
      setChangeError_(change, e.message);
      if (isGlobalApiError_(e)) stopMessage = e.message;
    }
  });
}

/**
 * Marks a change as successfully applied.
 * @param {!Object} change
 * @param {?Object} entity The entity returned by the API, if any.
 */
function setChangeSuccess_(change, entity) {
  change.done = true;
  change.resultEntity = entity || change.entity;
  const info = GTM_ENTITY_KINDS[change.kind];
  const id = entity ? String(entity[info.idField] || '') : '';
  change.ops.forEach(op => {
    if (op.result && op.result.status === 'error') return;
    let message;
    const diffs = op.diffs || change.diffs;
    if (op.type === 'delete') {
      message = change.changeStatus === 'deleted' ? '✔ Deleted.' :
          '✔ Removed.';
    } else if (op.type === 'create' && change.ops.length === 1 &&
               change.changeStatus === 'added') {
      message = '✔ Created' + (id ? ` (ID ${id}).` : '.');
    } else if (op.type === 'create') {
      message = '✔ Added.';
    } else {
      message = '✔ Updated: ' + formatDiffs_(diffs, false) + '.';
    }
    if (op.warnings && op.warnings.length) {
      message += ' ' + op.warnings.join(' ');
    }
    op.result = {status: 'success', message: message};
    if (op.type === 'create' && id) op.result.id = id;
    if (op.newMeta !== undefined) op.result.meta = op.newMeta;
  });
}

/**
 * Marks a change as failed.
 * @param {!Object} change
 * @param {string} message
 */
function setChangeError_(change, message) {
  change.done = true;
  change.ops.forEach(op => {
    op.result = {status: 'error', message: '✖ ' + message};
  });
}

/**
 * Marks a change as not applied because the execution time is running out.
 * @param {!Object} change
 */
function setChangeTimeLimit_(change) {
  change.done = true;
  change.ops.forEach(op => {
    op.result = {status: 'warning', message: '⚠ Not applied: the time limit ' +
        'of this run was reached. Use "Apply changes" again to continue.'};
  });
}
