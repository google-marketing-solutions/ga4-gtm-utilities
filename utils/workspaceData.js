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
 * @fileoverview Cached access to the entities of a workspace and conversions
 * between IDs and names (triggers, folders, templates, ...).
 */

/** Prefix of placeholder IDs for folders that still need to be created. */
const NEW_FOLDER_PREFIX = '__new_folder__:';

/**
 * Entities of the selected workspace. Lists are loaded lazily from the API
 * and cached for the duration of the script execution.
 */
class WorkspaceData {
  /**
   * @param {!Object} ctx The workspace context (see getWorkspaceContext).
   */
  constructor(ctx) {
    this.ctx = ctx;
    this.cache_ = {};
    this.pendingFolders_ = {};
  }

  /**
   * Returns all entities of a kind in the workspace.
   * @param {string} kind A key of GTM_ENTITY_KINDS.
   * @return {!Array<!Object>}
   */
  list(kind) {
    if (!this.cache_[kind]) {
      this.cache_[kind] = workspaceSupportsKind(this.ctx, kind) ?
          gtmListEntities(this.ctx.workspacePath, kind) : [];
    }
    return this.cache_[kind];
  }

  /**
   * Forgets the cached entities of a kind.
   * @param {string} kind
   */
  invalidate(kind) {
    delete this.cache_[kind];
  }

  /**
   * Returns an entity by ID.
   * @param {string} kind
   * @param {string} id
   * @return {!Object|undefined}
   */
  getById(kind, id) {
    const idField = GTM_ENTITY_KINDS[kind].idField;
    const target = String(id);
    return this.list(kind).find(entity => String(entity[idField]) === target);
  }

  /**
   * Returns an entity by name (exact match).
   * @param {string} kind
   * @param {string} name
   * @return {!Object|undefined}
   */
  getByName(kind, name) {
    const target = toText(name);
    return this.list(kind).find(entity => entity.name === target);
  }

  // -------------------------------------------------------------------------
  // Triggers
  // -------------------------------------------------------------------------

  /**
   * Returns the built-in triggers for the container type.
   * @return {!Array<{triggerId: string, name: string}>}
   */
  getBuiltInTriggers() {
    const context = this.ctx.isServer ? 'server' : 'web';
    return BUILT_IN_TRIGGERS.filter(trigger => trigger.context === context);
  }

  /**
   * Returns the name of a trigger.
   * @param {string} id
   * @return {string}
   */
  getTriggerName(id) {
    const target = String(id);
    const trigger = this.list('trigger').find(t => t.triggerId === target) ||
        BUILT_IN_TRIGGERS.find(t => t.triggerId === target);
    return trigger ? trigger.name : `Unknown trigger (${target})`;
  }

  /**
   * Converts trigger IDs to the text shown in a cell.
   * @param {?Array<string>|undefined} ids
   * @return {string}
   */
  triggerIdsToText(ids) {
    return formatList((ids || []).map(id => this.getTriggerName(id)));
  }

  /**
   * Converts the trigger names typed in a cell to trigger IDs.
   * @param {*} value
   * @return {!Array<string>}
   */
  triggerTextToIds(value) {
    const names = parseList(value);
    const missing = [];
    const ids = names.map(name => {
      const unknown = /^Unknown trigger \((\d+)\)$/.exec(name);
      if (unknown) return unknown[1];
      const trigger = this.list('trigger').find(t => t.name === name) ||
          this.getBuiltInTriggers().find(t => t.name === name) ||
          this.list('trigger').find(
              t => t.name.toLowerCase() === name.toLowerCase());
      if (!trigger) {
        missing.push(name);
        return null;
      }
      return trigger.triggerId;
    });
    if (missing.length) {
      throw new UserError(`Trigger not found: ${missing.join(', ')}. ` +
          'Create the trigger first or check the spelling.');
    }
    return ids;
  }

  // -------------------------------------------------------------------------
  // Folders
  // -------------------------------------------------------------------------

  /**
   * Returns the name of a folder ('' if the ID is empty or unknown).
   * @param {string|undefined} id
   * @return {string}
   */
  getFolderName(id) {
    if (!id) return '';
    if (String(id).indexOf(NEW_FOLDER_PREFIX) === 0) {
      return String(id).slice(NEW_FOLDER_PREFIX.length);
    }
    const folder = this.list('folder').find(f => f.folderId === String(id));
    return folder ? folder.name : '';
  }

  /**
   * Returns the ID of the folder with the given name. If the folder does not
   * exist, a placeholder ID is returned and the folder is created before the
   * changes are applied.
   * @param {*} name
   * @return {string} The folder ID, a placeholder, or '' for no folder.
   */
  getFolderIdForName(name) {
    const text = toText(name);
    if (!text) return '';
    const folder = this.list('folder').find(f => f.name === text) ||
        this.list('folder').find(
            f => f.name.toLowerCase() === text.toLowerCase());
    if (folder) return folder.folderId;
    this.pendingFolders_[text] = true;
    return NEW_FOLDER_PREFIX + text;
  }

  /**
   * Returns the names of folders that need to be created.
   * @return {!Array<string>}
   */
  getPendingFolderNames() {
    return Object.keys(this.pendingFolders_);
  }

  // -------------------------------------------------------------------------
  // Custom templates
  // -------------------------------------------------------------------------

  /**
   * Returns the name of the custom template used by a "cvt_" type.
   * @param {string} type E.g. "cvt_12345_67".
   * @return {string}
   */
  getTemplateName(type) {
    const match = /^cvt_(?:\d+_)?(\d+)$/.exec(type || '');
    if (!match) return '';
    const template = this.list('template').find(
        t => String(t.templateId) === match[1]);
    return template ? template.name : '';
  }

  /**
   * Returns the "cvt_" type of the custom template with the given name.
   * @param {string} name
   * @return {string}
   */
  getTemplateTypeByName(name) {
    const lower = toText(name).toLowerCase();
    const template = this.list('template').find(
        t => t.name.toLowerCase() === lower);
    return template ?
        `cvt_${this.ctx.containerId}_${template.templateId}` : '';
  }
}

/**
 * Returns true if the workspace supports an entity kind.
 * @param {!Object} ctx
 * @param {string} kind
 * @return {boolean}
 */
function workspaceSupportsKind(ctx, kind) {
  if (ctx.isGoogleTag) return kind === 'gtagConfig';
  switch (kind) {
    case 'client':
    case 'transformation':
      return ctx.isServer;
    case 'zone':
      return ctx.isWeb;
    case 'gtagConfig':
      return ctx.isWeb;
    default:
      return true;
  }
}
