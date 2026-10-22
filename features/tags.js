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
 * @fileoverview The "Tags" sheet: settings that all tags have in common, for
 * every tag type (Google and non-Google tags, web and server containers).
 */

/**
 * Returns the definition of the Tags sheet.
 * @return {!Object}
 */
function tagsSheetDef_() {
  return {
    key: 'tags',
    sheetName: 'Tags',
    title: 'ALL TAGS',
    tabColor: TAB_COLORS.general,
    mode: 'entity',
    kind: 'tag',
    contexts: ['web', 'server'],
    itemLabel: 'tag',
    itemLabelPlural: 'tags',
    idHeader: 'Tag ID',
    nameColumn: 'name',
    columns: [
      COLUMNS.name('Tag', {group: 'TAG'}),
      COLUMNS.typeReadOnly(TAG_TYPE_NAMES, 'Tag', {group: 'TAG'}),
      COLUMNS.folder({group: 'TAG'}),
      COLUMNS.paused({group: 'TAG'}),
      COLUMNS.triggers('firingTriggerId', {group: 'TRIGGERING'}),
      COLUMNS.triggers('blockingTriggerId', {group: 'TRIGGERING'}),
      consentStatusColumn_({group: 'CONSENT SETTINGS'}),
      consentTypesColumn_({group: 'CONSENT SETTINGS'}),
      firingOptionColumn_({group: 'ADVANCED SETTINGS'}),
      priorityColumn_({group: 'ADVANCED SETTINGS'}),
      scheduleColumn_('scheduleStartMs', {group: 'ADVANCED SETTINGS'}),
      scheduleColumn_('scheduleEndMs', {group: 'ADVANCED SETTINGS'}),
      liveOnlyColumn_({group: 'ADVANCED SETTINGS'}),
      sequenceTagColumn_('setupTag', {group: 'TAG SEQUENCING'}),
      sequenceStopColumn_('setupTag', {group: 'TAG SEQUENCING'}),
      sequenceTagColumn_('teardownTag', {group: 'TAG SEQUENCING'}),
      sequenceStopColumn_('teardownTag', {group: 'TAG SEQUENCING'}),
      COLUMNS.notes({group: 'NOTES'})
    ],
    validate: draft => validateConsentSettings_(draft)
  };
}

/**
 * Columns shared by the type-specific tag sheets.
 * @param {string=} group Group label.
 * @return {!Array<!Object>}
 */
function commonTagColumns_(group) {
  const opts = {group: group || 'TAG SETTINGS'};
  return [
    COLUMNS.triggers('firingTriggerId', opts),
    COLUMNS.triggers('blockingTriggerId', opts),
    COLUMNS.folder(opts),
    COLUMNS.paused(opts),
    COLUMNS.notes(opts)
  ];
}

/**
 * Column: consent status of a tag (Additional consent checks).
 * @param {!Object=} options
 * @return {!Object}
 */
function consentStatusColumn_(options) {
  const labels = CONSENT_STATUS_LABELS;
  return buildColumn_({
    key: 'consentStatus',
    header: 'Additional Consent Checks',
    width: 210,
    dropdown: Object.keys(labels).map(key => labels[key]),
    strictDropdown: true,
    note: 'Not set / No additional consent required / Require additional ' +
        'consent (then list the consent types in the next column).',
    appliesTo: (entity, data, ctx) => !ctx.isServer,
    get: entity => {
      const status = (entity.consentSettings &&
          entity.consentSettings.consentStatus) || 'notSet';
      return labels[status] || status;
    },
    set: (draft, value) => {
      const text = toText(value).toLowerCase();
      const status = Object.keys(labels).find(key =>
        labels[key].toLowerCase() === text || key.toLowerCase() === text) ||
        (text ? null : 'notSet');
      if (!status) {
        throw new UserError(`Unknown consent setting "${toText(value)}".`);
      }
      draft.consentSettings = draft.consentSettings || {};
      draft.consentSettings.consentStatus = status;
      if (status !== 'needed') delete draft.consentSettings.consentType;
    }
  }, options);
}

/**
 * Column: consent types required by a tag.
 * @param {!Object=} options
 * @return {!Object}
 */
function consentTypesColumn_(options) {
  return buildColumn_({
    key: 'consentTypes',
    header: 'Required Consent Types',
    width: 220,
    kind: 'list',
    note: 'Consent types that must be granted for the tag to fire, ' +
        'separated by commas: ' + CONSENT_TYPES.join(', ') + '. Setting ' +
        'consent types sets "Additional Consent Checks" to "Require ' +
        'additional consent".',
    appliesTo: (entity, data, ctx) => !ctx.isServer,
    get: entity => {
      const settings = entity.consentSettings;
      if (!settings || settings.consentStatus !== 'needed' ||
          !settings.consentType) {
        return '';
      }
      return formatList((settings.consentType.list || [])
          .map(item => item.value));
    },
    set: (draft, value) => {
      const types = parseList(value);
      draft.consentSettings = draft.consentSettings || {};
      if (types.length) {
        draft.consentSettings.consentStatus = 'needed';
        draft.consentSettings.consentType = {
          type: 'list',
          list: types.map(type => ({type: 'template', value: type}))
        };
      } else {
        delete draft.consentSettings.consentType;
      }
    }
  }, options);
}

/**
 * Validates the consent settings of a tag.
 * @param {!Object} draft
 * @return {!Array<string>} Errors.
 */
function validateConsentSettings_(draft) {
  const settings = draft.consentSettings;
  if (settings && settings.consentStatus === 'needed' &&
      !(settings.consentType && (settings.consentType.list || []).length)) {
    return ['"Require additional consent" needs at least one consent type ' +
        'in "Required Consent Types".'];
  }
  return [];
}

/**
 * Column: tag firing option.
 * @param {!Object=} options
 * @return {!Object}
 */
function firingOptionColumn_(options) {
  const labels = TAG_FIRING_OPTION_LABELS;
  return buildColumn_({
    key: 'tagFiringOption',
    header: 'Tag Firing Option',
    width: 140,
    dropdown: Object.keys(labels).map(key => labels[key]),
    strictDropdown: true,
    note: 'Once per event / Once per page / Unlimited.',
    appliesTo: (entity, data, ctx) => !ctx.isServer,
    get: entity => labels[entity.tagFiringOption] || '',
    set: (draft, value) => {
      const text = toText(value).toLowerCase();
      if (!text) {
        delete draft.tagFiringOption;
        return;
      }
      const option = Object.keys(labels).find(key =>
        labels[key].toLowerCase() === text || key.toLowerCase() === text);
      if (!option) {
        throw new UserError(`Unknown tag firing option "${toText(value)}".`);
      }
      draft.tagFiringOption = option;
    }
  }, options);
}

/**
 * Column: tag firing priority.
 * @param {!Object=} options
 * @return {!Object}
 */
function priorityColumn_(options) {
  return buildColumn_({
    key: 'priority',
    header: 'Tag Firing Priority',
    width: 110,
    note: 'Whole number; tags with a higher priority fire first. Leave empty ' +
        'for the default (0).',
    get: entity => entity.priority && entity.priority.value !== undefined ?
        String(entity.priority.value) : '',
    set: (draft, value) => {
      const text = toText(value);
      if (!text) {
        delete draft.priority;
      } else if (/^-?\d+$/.test(text)) {
        draft.priority = {type: 'integer', value: text};
      } else if (/^\{\{.+\}\}$/.test(text)) {
        draft.priority = {type: 'template', value: text};
      } else {
        throw new UserError('Tag Firing Priority must be a whole number.');
      }
    }
  }, options);
}

/**
 * Column: setup or cleanup (teardown) tag of a tag.
 * @param {string} field 'setupTag' or 'teardownTag'.
 * @param {!Object=} options
 * @return {!Object}
 */
function sequenceTagColumn_(field, options) {
  const isSetup = field === 'setupTag';
  const stopField = isSetup ? 'stopOnSetupFailure' : 'stopTeardownOnFailure';
  return buildColumn_({
    key: field,
    header: isSetup ? 'Setup Tag (fires before)' :
        'Cleanup Tag (fires after)',
    width: 200,
    dropdown: 'Tags',
    note: isSetup ? 'Name of a tag that fires before this tag.' :
        'Name of a tag that fires after this tag.',
    get: entity => entity[field] && entity[field][0] ?
        entity[field][0].tagName || '' : '',
    set: (draft, value, data) => {
      const name = toText(value);
      if (!name) {
        delete draft[field];
        return;
      }
      if (!data.getByName('tag', name)) {
        throw new UserError(`Tag "${name}" was not found.`);
      }
      const current = draft[field] && draft[field][0] ? draft[field][0] : {};
      const item = {tagName: name};
      item[stopField] = Boolean(current[stopField]);
      draft[field] = [item];
    }
  }, options);
}

/**
 * Column: whether the tag sequence stops when the setup / tag fails.
 * @param {string} field 'setupTag' or 'teardownTag'.
 * @param {!Object=} options
 * @return {!Object}
 */
function sequenceStopColumn_(field, options) {
  const isSetup = field === 'setupTag';
  const stopField = isSetup ? 'stopOnSetupFailure' : 'stopTeardownOnFailure';
  return buildColumn_({
    key: field + ':' + stopField,
    header: isSetup ? "Don't Fire if Setup Tag Fails" :
        "Don't Fire Cleanup Tag if This Tag Fails",
    width: 120,
    kind: 'checkbox',
    get: entity => Boolean(entity[field] && entity[field][0] &&
        entity[field][0][stopField]),
    set: (draft, value) => {
      if (!draft[field] || !draft[field][0]) {
        if (toBool(value)) {
          throw new UserError(`Choose a ${isSetup ? 'setup' : 'cleanup'} ` +
              'tag first.');
        }
        return;
      }
      draft[field][0][stopField] = toBool(value);
    }
  }, options);
}

/**
 * Column: start or end of the tag firing schedule.
 * @param {string} field 'scheduleStartMs' or 'scheduleEndMs'.
 * @param {!Object=} options
 * @return {!Object}
 */
function scheduleColumn_(field, options) {
  const isStart = field === 'scheduleStartMs';
  return buildColumn_({
    key: field,
    header: isStart ? 'Schedule Start' : 'Schedule End',
    width: 130,
    note: 'Date and time in the format yyyy-MM-dd HH:mm, in the time zone ' +
        'of this spreadsheet. Leave empty for no schedule.',
    appliesTo: (entity, data, ctx) => !ctx.isServer,
    get: (entity, data, ctx) => entity[field] ?
        Utilities.formatDate(new Date(Number(entity[field])), ctx.timeZone,
            'yyyy-MM-dd HH:mm') : '',
    set: (draft, value, data, ctx) => {
      const text = toText(value);
      if (!text) {
        delete draft[field];
        return;
      }
      const match = /^(\d{4}-\d{2}-\d{2})(?:[ T](\d{1,2}:\d{2}))?$/.exec(text);
      if (!match) {
        throw new UserError(`${isStart ? 'Schedule Start' : 'Schedule End'} ` +
            'must be in the format yyyy-MM-dd HH:mm.');
      }
      const date = Utilities.parseDate(`${match[1]} ${match[2] || '00:00'}`,
          ctx.timeZone, 'yyyy-MM-dd HH:mm');
      draft[field] = String(date.getTime());
    }
  }, options);
}

/**
 * Column: whether the tag only fires in published containers.
 * @param {!Object=} options
 * @return {!Object}
 */
function liveOnlyColumn_(options) {
  return buildColumn_({
    key: 'liveOnly',
    header: 'Only Fire in Published Containers',
    width: 120,
    kind: 'checkbox',
    get: entity => Boolean(entity.liveOnly),
    set: (draft, value) => {
      draft.liveOnly = toBool(value);
    }
  }, options);
}
