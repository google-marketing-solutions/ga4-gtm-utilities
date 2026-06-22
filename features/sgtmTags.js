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
 * Writes the existing sGTM tags to the sheet.
 */
function writeSgtmTagsToSheet() {
  const workspacePath = getSelectedWorkspacePath();
  const workspaceName = getSelectedWorkspaceName();
  const tags = listGTMResources('tags', workspacePath);
  const triggers = listGTMResources('triggers', workspacePath);
  const folders = listGTMResources('folders', workspacePath);

  if (!tags) {
    clearRangeContent('sgtmTags', 'tags');
    return;
  }

  const formattedTags = tags.reduce((arr, tag) => {
    const tempArray = [];
    while (tempArray.length < 14) {
      tempArray.push('');
    }

    const params = getTagParams(tag);

    tempArray[0] = workspaceName;
    tempArray[1] = workspacePath;
    tempArray[2] = tag.name;
    tempArray[3] = tag.tagId;
    tempArray[4] = tag.type;
    tempArray[5] = tag.paused || false;
    tempArray[6] = params['measurementId'] || '';
    tempArray[7] = params['parametersToExclude'] || '';
    tempArray[8] = params['parametersToOverride'] || params['parameter'] || '';
    tempArray[9] = params['serverContainerUrl'] || '';

    if (tag.firingTriggerId !== undefined) {
      tempArray[10] = convertTriggerIdsToNames(tag.firingTriggerId, triggers).join(', ');
    }
    if (tag.blockingTriggerId !== undefined) {
      tempArray[11] = convertTriggerIdsToNames(tag.blockingTriggerId, triggers).join(', ');
    }
    tempArray[12] = convertFolderIdsToNames(tag.parentFolderId, folders);
    tempArray[13] = tag.notes || '';

    arr.push(tempArray);
    return arr;
  }, []);

  clearRangeContent('sgtmTags', 'tags');
  writeToSheet(formattedTags, 'sgtmTags', 'tags');
}

/**
 * Helper to parse tag parameters into a flat key-value dictionary.
 * @param {!Object} tag Tag resource object.
 * @return {!Object} Key-value pairs dictionary of parameters.
 */
function getTagParams(tag) {
  const params = {};
  if (tag.parameter) {
    tag.parameter.forEach(p => {
      if (p.type === 'template' || p.type === 'boolean' || p.type === 'integer') {
        params[p.key] = p.value;
      } else if (p.type === 'list' && p.list) {
        if (p.list.every(item => item.type === 'template')) {
          params[p.key] = p.list.map(item => item.value).join(', ');
        } else {
          params[p.key] = JSON.stringify(p.list);
        }
      } else if (p.type === 'map' && p.map) {
        params[p.key] = JSON.stringify(p.map);
      }
    });
  }
  return params;
}

/**
 * Either creates, updates, or removes existing sGTM tags from a workspace.
 */
function modifySgtmTags() {
  const workspacePath = getSelectedWorkspacePath();
  const tagSettingsData = getDataWithRowNumbers('sgtmTags', 'tags');
  const tags = listGTMResources('tags', workspacePath);
  const triggers = listGTMResources('triggers', workspacePath);
  const folders = listGTMResources('folders', workspacePath);

  tagSettingsData.forEach(rowInfo => {
    const newSettings = rowInfo.values;
    const rowNumber = rowInfo.rowNumber;
    const create = newSettings[14];
    const update = newSettings[15];
    const remove = newSettings[16];

    if (create || update || remove) {
      if (create && !update && !remove) {
        const payload = buildSgtmTagPayload(newSettings, triggers, tags, folders, 'create');
        const response = createGTMResource('tags', newSettings[1], payload);
        if (response && response.tagManagerUrl) {
          logChange(response.name, response.type, response.tagId, 'created', response.tagManagerUrl);
          writeActionStatus('sgtmTags', rowNumber, 18, 'Created successfully');
        } else {
          const errMsg = errorResponse(response);
          logChange(newSettings[2], '', '', errMsg, '');
          writeActionStatus('sgtmTags', rowNumber, 18, errMsg);
        }
      } else if (!create && update && !remove) {
        const payload = buildSgtmTagPayload(newSettings, triggers, tags, folders, 'update');
        const response = updateGTMResource('tags', `${newSettings[1]}/tags/${newSettings[3]}`, payload);
        if (response && response.tagManagerUrl) {
          logChange(response.name, response.type, response.tagId, 'updated', response.tagManagerUrl);
          writeActionStatus('sgtmTags', rowNumber, 18, 'Updated successfully');
        } else {
          const errMsg = errorResponse(response);
          logChange(newSettings[2], '', '', errMsg, '');
          writeActionStatus('sgtmTags', rowNumber, 18, errMsg);
        }
      } else if (!create && !update && remove) {
        const response = removeGTMResource('tags', `${newSettings[1]}/tags/${newSettings[3]}`);
        if (response && response.details) {
          const errMsg = errorResponse(response);
          logChange(newSettings[2], '', '', errMsg, '');
          writeActionStatus('sgtmTags', rowNumber, 18, errMsg);
        } else {
          logChange(newSettings[2], 'tag', newSettings[3], 'removed', '');
          writeActionStatus('sgtmTags', rowNumber, 18, 'Deleted successfully');
        }
      } else {
        writeActionStatus('sgtmTags', rowNumber, 18, 'Error: multiple actions selected');
      }
    }
  });
}

/**
 * Builds the GTM Server Tag payload.
 * @param {!Array} settings Tag settings from sheet.
 * @param {!Array} triggers Triggers list.
 * @param {!Array} tags Tags list.
 * @param {!Array} folders Folders list.
 * @param {string} modifyOption Modify action (create or update).
 * @return {!Object} Tag payload.
 */
function buildSgtmTagPayload(settings, triggers, tags, folders, modifyOption) {
  const tagName = settings[2];
  const tagType = settings[4];
  const paused = Boolean(settings[5]);
  const measurementId = settings[6].trim();
  const parametersToExclude = settings[7].trim();
  const parametersToOverrideStr = settings[8].trim();
  const serverContainerUrl = settings[9].trim();
  const firingTriggerNames = settings[10];
  const blockingTriggerNames = settings[11];
  const parentFolderName = settings[12];
  const notes = settings[13];

  const firingTriggerIds = convertTriggerNamesToIds(firingTriggerNames, triggers);
  const blockingTriggerIds = convertTriggerNamesToIds(blockingTriggerNames, triggers);
  const parentFolderId = convertFolderNamesToIds(parentFolderName, folders);

  const newTag = {
    name: tagName,
    type: tagType,
    paused: paused,
    firingTriggerId: firingTriggerIds,
    blockingTriggerId: blockingTriggerIds,
    parentFolderId: parentFolderId,
    notes: notes,
    parameter: []
  };

  if (measurementId) {
    newTag.parameter.push({
      type: 'template',
      key: 'measurementId',
      value: measurementId
    });
  }

  if (serverContainerUrl) {
    newTag.parameter.push({
      type: 'template',
      key: 'serverContainerUrl',
      value: serverContainerUrl
    });
  }

  if (parametersToExclude) {
    newTag.parameter.push({
      type: 'list',
      key: 'parametersToExclude',
      list: parametersToExclude.split(',').map(param => ({
        type: 'template',
        value: param.trim()
      }))
    });
  }

  if (parametersToOverrideStr) {
    try {
      const parsedOverride = JSON.parse(parametersToOverrideStr);
      newTag.parameter.push({
        type: Array.isArray(parsedOverride) ? 'list' : 'map',
        key: 'parametersToOverride',
        ...(Array.isArray(parsedOverride) ? { list: parsedOverride } : { map: parsedOverride })
      });
    } catch (e) {
      const mapList = [];
      parametersToOverrideStr.split(';').forEach(pair => {
        const parts = pair.split(':');
        if (parts.length >= 2) {
          const key = parts[0].trim();
          const val = parts.slice(1).join(':').trim();
          mapList.push({
            type: 'map',
            map: [
              { type: 'template', key: 'parameter', value: key },
              { type: 'template', key: 'parameterValue', value: val }
            ]
          });
        }
      });
      if (mapList.length > 0) {
        newTag.parameter.push({
          type: 'list',
          key: 'parametersToOverride',
          list: mapList
        });
      }
    }
  }

  if (modifyOption === 'update') {
    const oldTag = tags.find(t => t.tagId === settings[3]);
    if (oldTag && oldTag.parameter) {
      oldTag.parameter.forEach(param => {
        const isOverridden = newTag.parameter.some(newParam => newParam.key === param.key);
        if (!isOverridden) {
          newTag.parameter.push(param);
        }
      });
    }
  }

  return newTag;
}
