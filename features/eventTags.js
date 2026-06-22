/**
 * Copyright 2023 Google LLC
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
 * Writes the existing GA4 event tags to the sheet.
 */
function writeGA4EventTagstoSheet() {
  const workspacePath = getSelectedWorkspacePath();
  const workspaceName = getSelectedWorkspaceName();
  const tags = listGTMResources('tags', workspacePath);
  const variables = listGTMResources('variables', workspacePath);
  const triggers = listGTMResources('triggers', workspacePath);
  const folders = listGTMResources('folders', workspacePath);
  writeGTMVariablesToValidationSheet();
  triggers.push({
    triggerId: '2147479553',
    name: 'All Pages'
  }, {
    triggerId: '2147479572',
    name: 'Consent Initialization - All Pages'
  }, {
    triggerId: '2147479573',
    name: 'Initialization - All Pages'
  });
  // Retrieve the GA4 event tags.
  const ga4EventTags = listGA4EventTags(tags);
  clearRangeContent('validation', 'event setting variable names');
  writeSettingVariableNamesToValidationSheet(variables);
  clearRangeContent('validation', 'tag names');
  writeTagNamesToValidationSheet(tags);
  // Add specific event tag information to a double array to be written to the 
  // sheet.
  const formattedGA4EventTags = ga4EventTags.reduce((arr, tag) => {
    const tempArray = [];
    while (tempArray.length < 24) {
      tempArray.push('');
    }
    const sendEcommerceData = tag.parameter.find(
      param => param.key == 'sendEcommerceData');
    const getEcommerceDataFrom = tag.parameter.find(
      param => param.key == 'getEcommerceDataFrom');

    tempArray[0] = workspaceName;
    tempArray[1] = workspacePath;
    tempArray[2] = tag.name;
    tempArray[3] = tag.tagId;
    tempArray[4] = tag.paused;
    const eventNameParam = tag.parameter.find(param => param.key == 'eventName');
    tempArray[5] = eventNameParam ? eventNameParam.value : '';

    if (sendEcommerceData != undefined) {
      tempArray[6] = sendEcommerceData.value;
    }
    if (getEcommerceDataFrom != undefined) {
      if (getEcommerceDataFrom.value == 'dataLayer') {
        tempArray[7] = 'dataLayer';
      } else {
        const ecommerceMacroData = tag.parameter.find(
          param => param.key == 'ecommerceMacroData');
        tempArray[7] = ecommerceMacroData ? ecommerceMacroData.value : '';
      }
    }
    const measurementIdOverride = tag.parameter.find(
      param => param.key == 'measurementIdOverride');
    tempArray[8] = measurementIdOverride ? measurementIdOverride.value : '';
    const eventSettingsVariable = tag.parameter.find(
      param => param.key == 'eventSettingsVariable');
    tempArray[9] = '';
    if (eventSettingsVariable) {
      tempArray[9] = eventSettingsVariable.value;
    } 
    if (tag.priority != undefined) {
      tempArray[10] = tag.priority.value;
    }
    tempArray[11] = tag.scheduleStartMs;
    tempArray[12] = tag.scheduleEndMs;
    tempArray[13] = tag.liveOnly;
    if (tag.setupTag != undefined) {
      tempArray[14] = tag.setupTag[0].tagName;
      tempArray[15] = tag.setupTag[0].stopOnSetupFailure;
    }
    if (tag.teardownTag != undefined) {
      tempArray[16] = tag.teardownTag[0].tagName;
      tempArray[17] = tag.teardownTag[0].stopTeardownOnFailure;
    }
    tempArray[18] = tag.tagFiringOption;
    tempArray[19] = tag.consentSettings.consentStatus;
    if (tag.consentSettings.consentStatus == 'needed') {
      tempArray[20] = tag.consentSettings.consentType.list
      .reduce((arr, consentType) => {
        arr.push(consentType.value); 
        return arr;
      }, []).join(', ');
    }
    tempArray[21] = tag.monitoringMetadataTagNameKey;
    if (tag.firingTriggerId != undefined) {
      tempArray[22] = convertTriggerIdsToNames(
        tag.firingTriggerId, triggers).join(', ');
    }
    if (tag.blockingTriggerId != undefined) {
      tempArray[23] = convertTriggerIdsToNames(
        tag.blockingTriggerId, triggers).join(', ');
    } 
    tempArray[24] = convertFolderIdsToNames(tag.parentFolderId, folders);
    tempArray[25] = tag.notes;
    arr.push(tempArray);
    return arr;
  }, []);
  // Deletes existing content in the sheet.
  clearRangeContent('eventTags', 'tags');
  // Writes data to the sheet.
  writeToSheet(formattedGA4EventTags, 'eventTags', 'tags');
}

/**
 * Either creates, updates, or removes existing GA4 event tags from a workspace.
 */
function modifyGA4EventTags() {
  const workspacePath = getSelectedWorkspacePath();
  const tagSettingsData = getDataWithRowNumbers('eventTags', 'tags');
  const tags = listGTMResources('tags', workspacePath);
  const triggers = listGTMResources('triggers', workspacePath);
  const folders = listGTMResources('folders', workspacePath);
  tagSettingsData.forEach(rowInfo => {
    const newSettings = rowInfo.values;
    const rowNumber = rowInfo.rowNumber;
    const create = newSettings[26];
    const update = newSettings[27];
    const remove = newSettings[28];
    if (create || update || remove) {
      if (create && !update && !remove) {
        // Create a new tag.
        const payload = buildGA4EventPayload(
          newSettings, triggers, tags, folders, 'create');
        const response = createGTMResource('tags', newSettings[1], payload);
        if (response.tagManagerUrl) {
          // Write succesful tag creation to sheet.
          logChange(response.name, response.type, response.tagId, 'created',
            response.tagManagerUrl);
          writeActionStatus('eventTags', rowNumber, 30, 'Created successfully');
        } else {
          // Write error message to sheet.
          const errMsg = errorResponse(response);
          logChange(newSettings[2], '',  '', errMsg, '');
          writeActionStatus('eventTags', rowNumber, 30, errMsg);
        }  
      } else if (!create && update && !remove) {
        // Update an existing tag.
        const payload = buildGA4EventPayload(
          newSettings, triggers, tags, folders, 'update');
        const response = updateGTMResource(
          'tags', `${newSettings[1]}/tags/${newSettings[3]}`, payload);
        if (response.tagManagerUrl) {
          // Write a successful tag update to sheet.
          logChange(response.name, response.type, response.tagId, 'updated',
            response.tagManagerUrl);
          writeActionStatus('eventTags', rowNumber, 30, 'Updated successfully');
        } else {
          // Write error message to sheet.
          const errMsg = errorResponse(response);
          logChange(newSettings[2], '',  '', errMsg, '');
          writeActionStatus('eventTags', rowNumber, 30, errMsg);
        }
      } else if (!create && !update && remove) {
        // Remove an existing tag.
        const response = removeGTMResource(
          'tags', `${newSettings[1]}/tags/${newSettings[3]}`);
        if (response.details) {
          // Write error message to sheet.
          const errMsg = errorResponse(response);
          logChange(newSettings[2], '',  '', errMsg, '');
          writeActionStatus('eventTags', rowNumber, 30, errMsg);
        } else {
          // Write successful tag deletion to sheet.
          logChange(newSettings[2], 'gaawe', newSettings[3], 'removed', '');
          writeActionStatus('eventTags', rowNumber, 30, 'Deleted successfully');
        }
      } else {
        writeActionStatus('eventTags', rowNumber, 30, 'Error: multiple actions selected');
      }
    }
  });
}

/**
 * Builds the GA4 event payload when either creating a new tag or updating an 
 * existing event tag.
 * @param {!Array} settings The settings from the sheet.
 * @param {!Array} triggers Existing triggers in the workspace.
 * @param {!Array} tags Existing tags in the workspace
 * @param {!Array} folders Existing folders in the workspace.
 * @param {string} modifyOption Either "update" or "create"
 * @return {!Object} The new tag payload object.
 */
function buildGA4EventPayload(settings, triggers, tags, folders, modifyOption) {
  const tagName = settings[2];
  const tagId = settings[3];
  const paused = Boolean(settings[4]);
  const eventName = settings[5];
  const sendEcommData = Boolean(settings[6]);
  const ecommObject = settings[7];
  const measurementIdOverride = settings[8].trim();
  const eventSettingsVariable = settings[9];
  const tagFiringPriority = settings[10];
  const scheduleStart = settings[11];
  const scheduleEnd = settings[12];
  const onlyFireInPublished = Boolean(settings[13]);
  const setupTag = settings[14];
  const stopOnSetupFailure = Boolean(settings[15]);
  const teardownTag = settings[16];
  const stopTeardownOnFailure = Boolean(settings[17]);
  const tagFiringOption = settings[18];
  const consentStatus = settings[19];
  const consentStatusType = settings[20];
  const metadataTagName = settings[21];
  const firingTriggerNames = settings[22];
  const blockingTriggerNames = settings[23];
  const parentFolderName = settings[24];
  const notes = settings[25];
  const firingTriggerIds = convertTriggerNamesToIds(
    firingTriggerNames, triggers);
  const blockingTriggerIds = convertTriggerNamesToIds(
    blockingTriggerNames, triggers);
  const parentFolderId = convertFolderNamesToIds(
    parentFolderName, folders);

  const newTag = {
    type: 'gaawe',
    name: tagName,
    paused: paused,
    liveOnly: onlyFireInPublished,
    tagFiringOption: tagFiringOption,
    consentSettings: {
      consentStatus: consentStatus
    },
    monitoringMetadataTagNameKey: metadataTagName,
    firingTriggerId: firingTriggerIds,
    blockingTriggerId: blockingTriggerIds,
    parentFolderId: parentFolderId,
    notes: notes,
    parameter: [{
      type: 'template',
      key: 'eventName',
      value: eventName
    }, {
      type: 'template',
      key: 'measurementIdOverride',
      value: measurementIdOverride
    }]
  }
  if (eventSettingsVariable.length > 0) {
    newTag.parameter.push({
      "type": "template",
      "key": "eventSettingsVariable",
      "value": eventSettingsVariable
    });
  }
  if (setupTag.length > 0) {
    newTag.setupTag = [{
      tagName: setupTag,
      stopOnSetupFailure: stopOnSetupFailure
    }];
  }
  if (teardownTag.length > 0) {
    newTag.teardownTag = [{
      tagName: teardownTag,
      stopTeardownOnFailure: stopTeardownOnFailure
    }];
  }
  if (/^[0-9]+$/.test(tagFiringPriority)) {
    newTag.priority = {
      type: 'integer',
      value: parseInt(tagFiringPriority)
    }
  }
  if (/^[0-9]+$/.test(scheduleStart)) {
    newTag.scheduleStartMs = parseInt(scheduleStart);
  }
  if (/^[0-9]+$/.test(scheduleEnd)) {
    newTag.scheduleEndMs = parseInt(scheduleEnd);
  }
  if (sendEcommData) {
    newTag.parameter.push({
      type: 'boolean',
      key: 'sendEcommerceData',
      value: 'true'
    });
    if (ecommObject == 'dataLayer') {
      newTag.parameter.push({
        type: 'template',
        key: 'getEcommerceDataFrom',
        value: ecommObject
      })
    } else {
      newTag.parameter.push({
        type: 'template',
        key: 'ecommerceMacroData',
        value: ecommObject
      }, {
        type: 'template',
        key: 'getEcommerceDataFrom',
        value: 'customObject'
      })
    }
  }
  if (consentStatus == 'needed') {
    newTag.consentSettings.consentType = {
      type: 'list',
      list: consentStatusType.split(',').reduce((arr, consentType) => {
        arr.push({
          type: 'template',
          value: consentType.trim()
        });
        return arr
      }, [])
    }
  }
  if (modifyOption == 'update') {
    const oldTag = tags.find(tag => tag.tagId == tagId);
    if (oldTag.parameter.length > 0) {
      oldTag.parameter.forEach(param => {
        if (!/sendEcommerceData|getEcommerceDataFrom|ecommerceMacroData|measurementId|eventName/.test(param.key)) {
          newTag.parameter.push(param);
        }
      });
    }
  }
  return newTag;
}

/**
 * Converts an array of trigger IDs to an array of trigger names.
 * @param {!Array<number>} ids An array of trigger IDs.
 * @param {!Array<!Object>} triggers An array of triggers.
 * @return {!Array<string>} An array of trigger names.
 */
function convertTriggerIdsToNames(ids, triggers) {
  const names = [];
  triggers.push({
    triggerId: '2147479553',
    name: 'All Pages'
  }, {
    triggerId: '2147479572',
    name: 'Consent Initialization - All Pages'
  }, {
    triggerId: '2147479573',
    name: 'Initialization - All Pages'
  });
  if (ids && ids.length > 0) {
    ids.forEach(id => {
      const idStr = String(id).trim();
      const foundTrigger = triggers.find(trigger => String(trigger.triggerId) == idStr);
      if (foundTrigger) {
        names.push(foundTrigger.name);
      } else {
        names.push('Unknown Trigger (' + idStr + ')');
      }
    });
  }
  return names;
}

/**
 * Converts a folder ID to a folder names.
 * @param {string} id A folder ID.
 * @param {!Array<!Object>} folders An array of folders.
 * @return {string} The folder name
 */
function convertFolderIdsToNames(id, folders) {
  if (id) {
    return folders.find(folder => folder.folderId == id).name;
  } else {
    return '';
  }
}

/**
 * Converts an array of trigger names to an array of trigger ids.
 * @param {!Array<string>} ids An array of trigger names.
 * @param {!Array<!Object>} triggers An array of triggers.
 * @return {!Array<number>} An array of trigger IDs.
 */
function convertTriggerNamesToIds(names, triggers) {
  const ids = []
  triggers.push({
    triggerId: '2147479553',
    name: 'All Pages'
  }, {
    triggerId: '2147479572',
    name: 'Consent Initialization - All Pages'
  }, {
    triggerId: '2147479573',
    name: 'Initialization - All Pages'
  });
  const namesStr = String(names || '').trim();
  if (namesStr.length > 0) {
    const nameList = namesStr.split(',');
    nameList.forEach(name => {
      const foundTrigger = triggers.find(trigger => trigger.name == name.trim());
      if (foundTrigger) {
        ids.push(foundTrigger.triggerId.toString());
      } else {
        console.log('Warning: Trigger not found: ' + name);
      }
    });
  }
  return ids;
}

/**
 * Converts a folder name to a folder parent ID.
 * @param {string} name A folder name.
 * @param {!Array<!Object>} folders An array of folders.
 * @return {string} The folder name
 */
function convertFolderNamesToIds(name, folders) {
  if (name) {
    return folders.find(folder => folder.name == name).folderId;
  } else {
    return '';
  }
}