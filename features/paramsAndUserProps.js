/**
 * Copyright 2022 Google LLC
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
 * Writes the existing parameters and user properties to the sheet.
 */
function writeExistingParamsAndUserPropsToSheet() {
  clearRangeContent('modifyParamsAndUserProps', 'settings');
  writeGTMVariablesToValidationSheet();
  writeToSheet(
    listParamsAndUserProperties(), 'modifyParamsAndUserProps', 'settings');
}

/**
 * Writes the existing GA4 event tags to the parameters and user properties 
 * sheet.
 */
function writeGA4TagsToParamsAndUserPropertiesSheet() {
  const tags = listGTMResources('tags', getSelectedWorkspacePath());
  // Retrieve the GA4 event tags.
  const ga4EventTags = listGA4EventTags(tags);
  // Add specific event tag information to a double array to be written to the 
  // sheet.
  const formattedGA4EventTags = ga4EventTags.reduce((arr, tag) => {
    const eventName = tag.parameter.find(
      param => param.key == 'eventName').value;
    arr.push([tag.name, tag.tagId, eventName, '', '', '']);
    return arr;
  }, []);
  // Deletes existing content in the sheet.
  clearRangeContent('modifyParamsAndUserProps', 'settings');
  // Writes data to the sheet.
  writeToSheet(formattedGA4EventTags, 'modifyParamsAndUserProps', 'settings');
}

/**
 * Creates a list of all parameters and user properties that exist
 * in all of the GA4 event properties in a Tag Manager workspace.
 * @return {!Array<!Array<string, number, string, string, string>>}
 */
function listParamsAndUserProperties() {
  const tags = listGTMResources('tags', getSelectedWorkspacePath());
  writeGTMVariablesToValidationSheet();
  const ga4EventTags = listGA4EventTags(tags);
  return ga4EventTags.reduce((arr, eventTag) => {
    const eventName = eventTag.parameter.find(
      param => param.key == 'eventName').value;
    eventTag.parameter.forEach(param => {
      if (param.type == 'list') {
        let mapType = '';
        if (param.key == 'eventSettingsTable') {
          mapType = 'parameter';
        } else if (param.key == 'userProperties') {
          mapType = 'user_property';
        }
        param.list.forEach(map => {
          if (map.map[0].key == 'name' || map.map[0].key == 'parameter') {
            arr.push([
              eventTag.name,
              eventTag.tagId,
              eventName,
              mapType,
              map.map[0].value,
              map.map[1].value
            ]);
          } else {
            arr.push([
              eventTag.name,
              eventTag.tagId,
              eventName,
              mapType,
              map.map[1].value,
              map.map[0].value
            ]); 
          }
        });
      }
    });
    return arr;
  }, []);
}

/** 
 * Either adds or removes a user property or parameter from an event tag before
 * sending an update request to change the tag in Tag Manager.
 */
function modifyParametersAndUserProperties() {
  const tags = listGTMResources('tags', getSelectedWorkspacePath());
  const ga4EventTags = listGA4EventTags(tags);
  const newSettings = organizeParamsAndUserPropsByTag();
  ga4EventTags.forEach(tag => {
    if (newSettings[tag.tagId]) {
      const paramAndUserPropertySettings = newSettings[tag.tagId];
      let parametersAdded = false;
      let userPropertiesAdded = false;
      tag.parameter.forEach(param => {
        if (param.type == 'list') {
          if (param.key == 'eventSettingsTable') {
            param.list = removeParams(
              param.list, paramAndUserPropertySettings.remove.parameter);
            param.list = param.list.concat(
              paramAndUserPropertySettings.create.parameter);
            parametersAdded = true;
          } else if (param.key == 'userProperties') {
            param.list = removeParams(
              param.list, paramAndUserPropertySettings.remove.user_property)
            param.list = param.list.concat(
              paramAndUserPropertySettings.create.user_property);
            userPropertiesAdded = true;
          }
        }
      });
      if (!parametersAdded && 
      paramAndUserPropertySettings.create.parameter.length > 0) {
        tag.parameter.push({
          type: 'list',
          key: 'eventSettingsTable',
          list: paramAndUserPropertySettings.create.parameter
        });
      }
      if (!userPropertiesAdded && 
      paramAndUserPropertySettings.create.user_property.length > 0) {
        tag.parameter.push({
          type: 'list',
          key: 'userProperties',
          list: paramAndUserPropertySettings.create.user_property
        });
      }
      const response = updateGTMResource(
        'tags', getSelectedWorkspacePath() + '/tags/' + tag.tagId, tag);
      const rows = paramAndUserPropertySettings.rows;
      if (response.tagManagerUrl) {
        const actionTaken = generateSuccessfullActionTakenText(
          newSettings[tag.tagId]);
        logChange(response.name, response.type, response.tagId, actionTaken, 
          response.tagManagerUrl);
        rows.forEach(r => {
          writeActionStatus('modifyParamsAndUserProps', r.rowNumber, 8, 'Success');
        });
      } else {
        const errMsg = errorResponse(response);
        logChange(tag.name, tag.type, tag.tagId, errMsg, tag.tagManagerUrl);
        rows.forEach(r => {
          writeActionStatus('modifyParamsAndUserProps', r.rowNumber, 8, errMsg);
        });
      }
    }
  });
}

/**
 * Organizes the user property and parameter settings from the spreadsheet
 * by tag insteady of by row so that the settings can be more easily applied to
 * a given tag.
 * @return {!Array<!Object>}
 */
function organizeParamsAndUserPropsByTag() {
  const data = getDataWithRowNumbers('modifyParamsAndUserProps', 'settings');
  return data.reduce((tagObj, rowInfo) => {
    const row = rowInfo.values;
    const rowNumber = rowInfo.rowNumber;
    const tagId = row[1];
    const type = row[3];
    const name = row[4];
    const value = row[5];
    const action = row[6];
    if (/Create|Delete/.test(action)) {
      if (tagObj[tagId] == undefined) {
        tagObj[tagId] = {
          create: {
            parameter: [],
            user_property: []
          },
          remove: {
            parameter: [],
            user_property: []
          },
          rows: []
        };
      }
      tagObj[tagId].rows.push({ rowNumber, action, type, name });
      if (action == 'Create') {
        if (type == 'parameter') {
          tagObj[tagId].create[type].push(buildParameterMapObject(name, value));
        } else if (type == 'user_property') {
          tagObj[tagId].create[type].push(buildUserPropertyMapObject(name, value));
        }
      } else if (action == 'Delete') {
        tagObj[tagId].remove[type].push(
          {name: name.toString(), value:value.toString()});
      }
    }
    return tagObj;
  }, []);
}

/**
 * Removes parameters or user properties from the params list in a tag.
 * @param {!Array} tagParamList A list of parameters that already exist in a 
 * Tag Manager tag.
 * @param {!Array} paramsToDelete A list of parameters that need to be removed 
 * from the Tag Manager list.
 * @return {!Array}
 */
function removeParams(tagParamList, paramsToDelete) {
  const newList = JSON.parse(JSON.stringify(tagParamList));
  if (paramsToDelete.length > 0) {
    tagParamList.forEach(param => {
      for (let i = 0; i < paramsToDelete.length; i++) {
        const originalParamName = param.map[0].value;
        const originalParamValue = param.map[1].value;
        const paramToDeleteName = paramsToDelete[i].name;
        const paramToDeleteValue = paramsToDelete[i].value;
        if (originalParamName == paramToDeleteName && 
        originalParamValue == paramToDeleteValue) {
          for (let newListIndex = 0; newListIndex < newList.length; 
          newListIndex++) {
            if(newList[newListIndex].map[0].value == paramToDeleteName && 
            newList[newListIndex].map[1].value == paramToDeleteValue) {
              newList.splice(newListIndex, 1);
              break;
            }
          }
          break;
        }
      }
    });
    return newList;
  } else {
    return tagParamList;
  }
}

/**
 * Create the change log text for the changes made to a given tag.
 * @param {!Array} paramAndUserPropertySettings A list of changes that were made
 * for a given tag.
 * @return {string}
 */
function generateSuccessfullActionTakenText(paramAndUserPropertySettings) {
  let changes = 'The following has been modified in this tag:';
  if (paramAndUserPropertySettings.create.parameter.length > 0) {
    changes += `
    - Parameters created:`
    paramAndUserPropertySettings.create.parameter.forEach(param => {
      changes += `
        - Name: ${param.map[0].value}, Value: ${param.map[1].value}`
    });
  }
  if (paramAndUserPropertySettings.create.user_property.length > 0) {
    changes += `
    - User Properties created:`
    paramAndUserPropertySettings.create.user_property.forEach(prop => {
      changes += `
        - Name: ${prop.map[0].value}, Value: ${prop.map[1].value}`
    });
  }
  if (paramAndUserPropertySettings.remove.parameter.length > 0) {
    changes += `
    - Parameters removed:`
    paramAndUserPropertySettings.remove.parameter.forEach(param => {
      changes += `
        - Name: ${param.name}, Value: ${param.value}`
    });
  }
  if (paramAndUserPropertySettings.remove.user_property.length > 0) {
    changes += `
    - User Properties removed:`
    paramAndUserPropertySettings.remove.user_property.forEach(prop => {
      changes += `
        - Name: ${prop.name}, Value: ${prop.value}`
    });
  }
  return changes;
}