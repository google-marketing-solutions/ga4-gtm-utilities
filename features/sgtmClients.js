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
 * Writes the existing sGTM clients to the sheet.
 */
function writeSgtmClientsToSheet() {
  const workspacePath = getSelectedWorkspacePath();
  const workspaceName = getSelectedWorkspaceName();
  const clients = listGTMResources('clients', workspacePath);
  const folders = listGTMResources('folders', workspacePath);
  
  if (!clients) {
    clearRangeContent('sgtmClients', 'clients');
    return;
  }

  // Format the clients data for the sheet
  const formattedClients = clients.reduce((arr, client) => {
    const tempArray = [];
    while (tempArray.length < 22) {
      tempArray.push('');
    }
    
    // Find parameters
    const params = getClientParams(client);
    
    tempArray[0] = workspaceName;
    tempArray[1] = workspacePath;
    tempArray[2] = client.type;
    tempArray[3] = client.clientId;
    tempArray[4] = convertFolderIdsToNames(client.parentFolderId, folders);
    tempArray[5] = client.notes || '';
    tempArray[6] = client.name;
    tempArray[7] = client.paused || false;
    tempArray[8] = client.priority !== undefined ? client.priority : '';
    
    // GA4 parameters
    tempArray[9] = params['activateDefaultPaths'] !== undefined ? params['activateDefaultPaths'] : '';
    tempArray[10] = params['cookieManagement'] || '';
    tempArray[11] = params['cookieName'] || '';
    tempArray[12] = params['cookieDomain'] || '';
    tempArray[13] = params['cookiePath'] || '';
    tempArray[14] = params['sameSite'] || '';
    tempArray[15] = params['cookieMaxAgeInSec'] || '';
    tempArray[16] = params['migrateFromJsClientId'] !== undefined ? params['migrateFromJsClientId'] : '';
    
    // GTM parameters
    tempArray[17] = params['allowedContainerIds'] || '';
    tempArray[18] = params['isPathMandatory'] !== undefined ? params['isPathMandatory'] : '';
    tempArray[19] = params['activateResponseCompression'] !== undefined ? params['activateResponseCompression'] : '';
    tempArray[20] = params['activateGeoResolution'] !== undefined ? params['activateGeoResolution'] : '';
    tempArray[21] = params['region'] || '';
    
    arr.push(tempArray);
    return arr;
  }, []);
  
  clearRangeContent('sgtmClients', 'clients');
  writeToSheet(formattedClients, 'sgtmClients', 'clients');
}

/**
 * Helper to parse client parameters into a flat key-value dictionary.
 * @param {!Object} client Client resource object.
 * @return {!Object} Key-value pairs dictionary of parameters.
 */
function getClientParams(client) {
  const params = {};
  if (client.parameter) {
    client.parameter.forEach(p => {
      if (p.type === 'template' || p.type === 'boolean' || p.type === 'integer') {
        params[p.key] = p.value;
      } else if (p.type === 'list' && p.list) {
        if (p.key === 'allowedContainerIds') {
          // Format as: GTM-11111:/path1; GTM-22222:/path2
          params[p.key] = p.list.map(item => {
            const containerIdParam = item.map.find(m => m.key === 'containerId');
            const pathParam = item.map.find(m => m.key === 'path');
            const cid = containerIdParam ? containerIdParam.value : '';
            const path = pathParam ? pathParam.value : '';
            return `${cid}:${path}`;
          }).join('; ');
        } else {
          params[p.key] = p.list.map(li => li.value).join('; ');
        }
      }
    });
  }
  return params;
}

/**
 * Either creates, updates, or removes existing sGTM clients from a workspace.
 */
function modifySgtmClients() {
  const workspacePath = getSelectedWorkspacePath();
  const clientSettingsData = getDataWithRowNumbers('sgtmClients', 'clients');
  const folders = listGTMResources('folders', workspacePath);
  const clients = listGTMResources('clients', workspacePath);

  clientSettingsData.forEach(rowInfo => {
    const newSettings = rowInfo.values;
    const rowNumber = rowInfo.rowNumber;
    const create = newSettings[22];
    const update = newSettings[23];
    const remove = newSettings[24];
    
    if (create || update || remove) {
      if (create && !update && !remove) {
        const payload = buildSgtmClientPayload(newSettings, folders, 'create');
        const response = createGTMResource('clients', newSettings[1], payload);
        if (response && response.tagManagerUrl) {
          logChange(response.name, 'client', response.clientId, 'created', response.tagManagerUrl);
          writeActionStatus('sgtmClients', rowNumber, 26, 'Created successfully');
        } else {
          const errMsg = errorResponse(response);
          logChange(newSettings[6], 'client', '', errMsg, '');
          writeActionStatus('sgtmClients', rowNumber, 26, errMsg);
        }
      } else if (!create && update && !remove) {
        const payload = buildSgtmClientPayload(newSettings, folders, 'update', clients);
        const response = updateGTMResource('clients', `${newSettings[1]}/clients/${newSettings[3]}`, payload);
        if (response && response.tagManagerUrl) {
          logChange(response.name, 'client', response.clientId, 'updated', response.tagManagerUrl);
          writeActionStatus('sgtmClients', rowNumber, 26, 'Updated successfully');
        } else {
          const errMsg = errorResponse(response);
          logChange(newSettings[6], 'client', '', errMsg, '');
          writeActionStatus('sgtmClients', rowNumber, 26, errMsg);
        }
      } else if (!create && !update && remove) {
        const response = removeGTMResource('clients', `${newSettings[1]}/clients/${newSettings[3]}`);
        if (response && response.details) {
          const errMsg = errorResponse(response);
          logChange(newSettings[6], 'client', '', errMsg, '');
          writeActionStatus('sgtmClients', rowNumber, 26, errMsg);
        } else {
          logChange(newSettings[6], 'client', newSettings[3], 'removed', '');
          writeActionStatus('sgtmClients', rowNumber, 26, 'Deleted successfully');
        }
      } else {
        writeActionStatus('sgtmClients', rowNumber, 26, 'Error: multiple actions selected');
      }
    }
  });
}

/**
 * Builds the GTM Client payload.
 * @param {!Array} settings Client settings from sheet.
 * @param {!Array} folders Folders array.
 * @param {string} modifyOption Modify action (create or update).
 * @param {!Array=} clients Existing clients in workspace (required for update).
 * @return {!Object} The Client payload structure.
 */
function buildSgtmClientPayload(settings, folders, modifyOption, clients) {
  const name = settings[6];
  const type = settings[2];
  const paused = Boolean(settings[7]);
  const priority = settings[8];

  // GA4 parameters
  const activateDefaultPaths = settings[9];
  const cookieManagement = settings[10].trim();
  const cookieName = settings[11].trim();
  const cookieDomain = settings[12].trim();
  const cookiePath = settings[13].trim();
  const sameSite = settings[14].trim();
  const cookieMaxAge = settings[15].toString().trim();
  const migrateFromJsClientId = settings[16];

  // GTM parameters
  const allowedContainersStr = settings[17].trim();
  const isPathMandatory = settings[18];
  const activateResponseCompression = settings[19];
  const activateGeoResolution = settings[20];
  const region = settings[21].trim();

  const parentFolderName = settings[4];
  const notes = settings[5];
  
  const parentFolderId = convertFolderNamesToIds(parentFolderName, folders);
  
  const newClient = {
    name: name,
    type: type,
    paused: paused,
    parentFolderId: parentFolderId,
    notes: notes,
    parameter: []
  };

  if (priority !== '') {
    newClient.priority = parseInt(priority);
  }

  if (type === 'gaaw_client') {
    if (cookieDomain) newClient.parameter.push({ type: 'template', key: 'cookieDomain', value: cookieDomain });
    if (cookieName) newClient.parameter.push({ type: 'template', key: 'cookieName', value: cookieName });
    if (cookiePath) newClient.parameter.push({ type: 'template', key: 'cookiePath', value: cookiePath });
    if (cookieMaxAge) newClient.parameter.push({ type: 'template', key: 'cookieMaxAgeInSec', value: cookieMaxAge });
    if (cookieManagement) newClient.parameter.push({ type: 'template', key: 'cookieManagement', value: cookieManagement });
    if (sameSite) newClient.parameter.push({ type: 'template', key: 'sameSite', value: sameSite });
    
    if (migrateFromJsClientId !== '') {
      newClient.parameter.push({ type: 'boolean', key: 'migrateFromJsClientId', value: String(migrateFromJsClientId) });
    }
    if (activateDefaultPaths !== '') {
      newClient.parameter.push({ type: 'boolean', key: 'activateDefaultPaths', value: String(activateDefaultPaths) });
    }
  } else if (type === 'gtm_client') {
    if (activateResponseCompression !== '') {
      newClient.parameter.push({ type: 'boolean', key: 'activateResponseCompression', value: String(activateResponseCompression) });
    }
    if (isPathMandatory !== '') {
      newClient.parameter.push({ type: 'boolean', key: 'isPathMandatory', value: String(isPathMandatory) });
    }
    if (activateGeoResolution !== '') {
      newClient.parameter.push({ type: 'boolean', key: 'activateGeoResolution', value: String(activateGeoResolution) });
      if (activateGeoResolution === true && region) {
        newClient.parameter.push({ type: 'template', key: 'region', value: region });
      }
    }

    if (allowedContainersStr) {
      const allowedContainerList = [];
      allowedContainersStr.split(';').forEach(pair => {
        const parts = pair.split(':');
        if (parts.length >= 1) {
          const cid = parts[0].trim();
          const path = parts[1] ? parts[1].trim() : '';
          allowedContainerList.push({
            type: 'map',
            map: [
              { type: 'template', key: 'containerId', value: cid },
              { type: 'template', key: 'path', value: path }
            ]
          });
        }
      });
      if (allowedContainerList.length > 0) {
        newClient.parameter.push({
          type: 'list',
          key: 'allowedContainerIds',
          list: allowedContainerList
        });
      }
    }
  }

  if (modifyOption === 'update' && clients) {
    const oldClient = clients.find(c => c.clientId === settings[3]);
    if (oldClient && oldClient.parameter) {
      oldClient.parameter.forEach(param => {
        const isOverridden = newClient.parameter.some(newParam => newParam.key === param.key);
        if (!isOverridden) {
          newClient.parameter.push(param);
        }
      });
    }
  }

  return newClient;
}
