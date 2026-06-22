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
    while (tempArray.length < 12) {
      tempArray.push('');
    }
    
    // Find parameters
    const params = getClientParams(client);
    
    tempArray[0] = workspaceName;
    tempArray[1] = workspacePath;
    tempArray[2] = client.name;
    tempArray[3] = client.clientId;
    tempArray[4] = client.type;
    tempArray[5] = client.priority !== undefined ? client.priority : '';
    tempArray[6] = client.paused || false;
    tempArray[7] = params['measurementId'] || '';
    tempArray[8] = params['cookiePaths'] || '';
    tempArray[9] = params['allowedOrigins'] || '';
    tempArray[10] = convertFolderIdsToNames(client.parentFolderId, folders);
    tempArray[11] = client.notes || '';
    
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
      if (p.type == 'template' || p.type == 'boolean' || p.type == 'integer') {
        params[p.key] = p.value;
      } else if (p.type == 'list' && p.list) {
        params[p.key] = p.list.map(li => li.value).join('; ');
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
  const clientSettingsData = getDataFromSheet('sgtmClients', 'clients');
  const folders = listGTMResources('folders', workspacePath);

  clientSettingsData.forEach(newSettings => {
    const create = newSettings[newSettings.length - 3];
    const update = newSettings[newSettings.length - 2];
    const remove = newSettings[newSettings.length - 1];
    
    if (create && !update && !remove) {
      const payload = buildSgtmClientPayload(newSettings, folders, 'create');
      const response = createGTMResource('clients', newSettings[1], payload);
      if (response && response.tagManagerUrl) {
        logChange(response.name, 'client', response.clientId, 'created', response.tagManagerUrl);
      } else {
        logChange(newSettings[2], 'client', '', errorResponse(response), '');
      }
    } else if (!create && update && !remove) {
      const payload = buildSgtmClientPayload(newSettings, folders, 'update');
      const response = updateGTMResource('clients', `${newSettings[1]}/clients/${newSettings[3]}`, payload);
      if (response && response.tagManagerUrl) {
        logChange(response.name, 'client', response.clientId, 'updated', response.tagManagerUrl);
      } else {
        logChange(newSettings[2], 'client', '', errorResponse(response), '');
      }
    } else if (!create && !update && remove) {
      const response = removeGTMResource('clients', `${newSettings[1]}/clients/${newSettings[3]}`);
      if (response && response.details) {
        logChange(newSettings[2], 'client', '', errorResponse(response), '');
      } else {
        logChange(newSettings[2], 'client', newSettings[3], 'removed', '');
      }
    }
  });
}

/**
 * Builds the GTM Client payload.
 * @param {!Array} settings Client settings from sheet.
 * @param {!Array} folders Folders array.
 * @param {string} modifyOption Modify action (create or update).
 * @return {!Object} The Client payload structure.
 */
function buildSgtmClientPayload(settings, folders, modifyOption) {
  const name = settings[2];
  const type = settings[4];
  const priority = settings[5];
  const paused = Boolean(settings[6]);
  const measurementId = settings[7].trim();
  const cookiePaths = settings[8].trim();
  const allowedOrigins = settings[9].trim();
  const parentFolderName = settings[10];
  const notes = settings[11];
  
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

  if (measurementId) {
    newClient.parameter.push({
      type: 'template',
      key: 'measurementId',
      value: measurementId
    });
  }

  if (cookiePaths) {
    newClient.parameter.push({
      type: 'list',
      key: 'cookiePaths',
      list: cookiePaths.split(';').map(path => ({
        type: 'template',
        value: path.trim()
      }))
    });
  }

  if (allowedOrigins) {
    newClient.parameter.push({
      type: 'list',
      key: 'allowedOrigins',
      list: allowedOrigins.split(';').map(origin => ({
        type: 'template',
        value: origin.trim()
      }))
    });
  }

  return newClient;
}
