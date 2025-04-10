function crowdinAPIGetResponseData(addAPIPath) {
  var url =
    addAPIPath != null
      ? `${apiBaseURL}/projects/${projectID}${addAPIPath}`
      : `${apiBaseURL}/projects/${projectID}`;
  var options = {
    muteHttpExceptions: true,
    headers: {
      Authorization: `Bearer ${token}`,
    },
  };

  var response = UrlFetchApp.fetch(url, options);
  var json = response.getContentText();
  var data = JSON.parse(json);

  if ("data" in data) {
    return data.data;
  }

  Logger.log(`No data in response for ${url}`);
  Logger.log(`Response: ${response}`);
  logMessage(`No data in response for ${url}`);
  return null;
}

function crowdinAPIFetchAllData(addAPIPath) {
  var offset = 0;
  var addURL = `${addAPIPath}?limit=${apiLimit}&offset=`;

  data = new Array();

  while (true) {
    newData = crowdinAPIGetResponseData(addURL + offset);
    if (newData.length == 0) {
      break;
    }
    data.push(...newData);
    if (newData.length < apiLimit) {
      break;
    }
    offset += apiLimit;
  }

  logMessage(`Fetched ${data.length} entries from Crowdin for ${addAPIPath}`);
  return data;
}

//
// Get Project Link ID
// (project name on Crowdin.com or alphanumeric project ID used for URLs in Crowdin Enterprise)
//

function getProjectLinkID() {
  data = crowdinAPIGetResponseData();

  if (!data || !("identifier" in data)) {
    Logger.log(`Error: Couldn't get project identifier... Data: ${data}`);
    logMessage(`Error: Couldn't get project identifier.`);
    return null;
  }

  Logger.log(`Fetched project identifier: ${data.identifier}`);
  logMessage(`Fetched project identifier: ${data.identifier}`);
  return data.identifier;
}

//
// Downloading file list from Crowdin (to have names)
//

function getFileNamesFromCrowdin() {
  var data = crowdinAPIFetchAllData("/files");

  if (data == null) {
    Logger.log(
      "Error: Something went wrong while trying to get file names from Crowdin."
    );
    logMessage(
      "Error: Something went wrong while trying to get file names from Crowdin."
    );
  }

  if (data.length == 0) {
    Logger.log(`Error: Got an empty list of files from Crowdin: ${data}`);
    logMessage("Error: Got an empty list of files from Crowdin.");
  }

  var files = new Map();

  data.forEach(function (item) {
    if (item.data.title) {
      files.set(item.data.id, item.data.title);
    } else {
      files.set(item.data.id, item.data.name);
    }
  });

  Logger.log(`Fetched file names from Crowdin: ${data.length} file(s)`);
  logMessage(`Fetched file names from Crowdin: ${data.length} file(s)`);

  return files;
}
