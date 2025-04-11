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

function getIssuesFromCrowdin() {
  if (!token) {
    logMessage("Error: No API token available.");
    ui.alert("Error: No API token available. Please set one via the menu.");
    return [];
  }

  var projectLinkID = getProjectLinkID();
  var files = getFileNamesFromCrowdin();
  var data = crowdinAPIFetchAllData("/comments");
  var issues = new Array();

  if (!data || data.length === 0) {
    logMessage("Error: No issues or comments fetched from Crowdin.");
    return issues;
  }

  for (issueData of data) {
    issueData = issueData.data;
    let issue = new Object();

    issue.id = issueData.id;
    issue.date = new Date(issueData.createdAt);
    // issue.dateString = new Intl.DateTimeFormat("en" , {dateStyle: "short"}).format(issue.date)

    issue.text = decodeHTMLEntities(issueData.text);
    issue.context = decodeHTMLEntities(issueData.string.context);

    if (issueData.languageId != null) {
      issue.language =
        issueData.languageId.charAt(0).toUpperCase() +
        issueData.languageId.slice(1);
    } else {
      issue.language = "—";
    }

    if (issueData.type == "comment") {
      issue.status = "Comment";
      issue.type = "Comment";
    } else {
      issue.status = issueData.issueStatus
        .replace(/unresolved/, "Open")
        .replace(/resolved/, "Closed");
      issue.type = (
        issueData.issueType.charAt(0).toUpperCase() +
        issueData.issueType.slice(1)
      ).replace("_", " ");
    }

    issue.stringKey = issueData.string.context.split("\n")[0];

    issue.stringID = issueData.stringId;
    issue.stringText = decodeHTMLEntities(issueData.string.text);

    issue.userID = issueData.userId;
    issue.user = issueData.user.username;
    if (
      issueData.user.fullName &
      (issueData.user.fullName != issueData.user.username)
    ) {
      issue.user += `\n${issueData.user.fullName}`;
    }
    issue.user += `\n\nLang: ${issue.language}`;

    issue.fileID = issueData.string.fileId;
    issue.fileIDandName = `${issueData.string.fileId}\n\n${files.get(
      issueData.string.fileId
    )}`;

    issue.link =
      org == ""
        ? `https://crowdin.com/translate/${projectLinkID}/all/en-XX#${issue.stringID}`
        : `https://${org}.crowdin.com/translate/${projectLinkID}/all/en-XX#${issue.stringID}`;

    issues.push(issue);
  }

  issues.sort((firstEl, secondEl) => {
    if (firstEl.date != secondEl.date) {
      return firstEl.date - secondEl.date;
    } else {
      return firstEl.id - secondEl.id;
    }
  });

  var issuesMap = new Map();
  for (issue of issues) {
    if (!issuesMap.has(issue.stringID)) {
      issue.firstForString = true;
      issuesMap.set(issue.stringID, [issue]);
    } else {
      issuesMap.get(issue.stringID).push(issue);
    }
  }

  issues = new Array();

  for ([key, value] of issuesMap) {
    for (item of value) {
      issues.push(item);
    }
  }

  logMessage(`Fetched ${issues.length} issues from Crowdin.`);
  return issues;
}
