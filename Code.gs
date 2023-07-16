// ---  Script Parameters: see Readme.md for details   ---
var org = null; // Replace with your organization name or leave blank if you're using Crowdin.com
// You can also set this via Set org in Loc Tools menu
// Hardcoded values take precedence over values set via menu

var projectID = null; // Replace with Project ID (under Tools → API on Crowdin)
// You can also set this via Set Project ID in Loc Tools menu
// Hardcoded values take precedence over values set via menu

var token = null;
// Leave empty and users will be prompted for it, and it'll be saved (per user)
// You can also set or delete this via Set or Delete API Key in Loc Tools menu
// --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
// You can also hardcode your token here but it's very insecure
// Hardcoded values take precedence over values set via menu
// WARNING: only you save your API token here if the spreadsheet is never shared outside your trusted circle
// WARNING: API token saved here will be freely available to anyone who has access to the spreadsheet
// --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---

const wrClip = SpreadsheetApp.WrapStrategy.CLIP;
const wrWrap = SpreadsheetApp.WrapStrategy.WRAP;

const header = [
  [
    "iID",
    "File ID / \nName",
    "Str\nID",
    "Date",
    "User",
    "Status",
    "Issue Type",
    "String",
    "Issue/Comment",
    "Context",
  ],
];
const colWidths = [40, 100, 50, 100, 110, 80, 80, 230, 450, 450];
const colFontColors = [
  "#555",
  "#555",
  "#1155cc",
  "#555",
  "#555",
  "#555",
  "#555",
  "#000",
  "#000",
  "#000",
];
const wrapStrategies = [
  wrClip,
  wrClip,
  wrClip,
  wrClip,
  wrClip,
  wrWrap,
  wrWrap,
  wrWrap,
  wrWrap,
  wrClip,
];

const apiLimit = 500;
const apiBaseURL =
  org == ""
    ? "https://api.crowdin.com/api/v2"
    : `https://${org}.api.crowdin.com/api/v2`;

var ss = SpreadsheetApp.getActiveSpreadsheet();
var ui = SpreadsheetApp.getUi();
var scriptProperties = PropertiesService.getScriptProperties();
var userProperties = PropertiesService.getUserProperties();

//
// Adding functions to the menu
//

function onOpen() {
  ui.createMenu("Loc Tools")
    .addItem(
      "Overwrite current sheet with comments from Crowdin",
      "overwriteWithIssuesFromCrowdin"
    )
    .addSeparator()
    .addItem("Set Crowdin API key (current Google user)", "setAPIKey")
    .addItem(
      "Delete saved Crowdin API key (current Google user)",
      "deleteAPIKey"
    )
    .addItem("View or set Crowdin project ID (all users)", "setProjectID")
    .addItem("View or set Crowdin org (all users)", "setOrg")
    .addToUi();

  if (token == null && userProperties.getProperty("crowdinAPIKey") != null) {
    token = userProperties.getProperty("crowdinAPIKey");
  }

  if (
    projectID == null &&
    scriptProperties.getProperty("crowdinProjectID") != null
  ) {
    projectID = scriptProperties.getProperty("crowdinProjectID");
  }

  if (org == null && scriptProperties.getProperty("crowdinOrg") != null) {
    org = scriptProperties.getProperty("crowdinOrg");
  }
}

//
// Get Project Link ID
// (project name on Crowdin.com or alphanumeric project ID used for URLs in Crowdin Enterprise)
//

function getProjectLinkID() {
  data = crowdinAPIGetResponseData();

  if (!data || !("identifier" in data)) {
    Logger.log(`Error: Couldn't get project identifier... Data: ${data}`);
    return null;
  }

  Logger.log(`Fetched project identifier: ${data.identifier}`);
  return data.identifier;
}

//
// Downloading file list from Crowdin (to have names)
//

function getFileNamesFromCrowdin() {
  var data = crowdinAPIFetchAllData("/files");

  if (data == null) {
    Logger.log(
      "Error: Something went wrong while trying to get file names from Crowdin..."
    );
  }

  if (data.length == 0) {
    Logger.log(`Error: Got an empty list of files from Crowdin: ${data}`);
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

  return files;
}

//
// Downloading issues and comments from Crowdin
//

function getIssuesFromCrowdin() {
  var projectLinkID = getProjectLinkID();
  var files = getFileNamesFromCrowdin();
  var data = crowdinAPIFetchAllData("/comments");
  var issues = new Array();

  for (issueData of data) {
    issueData = issueData.data;
    let issue = new Object();

    issue.id = issueData.id;
    issue.date = new Date(issueData.createdAt);
    // issue.dateString = new Intl.DateTimeFormat("en" , {dateStyle: "short"}).format(issue.date)

    issue.text = decodeHTMLEntities(issueData.text);
    issue.context = decodeHTMLEntities(issueData.string.context);

    issue.language =
      issueData.languageId.charAt(0).toUpperCase() +
      issueData.languageId.slice(1);

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

  return issues;
}

function overwriteWithIssuesFromCrowdin() {
  if (org == null) {
    setOrg();
    org = scriptProperties.getProperty("crowdinOrg");
    if (org == null) {
      Logger.log("Error: Couldn't get the organization name from the user");
      ui.alert(
        "Error: We need a valid or empty organization name to fetch data from Crowdin. Please try again."
      );
      return;
    }
  }

  if (projectID == null) {
    setProjectID();
    projectID = scriptProperties.getProperty("crowdinProjectID");
    if (projectID == null) {
      Logger.log("Error: Couldn't get the project ID from the user");
      ui.alert(
        "Error: We need a valid project ID to fetch data from Crowdin. Please try again."
      );
      return;
    }
  }

  if (token == null) {
    setAPIKey();
    token = userProperties.getProperty("crowdinAPIKey");
    if (token == null) {
      Logger.log("Error: Couldn't get the API token from the user");
      ui.alert(
        "Error: We need a valid token to fetch data from Crowdin. Please try again."
      );
      return;
    }
  }

  var issues = getIssuesFromCrowdin();

  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var lastColumn = header[0].length;
  var maxColumns = sheet.getMaxColumns();

  if (maxColumns < header[0].length) {
    sheet.insertColumnsAfter(maxColumns, header[0].length - maxColumns);
  }

  sheet
    .getRange(1, 1, 1, lastColumn)
    .setValues(header)
    .setFontColor("#f3f3f3")
    .setBackground("#434343")
    .setFontWeight("bold");
  colWidths.map((width, col) => {
    sheet.setColumnWidth(col + 1, width);
  });
  sheet.setFrozenRows(1);

  sheet
    .getRange(2, 1, sheet.getMaxRows() - 1, sheet.getMaxColumns())
    .clearContent();
  sheet
    .getRange(2, 1, sheet.getMaxRows() - 1, sheet.getMaxColumns())
    .setBorder(false, false, false, false, false, false);

  issueArray = new Array();
  stringIDsAndLinks = new Array();

  for (i of issues) {
    issueArray.push([
      i.id,
      i.fileIDandName,
      i.stringID,
      i.date,
      i.user,
      i.status,
      i.type,
      i.stringText,
      i.text,
      i.context,
    ]);
    stringIDsAndLinks.push([createRichTextLink(i.stringID, i.link)]);

    /*  let styleGrey = SpreadsheetApp.newTextStyle().setForegroundColor('#aaa').build()
    let fileIDandName = SpreadsheetApp.newRichTextValue()
                    .setText(i.fileIDandName)
                    .setTextStyle(0, String(i.fileID).length, styleGrey)
                    .build()
    var row = [i.id, fileIDandName, i.stringID, i.date, i.user, i.status, i.type, i.text, i.stringText, i.context]
    row = row.map(function(item){
      if (item.getRuns) {
        return item
      } else {
        return SpreadsheetApp.newRichTextValue().setText(item).build()
      }
    })

    issueArray.push(row) */
  }

  sheet.getRange(2, 1, dataRange.getLastRow(), lastColumn).clearContent();
  sheet
    .getRange(2, 1, issueArray.length, issueArray[0].length)
    .setValues(issueArray);

  sheet.getRange(2, 3, issueArray.length).setRichTextValues(stringIDsAndLinks);

  for (let [index, val] of issues.entries()) {
    if (val.firstForString) {
      sheet
        .getRange(index + 2, 1, index + 2, lastColumn)
        .setBorder(
          true,
          null,
          null,
          null,
          null,
          null,
          "#000000",
          SpreadsheetApp.BorderStyle.SOLID_MEDIUM
        );
    }
  }

  // TODO: Extract font to config/params?
  sheet
    .getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
    .setFontFamily("Calibri")
    .setFontSize(11);

  wrapStrategies.map((strat, col) => {
    sheet
      .getRange(2, col + 1, sheet.getDataRange().getLastRow(), 1)
      .setWrapStrategy(strat);
  });

  colFontColors.map((color, col) => {
    sheet
      .getRange(2, col + 1, sheet.getDataRange().getLastRow(), 1)
      .setFontColor(color);
  });

  // Apply conditional formatting rules
  // TODO: Extract the rules to the top/config part of the script
  sheet.setConditionalFormatRules([
    // Hightlight Source mistakes
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($F2="Open",$G2="Source Mistake")')
      .setBackground("#fff2cc")
      .setRanges([sheet.getRange(2, 1, sheet.getLastRow(), 10)])
      .build(),
    // Hide (white font color) duplicate text in File, String, Context
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=$C2=$C1")
      .setFontColor("#ffffff")
      .setRanges([
        sheet.getRange(2, 2, sheet.getLastRow(), 1),
        sheet.getRange(2, 8, sheet.getLastRow(), 1),
        sheet.getRange(2, 10, sheet.getLastRow(), 1),
      ])
      .build(),
  ]);

  // Recreate filter to cover the whole range
  sheet.getFilter().remove();
  sheet.getDataRange().createFilter();
}

//
// ==============================================================================================================================================
//
// Service Functions
//

//
// Get API token from user
//

function getTokenPrompt() {
  var result = ui.prompt(
    "Please enter the Crowdin API token (don't share your token anywhere!)"
  );

  //Get the button that the user pressed.
  var button = result.getSelectedButton();

  if (button === ui.Button.OK) {
    Logger.log("Got the token from the user.");
    return result.getResponseText();
  } else if (button === ui.Button.CLOSE) {
    Logger.log("Script cancelled, user haven't supplied the API token");
    return null;
  }
}

function setAPIKey() {
  var result = ui.prompt(
    "Please enter the Crowdin API token (don't share your token anywhere!)",
    ui.ButtonSet.OK_CANCEL
  );
  var button = result.getSelectedButton();
  if (button === ui.Button.OK) {
    userProperties.setProperty("crowdinAPIKey", result.getResponseText());
    Logger.log("API key saved");
  }
}

function deleteAPIKey() {
  var result = ui.prompt(
    "Are you sure you want to delete the API key? (it will be deleted for all users)",
    ui.ButtonSet.YES_NO
  );
  var button = result.getSelectedButton();
  if (button === ui.Button.YES) {
    userProperties.deleteProperty("crowdinAPIKey");
    Logger.log("API key deleted");
  }
}

function setProjectID() {
  var result = ui.prompt(
    "Please enter the Crowdin Project ID",
    ui.ButtonSet.OK_CANCEL
  );
  var button = result.getSelectedButton();
  if (button === ui.Button.OK) {
    scriptProperties.setProperty("crowdinProjectID", result.getResponseText());
    Logger.log("Project ID saved");
  }
}

function setOrg() {
  var result = ui.prompt(
    "Please enter the Crowdin organization name (leave blank if you're using Crowdin.com)",
    ui.ButtonSet.OK_CANCEL
  );
  var button = result.getSelectedButton();
  if (button === ui.Button.OK) {
    scriptProperties.setProperty("crowdinOrg", result.getResponseText());
    Logger.log("Organization name saved");
  }
}

//
// Make an API request and return 'data' if it exists
// Return null if data isn't present
//

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
  return null;
}

//
// Make an API request and fetch all 'data' into an array
//

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

  return data;
}

function prettylog(object) {
  Logger.log(JSON.stringify(object, null, 2));
}

function decodeHTMLEntities(string) {
  return string.replace(/&quot;/g, '"').replace(/&amp;/g, "&");
  // return XmlService.parse(`<d>${string}</d>`).getRootElement().getText() // Doesn't work with Unreal rich text tags (<hunter>...</>)
}

function createRichTextLink(string, link) {
  if (link) {
    return SpreadsheetApp.newRichTextValue()
      .setText(string)
      .setLinkUrl(0, String(string).length, link)
      .build();
  }
}
