// --- Script Parameters: see Readme.md for details ---
var org = null; // Replace with your organization name or leave blank if you're using Crowdin.com
var projectID = null; // Replace with Project ID (under Tools → API on Crowdin)
var token = null;
// Leave empty and users will be prompted for it, and it'll be saved (per user)
// You can also set or delete this via Set or Delete API token in Loc Tools menu
// --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
// You can also hardcode your token here but it's very insecure
// Hardcoded values take precedence over values set via menu
// WARNING: only you save your API token here if the spreadsheet is never shared outside your trusted circle
// WARNING: API token saved here will be freely available to anyone who has access to the spreadsheet
// --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---

// ===========================================================
// Script Parameters
// ===========================================================

const logsSheetName = "Logs";
const issuesSheetName = "All Issues";

const fontName = "Calibri";

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

// ===========================================================
// Initialization and Menu Setup
// ===========================================================

var scriptProperties = PropertiesService.getScriptProperties();
var userProperties = PropertiesService.getUserProperties();

if (token == null) {
  token = userProperties.getProperty("crowdinAPIKey");
}

if (token == null) {
  token = scriptProperties.getProperty("crowdinDocAPIKey"); // Check document-level token
}

if (projectID == null) {
  projectID = scriptProperties.getProperty("crowdinProjectID");
}

if (org == null) {
  org = scriptProperties.getProperty("crowdinOrg");
}

const apiLimit = 500;
const apiBaseURL =
  org == ""
    ? "https://api.crowdin.com/api/v2"
    : `https://${org}.api.crowdin.com/api/v2`;

var ss = SpreadsheetApp.getActiveSpreadsheet();
var ui = SpreadsheetApp.getUi();

function onOpen() {
  ui.createMenu("Loc Tools")
    .addItem(
      "Pull comments and issues from Crowdin",
      "overwriteWithIssuesFromCrowdin"
    )
    .addSeparator()
    .addItem("Set Crowdin API token (current Google user)", "setAPIKey")
    .addItem("Set Crowdin API token (document-wide)", "setDocAPIKey")
    .addItem("Delete API tokens (user and document-wide)", "deleteAPIKey")
    .addSeparator()
    .addItem("View or set Crowdin project ID (all users)", "setProjectID")
    .addItem("View or set Crowdin org (all users)", "setOrg")
    .addToUi();
}

// ===========================================================
// API Credential Management
// ===========================================================

function checkAPICredentials() {
  if (org == null) {
    setOrg();
    org = scriptProperties.getProperty("crowdinOrg");
    if (org == null) {
      logMessage("Error: Couldn't get the organization name from the user");
      ui.alert(
        "Error: We need a valid or empty organization name to fetch data from Crowdin. Please try again."
      );
      return false;
    }
  }

  if (projectID == null) {
    setProjectID();
    projectID = scriptProperties.getProperty("crowdinProjectID");
    if (projectID == null) {
      logMessage("Error: Couldn't get the project ID from the user");
      ui.alert(
        "Error: We need a valid project ID to fetch data from Crowdin. Please try again."
      );
      return false;
    }
  }

  if (token == null) {
    setAPIKey();
    token = userProperties.getProperty("crowdinAPIKey");
    if (token == null) {
      logMessage("Error: Couldn't get the API token from the user");
      ui.alert(
        "Error: We need a valid token to fetch data from Crowdin. Please try again."
      );
      return false;
    }
  }

  logMessage("API credentials successfully validated.");
  return true;
}

function setAPIKey() {
  var result = ui.prompt(
    "Please enter the Crowdin API token (don't share your token anywhere!)",
    ui.ButtonSet.OK_CANCEL
  );
  var button = result.getSelectedButton();
  if (button === ui.Button.OK) {
    userProperties.setProperty("crowdinAPIKey", result.getResponseText());
    logMessage("API token saved");
  }
}

function setDocAPIKey() {
  var result = ui.prompt(
    "Please enter the Crowdin API token for this document (don't share your token anywhere!)",
    ui.ButtonSet.OK_CANCEL
  );
  var button = result.getSelectedButton();
  if (button === ui.Button.OK) {
    scriptProperties.setProperty("crowdinDocAPIKey", result.getResponseText());
    logMessage("Document-wide API token saved");
  }
}

function deleteAPIKey() {
  var result = ui.prompt(
    "Are you sure you want to delete the API token? (it will be deleted for you and all users)",
    ui.ButtonSet.YES_NO
  );
  var button = result.getSelectedButton();
  if (button === ui.Button.YES) {
    userProperties.deleteProperty("crowdinAPIKey");
    scriptProperties.deleteProperty("crowdinDocAPIKey");
    logMessage("API token deleted for you and all users");
  }
}

function setProjectID() {
  var text = "Please enter the Crowdin Project ID";
  if (projectID != null) {
    text += `.\nCurrent project ID: ${projectID}. Click Cancel to keep it`;
  }
  var result = ui.prompt(text, ui.ButtonSet.OK_CANCEL);
  var button = result.getSelectedButton();
  if (button === ui.Button.OK) {
    scriptProperties.setProperty("crowdinProjectID", result.getResponseText());
    logMessage("Project ID saved");
  }
}

function setOrg() {
  var text =
    "Please enter the Crowdin organization name (leave blank if you're using Crowdin.com)";
  if (org != null) {
    text += `.\nCurrent organization name: ${
      org ? org : "<Crowdin.com>"
    }. Click Cancel to keep it`;
  }
  var result = ui.prompt(text, ui.ButtonSet.OK_CANCEL);
  var button = result.getSelectedButton();
  if (button === ui.Button.OK) {
    scriptProperties.setProperty("crowdinOrg", result.getResponseText());
    logMessage("Organization name saved");
  }
}
