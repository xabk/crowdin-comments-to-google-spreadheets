// --- Script Parameters: see Readme.md for details ---
var org = null; // Replace with your organization name or leave blank if you're using Crowdin.com
var projectID = null; // Replace with Project ID (under Tools → API on Crowdin)
var token = null;
// Leave empty and users will be prompted for it, and it'll be saved (per user)
// You can also set or delete this via Set or Delete API Key in Loc Tools menu
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
}

// ===========================================================
// API Credential Management
// ===========================================================

function checkAPICredentials() {
  if (org == null) {
    setOrg();
    org = scriptProperties.getProperty("crowdinOrg");
    if (org == null) {
      Logger.log("Error: Couldn't get the organization name from the user");
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
      Logger.log("Error: Couldn't get the project ID from the user");
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
      Logger.log("Error: Couldn't get the API token from the user");
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
    Logger.log("API key saved");
    logMessage("API key saved");
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
    logMessage("API key deleted");
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
    Logger.log("Project ID saved");
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
    Logger.log("Organization name saved");
    logMessage("Organization name saved");
  }
}
