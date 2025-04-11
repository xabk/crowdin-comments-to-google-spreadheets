function overwriteWithIssuesFromCrowdin() {
  if (!checkAPICredentials()) {
    return false;
  }

  logMessage("Starting to overwrite sheet with issues from Crowdin...");
  var issues = getIssuesFromCrowdin();

  if (issues.length === 0) {
    logMessage("No issues found to write to the sheet.");
    return;
  }

  // Create or update the main issues sheet
  createAllIssuesSheet(issues);

  // Create source issues sheet
  createSourceIssuesSheet(issues);

  // Create context and general questions sheet
  createContextAndGeneralQuestionsSheet(issues);

  // Create language-specific sheets
  createLanguageSheets(issues);
}

function createAllIssuesSheet(issues) {
  sheetName = ISSUES_SHEET_NAME_ALL;

  createOrUpdateSheet(sheetName, issues);

  sheet = ss.getSheetByName(sheetName);
  lastColumn = sheet.getLastColumn();

  // Add borders between strings to group issues per string visually
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

  // Apply conditional formatting rules
  // TODO: Extract the rules to the top/config part of the script
  const existingRules = sheet.getConditionalFormatRules();

  // Hide some repeating data to make the spreadsheet more readable
  const newRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$C2=$C1")
    .setFontColor("#ffffff")
    .setRanges([
      sheet.getRange(2, 2, sheet.getLastRow(), 1),
      sheet.getRange(2, 8, sheet.getLastRow(), 1),
      sheet.getRange(2, 10, sheet.getLastRow(), 1),
    ])
    .build();

  sheet.setConditionalFormatRules([...existingRules, newRule]);

  logMessage("Successfully overwrote the sheet with issues from Crowdin.");
}

function createSourceIssuesSheet(issues) {
  const sourceMistakes = issues.filter(
    (issue) => issue.type === "Source mistake"
  );

  if (sourceMistakes.length > 0) {
    const sheetName = ISSUES_SHEET_NAME_SOURCE;

    createOrUpdateSheet(sheetName, sourceMistakes);

    logMessage("Created/updated Source Issues sheet.");
  }
}

function createContextAndGeneralQuestionsSheet(issues) {
  const contextAndGeneralQuestions = issues.filter(
    (issue) =>
      issue.type === "Context request" || issue.type === "General question"
  );

  if (contextAndGeneralQuestions.length > 0) {
    const sheetName = ISSUES_SHEET_NAME_CONTEXT;

    createOrUpdateSheet(sheetName, contextAndGeneralQuestions);

    logMessage("Created/updated Context and General Questions sheet.");
  }
}

function createLanguageSheets(issues) {
  const translationMistakes = issues.filter(
    (issue) => issue.type === "Translation mistake"
  );

  const issuesByLanguage = groupIssuesByLanguage(translationMistakes);

  for (const [language, languageIssues] of issuesByLanguage.entries()) {
    const sheetName = ISSUES_SHEET_NAME_LANG.replace("%lang%", language);

    createOrUpdateSheet(sheetName, languageIssues);

    logMessage(`Created/updated sheet for language: ${language}`);
  }
}

function groupIssuesByLanguage(issues) {
  const issuesByLanguage = new Map();

  for (const issue of issues) {
    const language = issue.language || "Unknown";

    if (!issuesByLanguage.has(language)) {
      issuesByLanguage.set(language, []);
    }

    issuesByLanguage.get(language).push(issue);
  }

  return issuesByLanguage;
}

function createOrUpdateSheet(sheetName, issues) {
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  // Prepare data for the sheet
  const issueArray = [];
  const stringIDsAndLinks = [];

  issues.forEach((issue) => {
    issueArray.push([
      issue.id,
      issue.fileIDandName,
      issue.stringID,
      issue.date,
      issue.user,
      issue.status,
      issue.type,
      issue.stringText,
      issue.text,
      issue.context,
    ]);

    stringIDsAndLinks.push([
      prepareIssueLinks(issue, ADD_LINK_URL, ADD_LINK_TEXT, ADD_LINK),
    ]);
  });

  const totalRows = issueArray.length + 1; // Include header row
  const totalColumns = header[0].length;

  // Resize the sheet only if necessary
  if (sheet.getMaxRows() < totalRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), totalRows - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < totalColumns) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      totalColumns - sheet.getMaxColumns()
    );
  }

  // Reset formatting for the entire sheet
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearFormat();

  // Set headers and data in a single batch
  sheet
    .getRange(1, 1, 1, totalColumns)
    .setValues(header)
    .setFontColor("#f3f3f3")
    .setBackground("#434343")
    .setFontWeight("bold");
  sheet.getRange(2, 1, issueArray.length, totalColumns).setValues(issueArray);

  // Set rich text links for the string ID column
  sheet.getRange(2, 3, issueArray.length).setRichTextValues(stringIDsAndLinks);

  // Apply column widths explicitly
  colWidths.forEach((width, col) => {
    sheet.setColumnWidth(col + 1, width);
  });

  // Apply formatting in bulk
  const dataRange = sheet.getRange(1, 1, totalRows, totalColumns);
  dataRange.setFontFamily(fontName).setFontSize(11);

  wrapStrategies.forEach((strat, col) => {
    sheet.getRange(2, col + 1, issueArray.length, 1).setWrapStrategy(strat);
  });

  colFontColors.forEach((color, col) => {
    sheet.getRange(2, col + 1, issueArray.length, 1).setFontColor(color);
  });

  // Apply conditional formatting rules
  sheet.setConditionalFormatRules([
    // Highlight open issues
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$F2="Open"')
      .setBackground(OPEN_ISSUE_COLOR)
      .setRanges([sheet.getRange(2, 1, sheet.getLastRow(), totalColumns)])
      .build(),
  ]);

  // Remove extra rows and columns if necessary
  if (sheet.getMaxRows() > totalRows) {
    sheet.deleteRows(totalRows + 1, sheet.getMaxRows() - totalRows);
  }
  if (sheet.getMaxColumns() > totalColumns) {
    sheet.deleteColumns(totalColumns + 1, sheet.getMaxColumns() - totalColumns);
  }

  // Recreate filter to cover the whole range
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }
  sheet.getDataRange().createFilter();
}
