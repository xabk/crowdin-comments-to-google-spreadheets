function createLanguageSheets(issues) {
  const translationMistakes = issues.filter(
    (issue) => issue.type === "Translation mistake"
  );

  const issuesByLanguage = groupIssuesByLanguage(translationMistakes);

  for (const [language, languageIssues] of issuesByLanguage.entries()) {
    const sheetName = `${issuesSheetName} (${language})`;
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    // Prepare data for the sheet
    const issueArray = languageIssues.map((issue) => [
      issue.id,
      issue.fileIDandName,
      `=HYPERLINK("${issue.link.replace(
        "/en-XX",
        `/en-${issue.language.toLowerCase()}`
      )}", "${issue.stringID}")`, // Adjust link for language
      issue.date,
      issue.user,
      issue.status,
      issue.type,
      issue.stringText,
      issue.text,
      issue.context,
    ]);

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
    sheet
      .getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
      .clearFormat();

    // Set headers and data in a single batch
    sheet
      .getRange(1, 1, 1, totalColumns)
      .setValues(header)
      .setFontColor("#f3f3f3")
      .setBackground("#434343")
      .setFontWeight("bold");
    sheet.getRange(2, 1, issueArray.length, totalColumns).setValues(issueArray);

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
        .setBackground("#fbfbd8")
        .setRanges([sheet.getRange(2, 1, sheet.getLastRow(), totalColumns)])
        .build(),
    ]);

    // Remove extra rows and columns if necessary
    if (sheet.getMaxRows() > totalRows) {
      sheet.deleteRows(totalRows + 1, sheet.getMaxRows() - totalRows);
    }
    if (sheet.getMaxColumns() > totalColumns) {
      sheet.deleteColumns(
        totalColumns + 1,
        sheet.getMaxColumns() - totalColumns
      );
    }

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
