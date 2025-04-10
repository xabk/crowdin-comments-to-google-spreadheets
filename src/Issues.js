function getOrCreateAllIssuesSheet() {
  let allIssuesSheet = ss.getSheetByName(issuesSheetName);
  if (!allIssuesSheet) {
    allIssuesSheet = ss.insertSheet(issuesSheetName);
  }
  return allIssuesSheet;
}

function getIssuesFromCrowdin() {
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

  var sheet = getOrCreateAllIssuesSheet();
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

  sheet
    .getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
    .setFontFamily(fontName)
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
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }
  sheet.getDataRange().createFilter();

  logMessage("Successfully overwrote the sheet with issues from Crowdin.");
}
