function getOrCreateLogsSheet() {
  let logsSheet = ss.getSheetByName(LOGS_SHEET_NAME);
  if (!logsSheet) {
    logsSheet = ss.insertSheet(LOGS_SHEET_NAME);
    logsSheet.appendRow(["Timestamp", "Message"]);
    logsSheet.setColumnWidth(1, 200);
    logsSheet.setColumnWidth(2, 600);
  }
  return logsSheet;
}

function logMessage(message) {
  Logger.log(message);
  const logsSheet = getOrCreateLogsSheet();
  logsSheet.insertRowBefore(2);
  logsSheet.getRange(2, 1, 1, 2).setValues([[new Date(), message]]);
}
