var LOCAL_CONFIG = {
  SPREADSHEET_ID: "PUT_YOUR_SPREADSHEET_ID_HERE",
};

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("🔔 Expiry Notifications")
    .addItem("Run Daily Check", "runDailyCheck")
    .addItem("Run Reply Scan", "runReplyScan")
    .addToUi();
}

function runDailyCheck() {
  return ExpiryLib.runDailyCheckForSpreadsheet(LOCAL_CONFIG.SPREADSHEET_ID);
}

function runReplyScan() {
  return ExpiryLib.runReplyScanForSpreadsheet(LOCAL_CONFIG.SPREADSHEET_ID);
}
