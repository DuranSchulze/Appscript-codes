// 42 Triggers

function getTriggersByHandler(handlerName) {
  var triggers = ScriptApp.getProjectTriggers();
  var list = [];
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === handlerName) {
      list.push(triggers[i]);
    }
  }
  return list;
}

function getDailyTriggerHour() {
  var stored = getPropString(PROP_KEYS.DAILY_TRIGGER_HOUR, "");
  var parsed = parseInt(stored, 10);
  return !isNaN(parsed) && parsed >= 0 && parsed <= 23
    ? parsed
    : CONFIG.TRIGGER_HOUR;
}


function getDailyTriggerMinute() {
  var stored = getPropString(PROP_KEYS.DAILY_TRIGGER_MINUTE, "");
  var parsed = parseInt(stored, 10);
  return !isNaN(parsed) && parsed >= 0 && parsed <= 59 ? parsed : 0;
}


function formatTwoDigits(value) {
  var num = parseInt(value, 10);
  if (isNaN(num) || num < 0) num = 0;
  return num < 10 ? "0" + num : String(num);
}


function formatDailyTriggerTime(hour, minute) {
  return hour + ":" + formatTwoDigits(minute);
}


function parseDailyTriggerTimeInput(rawText) {
  var text = String(rawText || "").trim();
  if (!text) return null;

  if (/^\d{1,2}$/.test(text)) {
    var hourOnly = parseInt(text, 10);
    if (hourOnly >= 0 && hourOnly <= 23) {
      return { hour: hourOnly, minute: 0 };
    }
    return null;
  }

  var match = text.match(/^(\d{1,2}):(\d{1,2})$/);
  if (!match) return null;

  var hour = parseInt(match[1], 10);
  var minute = parseInt(match[2], 10);
  if (hour < 0 || hour > 23 || minute < 0 || minute > 59) return null;

  return { hour: hour, minute: minute };
}

function setDailyTriggerHourDialog() {
  var ui = SpreadsheetApp.getUi();
  var current = getDailyTriggerHour();
  var currentMinute = getDailyTriggerMinute();
  var response = ui.prompt(
    "Set Daily Send Time",
    "Enter the daily email schedule time in 24-hour format.\n\nExamples: 8 = 8:00 AM, 8:30 = 8:30 AM, 14:15 = 2:15 PM\n\nCurrent: " +
      formatDailyTriggerTime(current, currentMinute),
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;
  var parsed = parseDailyTriggerTimeInput(response.getResponseText());
  if (!parsed) {
    ui.alert(
      "Invalid time. Please enter a valid 24-hour time like 8, 8:30, or 14:15.",
    );
    return;
  }
  setPropString(PROP_KEYS.DAILY_TRIGGER_HOUR, String(parsed.hour));
  setPropString(PROP_KEYS.DAILY_TRIGGER_MINUTE, String(parsed.minute));
  ui.alert(
    "Send Time Updated",
    "Daily send time set to " +
      formatDailyTriggerTime(parsed.hour, parsed.minute) +
      " Philippine Time.\n\nClick 'Activate Daily Schedule' to apply the new time.",
    ui.ButtonSet.OK,
  );
}


function installTrigger() {
  rememberSpreadsheetId(SpreadsheetApp.getActiveSpreadsheet());
  removeTrigger();

  var hour = getDailyTriggerHour();
  var minute = getDailyTriggerMinute();

  ScriptApp.newTrigger("runDailyCheck")
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .nearMinute(minute)
    .inTimezone("Asia/Manila")
    .create();

  var msg =
    "Daily schedule activated. runDailyCheck() will run automatically every day at " +
    formatDailyTriggerTime(hour, minute) +
    " Philippine Time (Asia/Manila).\n\nNote: Apps Script time-based triggers run in a scheduled window, so the actual fire time may be near that minute rather than exact to the second.";
  Logger.log(msg);
  try {
    SpreadsheetApp.getUi().alert(
      "Schedule Activated",
      msg,
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  } catch (e) {}
}

function installReplyScanTrigger() {
  rememberSpreadsheetId(SpreadsheetApp.getActiveSpreadsheet());
  removeReplyScanTrigger();

  ScriptApp.newTrigger("runReplyScan")
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .inTimezone("Asia/Manila")
    .create();

  ScriptApp.newTrigger("runReplyScan")
    .timeBased()
    .everyDays(1)
    .atHour(15)
    .inTimezone("Asia/Manila")
    .create();

  var msg =
    "Reply scan activated: runs twice daily at 9:00 AM and 3:00 PM Philippine Time.";
  try {
    SpreadsheetApp.getUi().alert(
      "Reply Tracking",
      msg,
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  } catch (e) {}
  Logger.log(msg);
}

function removeReplyScanTrigger() {
  var triggers = getTriggersByHandler("runReplyScan");
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  var msg =
    triggers.length > 0
      ? "Reply scan schedule deactivated. " +
        triggers.length +
        " trigger(s) removed."
      : "No active reply scan schedule found.";
  try {
    SpreadsheetApp.getUi().alert(
      "Reply Tracking",
      msg,
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  } catch (e) {}
  Logger.log(msg);
}

function removeTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var count = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "runDailyCheck") {
      ScriptApp.deleteTrigger(triggers[i]);
      count++;
    }
  }
  var msg =
    count > 0
      ? "Daily schedule deactivated. " + count + " trigger(s) removed."
      : "No active daily schedule found. Nothing to remove.";
  Logger.log(msg);
  try {
    SpreadsheetApp.getUi().alert(
      "Schedule Deactivated",
      msg,
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  } catch (e) {}
}
