// 20 Properties

function getAutomationProperties() {
  return PropertiesService.getDocumentProperties();
}

function rememberSpreadsheetId(ss) {
  try {
    if (!ss) return;
    var id = ss.getId();
    if (id) setPropString(PROP_KEYS.SPREADSHEET_ID, id);
  } catch (e) {}
}

function getPropString(key, fallbackValue) {
  var value = getAutomationProperties().getProperty(key);
  if (value === null || value === undefined || String(value).trim() === "") {
    return fallbackValue || "";
  }
  return String(value).trim();
}

function setPropString(key, value) {
  var props = getAutomationProperties();
  var text = String(value === null || value === undefined ? "" : value).trim();
  if (!text) {
    props.deleteProperty(key);
    return;
  }
  props.setProperty(key, text);
}

function getPropBoolean(key, fallbackValue) {
  var raw = getAutomationProperties().getProperty(key);
  if (raw === null || raw === undefined || String(raw).trim() === "") {
    return !!fallbackValue;
  }
  return String(raw).toLowerCase().trim() === "true";
}

function setPropBoolean(key, value) {
  getAutomationProperties().setProperty(key, value ? "true" : "false");
}
