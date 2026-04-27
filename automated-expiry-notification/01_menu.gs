var ROOT_MENU_NAME = "🔔 Expiry Notifications";

var ROOT_MENU_SPEC = [
  { type: "item", label: "🚀 Setup This Sheet for Automation", fn: "runSetupWizard" },
  { type: "separator" },
  {
    type: "submenu",
    label: "Status & Overview",
    items: [
      { type: "item", label: "Show System Status", fn: "showSystemStatus" },
      { type: "item", label: "View Last Run Summary", fn: "checkScheduleStatus" },
      { type: "item", label: "Show All Tabs Info", fn: "showAllTabsInfo" },
    ],
  },
  { type: "separator" },
  {
    type: "submenu",
    label: "Tab Management",
    items: [
      { type: "item", label: "Configure Automation Tabs", fn: "configureAutomationSheets" },
      { type: "item", label: "Select Working Tab", fn: "selectWorkingTab" },
      { type: "item", label: "Setup Tab Dropdowns", fn: "setupSheetDropdowns" },
      { type: "item", label: "Map Tab Columns", fn: "mapTabColumns" },
      { type: "item", label: "Set Tab Header Row", fn: "setTabHeaderRowDialog" },
    ],
  },
  { type: "separator" },
  {
    type: "submenu",
    label: "Automation Settings",
    items: [
      { type: "item", label: "Activate Daily Schedule", fn: "installTrigger" },
      { type: "item", label: "Set Send Time", fn: "setDailyTriggerHourDialog" },
      { type: "item", label: "Deactivate Daily Schedule", fn: "removeTrigger" },
      { type: "item", label: "Set Default CC Emails", fn: "setDefaultCcEmailsDialog" },
      { type: "separator" },
      { type: "item", label: "Set Reply Keywords", fn: "setReplyKeywords" },
      {
        type: "item",
        label: "Activate Reply Scan (2x Daily)",
        fn: "installReplyScanTrigger",
      },
      { type: "item", label: "Deactivate Reply Scan", fn: "removeReplyScanTrigger" },
      { type: "separator" },
      { type: "item", label: "Set Open Tracking URL", fn: "setOpenTrackingBaseUrl" },
    ],
  },
  { type: "separator" },
  {
    type: "submenu",
    label: "Run & Diagnostics",
    items: [
      { type: "item", label: "Run Manual Check Now", fn: "manualRunNow" },
      { type: "item", label: "Preview Target Dates", fn: "previewTargetDates" },
      { type: "separator" },
      { type: "item", label: "Inspect Row...", fn: "diagnosticInspectRow" },
      { type: "item", label: "Send Test Email by No....", fn: "diagnosticSendTestRow" },
      { type: "separator" },
      { type: "item", label: "Test Gmail Send", fn: "testGmailSend" },
      { type: "item", label: "Test Drive Access", fn: "testDriveAccess" },
      { type: "item", label: "Test All Connections", fn: "testAllConnections" },
      { type: "separator" },
      { type: "item", label: "Check Column Mappings", fn: "checkColumnMappings" },
      { type: "item", label: "Validate Tab Structure", fn: "validateActiveTabStructure" },
      { type: "item", label: "View Tab Configuration", fn: "viewTabConfiguration" },
      { type: "item", label: "Check Reply Tracking Setup", fn: "checkReplyTrackingSetup" },
      {
        type: "item",
        label: "Preview Fallback Template",
        fn: "previewFallbackTemplateBody",
      },
      { type: "separator" },
      { type: "item", label: "System Diagnostics", fn: "runSystemDiagnostics" },
    ],
  },
  { type: "separator" },
  {
    type: "submenu",
    label: "Help",
    items: [
      {
        type: "item",
        label: "Want to integrate this to another google sheet?",
        fn: "showIntegrationLinkDialog",
      },
      { type: "separator" },
      { type: "item", label: "View Documentation", fn: "viewDocumentation" },
      { type: "item", label: "About", fn: "showAbout" },
    ],
  },
];

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  rememberSpreadsheetId(SpreadsheetApp.getActiveSpreadsheet());
  buildRootMenu_(ui, ROOT_MENU_NAME, ROOT_MENU_SPEC).addToUi();
}

function buildRootMenu_(ui, menuName, spec) {
  var menu = ui.createMenu(menuName);
  applyMenuSpec_(ui, menu, spec);
  return menu;
}

function applyMenuSpec_(ui, menu, spec) {
  for (var i = 0; i < spec.length; i++) {
    var item = spec[i];
    if (!item || !item.type) continue;

    if (item.type === "separator") {
      menu.addSeparator();
    } else if (item.type === "item") {
      menu.addItem(item.label, item.fn);
    } else if (item.type === "submenu") {
      var subMenu = ui.createMenu(item.label);
      applyMenuSpec_(ui, subMenu, item.items || []);
      menu.addSubMenu(subMenu);
    }
  }
}
