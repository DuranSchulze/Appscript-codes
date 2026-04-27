# Automated Expiry Notification

Behavior-safe Apps Script modularization of the spreadsheet automation for expiry reminders, reply tracking, and open tracking.

## Architecture Map

- [00_config.gs](/Users/zafajardo/Documents/Development/dds-all-system/automated-expiry-notification/00_config.gs): constants, enums, property keys, header aliases, defaults
- [01_menu.gs](/Users/zafajardo/Documents/Development/dds-all-system/automated-expiry-notification/01_menu.gs): `onOpen`, menu structure, menu builder helpers
- [10_setup_wizard.gs](/Users/zafajardo/Documents/Development/dds-all-system/automated-expiry-notification/10_setup_wizard.gs): guided setup flow
- [11_status_ui.gs](/Users/zafajardo/Documents/Development/dds-all-system/automated-expiry-notification/11_status_ui.gs): status dialogs, docs, about, integration link UI
- [12_tab_management.gs](/Users/zafajardo/Documents/Development/dds-all-system/automated-expiry-notification/12_tab_management.gs): configured tab selection and working-tab helpers
- [13_tab_mapping.gs](/Users/zafajardo/Documents/Development/dds-all-system/automated-expiry-notification/13_tab_mapping.gs): mapping UI, header row selection, mapping diagnostics
- [14_dropdowns.gs](/Users/zafajardo/Documents/Development/dds-all-system/automated-expiry-notification/14_dropdowns.gs): status/send-mode/notice-date dropdown setup
- [20_properties.gs](/Users/zafajardo/Documents/Development/dds-all-system/automated-expiry-notification/20_properties.gs): document property access and typed wrappers
- [21_sheet_resolution.gs](/Users/zafajardo/Documents/Development/dds-all-system/automated-expiry-notification/21_sheet_resolution.gs): configured sheet storage and spreadsheet resolution
- [22_column_mapping_core.gs](/Users/zafajardo/Documents/Development/dds-all-system/automated-expiry-notification/22_column_mapping_core.gs): column-map persistence, alias matching, fuzzy detection
- [23_sheet_row_ops.gs](/Users/zafajardo/Documents/Development/dds-all-system/automated-expiry-notification/23_sheet_row_ops.gs): row lookups, status writes, metadata writes, ensure-column helpers
- [30_run_daily_check.gs](/Users/zafajardo/Documents/Development/dds-all-system/automated-expiry-notification/30_run_daily_check.gs): manual run entrypoint and daily orchestration
- [31_notification_rules.gs](/Users/zafajardo/Documents/Development/dds-all-system/automated-expiry-notification/31_notification_rules.gs): send-mode rule helpers
- [32_email_content.gs](/Users/zafajardo/Documents/Development/dds-all-system/automated-expiry-notification/32_email_content.gs): subject/body composition, template fallback, token replacement
- [33_email_delivery.gs](/Users/zafajardo/Documents/Development/dds-all-system/automated-expiry-notification/33_email_delivery.gs): Gmail delivery, sender resolution, CC resolution
- [34_attachments.gs](/Users/zafajardo/Documents/Development/dds-all-system/automated-expiry-notification/34_attachments.gs): Drive attachment parsing and fallback link handling
- [35_ai_generation.gs](/Users/zafajardo/Documents/Development/dds-all-system/automated-expiry-notification/35_ai_generation.gs): Gemini settings, fetch, diagnostics, fallback template controls
- [40_reply_tracking.gs](/Users/zafajardo/Documents/Development/dds-all-system/automated-expiry-notification/40_reply_tracking.gs): reply scan orchestration and reply matching
- [41_open_tracking.gs](/Users/zafajardo/Documents/Development/dds-all-system/automated-expiry-notification/41_open_tracking.gs): tracking URL config, `doGet`, token lookup, open writes
- [42_triggers.gs](/Users/zafajardo/Documents/Development/dds-all-system/automated-expiry-notification/42_triggers.gs): trigger install/remove and schedule parsing
- [43_logs.gs](/Users/zafajardo/Documents/Development/dds-all-system/automated-expiry-notification/43_logs.gs): logs sheet bootstrap and append helpers
- [44_diagnostics.gs](/Users/zafajardo/Documents/Development/dds-all-system/automated-expiry-notification/44_diagnostics.gs): connectivity tests, previews, row inspection, diagnostics actions
- [90_shared_utils.gs](/Users/zafajardo/Documents/Development/dds-all-system/automated-expiry-notification/90_shared_utils.gs): cross-domain pure helpers for parsing, normalization, and dates
- [code.gs](/Users/zafajardo/Documents/Development/dds-all-system/automated-expiry-notification/code.gs): placeholder only
- [appsscript.json](/Users/zafajardo/Documents/Development/dds-all-system/automated-expiry-notification/appsscript.json): Apps Script manifest

## Stable Entrypoints

These public functions remain available for menus, triggers, or web app deployment:

- `onOpen`
- `runDailyCheck`
- `manualRunNow`
- `runReplyScan`
- `doGet`
- all existing menu-exposed setup and diagnostics functions referenced in `01_menu.gs`

## Dependency Direction

- Config: `00_config.gs`
- Shared pure utilities: `90_shared_utils.gs`
- Services and persistence: `20_*`, `21_*`, `22_*`, `23_*`, `33_*`, `34_*`, `43_*`
- Domain logic: `30_*`, `31_*`, `32_*`, `35_*`, `40_*`, `41_*`, `42_*`
- UI and entrypoints: `01_*`, `10_*`, `11_*`, `12_*`, `13_*`, `14_*`, `44_*`

Diagnostics may depend on any module. Other modules should not depend on diagnostics.

## Main Flows

### Daily Run

1. `runDailyCheck` resolves the automation spreadsheet and configured tabs.
2. Each tab loads its header/data row settings and effective column map.
3. Each row is checked for status, send mode, required fields, target date, and reply state.
4. Notice or final-stage content is composed, delivered, and logged.
5. Row metadata and logs are updated through centralized sheet helpers.

### Reply Scan

1. `runReplyScan` resolves configured tabs and required reply-tracking columns.
2. Sent thread metadata is read from the sheet.
3. Gmail threads are scanned for matching reply keywords.
4. Reply status columns and logs are updated when matches are found.

### Tab Setup

1. `runSetupWizard` handles tab selection or creation.
2. Column detection and optional manual mapping run through the mapping core.
3. Automation-managed columns are ensured by row ops.
4. Dropdowns and schedule setup are applied through their dedicated modules.

### Open Tracking

1. Email content injects an open-tracking pixel when a tracking base URL exists.
2. `doGet(e)` handles `open` and `click` modes.
3. Token lookup resolves the matching row across configured tabs.
4. Open counters, timestamps, reply markers, and logs are updated centrally.

## Developer Conventions

- Keep menu labels and public entrypoint names stable unless compatibility is intentionally changed.
- Route property access through `20_properties.gs`.
- Route sheet mutations and ensure-column behavior through `23_sheet_row_ops.gs`.
- Keep Gmail, Drive, and UrlFetch integrations isolated from rule decisions when touching behavior.
- Prefer adding new pure helpers to `90_shared_utils.gs` only when they are truly cross-domain.
