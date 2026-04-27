// System Version: April 27, 2026

var CONFIG = {
  SHEET_NAME: "VISA automation",
  LOGS_SHEET_NAME: "LOGS",
  AUTOMATION_SHEET_PROPERTY_KEY: "AUTOMATION_SHEET_NAMES",
  HEADER_ROW: 2,
  DATA_START_ROW: 3,
  TRIGGER_HOUR: 8,
  REPLY_SCAN_TRIGGER_HOUR: 9,
  SENDER_NAME: "Office",
  STATIC_REDIRECT_URL: "https://pastebin.com/8n85J6k6",
};

var HEADERS = {
  NO: "No.",
  CLIENT_NAME: "Client Name",
  CLIENT_EMAIL: "Client Email",
  DOC_TYPE: "Type of ID/Document",
  EXPIRY_DATE: "Expiry Date",
  NOTICE_DATE: "Notice Date",
  REMARKS: "Remarks",
  ATTACHMENTS: "Attached Files",
  STATUS: "Status",
  STAFF_EMAIL: "Staff Email",
  SEND_MODE: "Send Mode",
  SENT_AT: "Sent At",
  SENT_THREAD_ID: "Sent Thread Id",
  SENT_MESSAGE_ID: "Sent Message Id",
  REPLY_STATUS: "Reply Status",
  REPLIED_AT: "Replied At",
  REPLY_KEYWORD: "Reply Keyword",
  OPEN_TOKEN: "Open Tracking Token",
  FIRST_OPENED_AT: "First Opened At",
  LAST_OPENED_AT: "Last Opened At",
  OPEN_COUNT: "Open Count",
  FINAL_NOTICE_SENT_AT: "Final Notice Sent At",
  FINAL_NOTICE_THREAD_ID: "Final Notice Thread Id",
  FINAL_NOTICE_MESSAGE_ID: "Final Notice Message Id",
};

var HEADER_ALIASES = {
  CLIENT_NAME: ["Seller", "Seller Name", "Buyer", "Buyer Name"],
  CLIENT_EMAIL: [
    "Seller Email",
    "Buyer Email",
    "Seller E-mail",
    "Buyer E-mail",
    "Seller Mail",
    "Buyer Mail",
  ],
  DOC_TYPE: ["Services", "Service"],
  EXPIRY_DATE: ["Due Date"],
  NOTICE_DATE: ["Remaining Days", "Reminder Days"],
  REMARKS: [
    "Reminder (Email Content)",
    "Reminder Email Content",
    "Reminder Content",
  ],
  ATTACHMENTS: [
    "Attached File",
    "Gsheet",
    "GSheet",
    "Google Sheet",
    "Google Sheets",
  ],
  STATUS: ["Project Status"],
  STAFF_EMAIL: ["Assigned Staff Email"],
  SENT_THREAD_ID: ["Sent Thread ID"],
  SENT_MESSAGE_ID: ["Sent Message ID"],
  SEND_MODE: ["Send Option", "Mode"],
  OPEN_TOKEN: ["Open Token"],
  FIRST_OPENED_AT: ["First Open At"],
  LAST_OPENED_AT: ["Last Open At"],
  FINAL_NOTICE_SENT_AT: ["Final Notice Date", "Final Sent At"],
  FINAL_NOTICE_THREAD_ID: ["Final Notice Thread ID", "Final Thread ID"],
  FINAL_NOTICE_MESSAGE_ID: ["Final Notice Message ID", "Final Message ID"],
};

var STATUS = {
  ACTIVE: "Active",
  NOTICE_SENT: "Notice Sent",
  SENT: "Sent",
  ERROR: "Error",
  SKIPPED: "Skipped",
};

var SEND_MODE = {
  AUTO: "Auto",
  HOLD: "Hold",
  MANUAL_ONLY: "Manual Only",
};

var REPLY_STATUS = {
  PENDING: "Pending",
  REPLIED: "Replied",
};

var PROP_KEYS = {
  SPREADSHEET_ID: "SPREADSHEET_ID",
  REPLY_KEYWORDS: "REPLY_KEYWORDS",
  AI_ENABLED: "AI_ENABLED",
  AI_PROVIDER: "AI_PROVIDER",
  AI_API_KEY: "AI_API_KEY",
  AI_MODEL: "AI_MODEL",
  FALLBACK_TEMPLATE_MODE: "FALLBACK_TEMPLATE_MODE",
  FALLBACK_TEMPLATE: "FALLBACK_TEMPLATE",
  OPEN_TRACKING_BASE_URL: "OPEN_TRACKING_BASE_URL",
  DEFAULT_CC_EMAILS: "DEFAULT_CC_EMAILS",
  DAILY_TRIGGER_HOUR: "DAILY_TRIGGER_HOUR",
  DAILY_TRIGGER_MINUTE: "DAILY_TRIGGER_MINUTE",
};

var AI_PROVIDER = {
  GEMINI: "gemini",
};

var FALLBACK_TEMPLATE_MODE = {
  HARDCODED: "HARDCODED",
  PROPERTY: "PROPERTY",
};

var DEFAULT_REPLY_KEYWORDS = ["ACK", "RECEIVED", "OK"];
var DEFAULT_AI_MODEL = "models/gemini-1.5-flash";
var DEFAULT_NOTICE_OPTIONS = [
  "7 days before",
  "14 days before",
  "30 days before",
  "60 days before",
  "90 days before",
  "1 year before",
  "2 years before",
  "On expiry date",
];
var PROP_KEY_NOTICE_OPTIONS_PREFIX = "NOTICE_OPTIONS_";

var LOG_COL = {
  TIMESTAMP: 1,
  TAB: 2,
  CLIENT_NAME: 3,
  ACTION: 4,
  DETAIL: 5,
};

var TAB_CONFIG = {
  PREFIX: "TAB_CONFIG_",
  DEFAULT_HEADER_ROW: 2,
};

var TAB_CONFIG_KEYS = {
  COLUMN_MAP: "COLUMN_MAP",
  HEADER_ROW: "HEADER_ROW",
  DATA_START_ROW: "DATA_START_ROW",
  NOTICE_OPTIONS: "NOTICE_OPTIONS",
  STATUS_OPTIONS: "STATUS_OPTIONS",
  SEND_MODE_OPTIONS: "SEND_MODE_OPTIONS",
  LAST_SELECTED: "LAST_SELECTED",
};

var FLEXIBLE_HEADER_ALIASES = {
  NO: [
    "No.",
    "No",
    "Number",
    "#",
    "ID",
    "Ref",
    "Reference",
    "Ref No",
    "Ref. No.",
    "Reference No",
  ],
  CLIENT_NAME: [
    "Client Name",
    "Name",
    "Full Name",
    "Client",
    "Applicant Name",
    "Applicant",
    "Person",
    "Contact Name",
    "Seller",
    "Seller Name",
    "Buyer",
    "Buyer Name",
  ],
  CLIENT_EMAIL: [
    "Client Email",
    "Email",
    "Email Address",
    "E-mail",
    "E-mail Address",
    "Contact Email",
    "Mail",
    "Seller Email",
    "Buyer Email",
    "Seller E-mail",
    "Buyer E-mail",
    "Seller Mail",
    "Buyer Mail",
  ],
  DOC_TYPE: [
    "Type of ID/Document",
    "Document Type",
    "Doc Type",
    "Type",
    "ID Type",
    "Visa Type",
    "Permit Type",
    "Document",
    "Services",
    "Service",
  ],
  EXPIRY_DATE: [
    "Expiry Date",
    "Expiration Date",
    "Expires On",
    "Valid Until",
    "End Date",
    "Date of Expiration",
    "Expiry",
    "Due Date",
  ],
  NOTICE_DATE: [
    "Notice Date",
    "Notice",
    "Reminder Date",
    "Send On",
    "Notify On",
    "Advance Notice",
    "Remaining Days",
    "Reminder Days",
  ],
  REMARKS: [
    "Remarks",
    "Notes",
    "Comments",
    "Message",
    "Body",
    "Email Body",
    "Note",
    "Reminder (Email Content)",
    "Reminder Email Content",
    "Reminder Content",
  ],
  ATTACHMENTS: [
    "Attached Files",
    "Attachments",
    "Files",
    "Docs",
    "Documents",
    "Attached",
    "Drive Links",
    "Attached File",
    "Gsheet",
    "GSheet",
    "Google Sheet",
    "Google Sheets",
  ],
  STATUS: [
    "Status",
    "State",
    "Send Status",
    "Processing Status",
    "Project Status",
  ],
  STAFF_EMAIL: [
    "Staff Email",
    "Assigned Staff Email",
    "Staff",
    "Handler Email",
    "Assigned To",
    "Owner Email",
  ],
  SEND_MODE: [
    "Send Mode",
    "Send Option",
    "Mode",
    "Send",
    "Auto Send",
    "Processing Mode",
  ],
  SENT_AT: ["Sent At", "Date Sent", "Sent On", "Processed At", "Last Sent"],
  SENT_THREAD_ID: [
    "Sent Thread Id",
    "Sent Thread ID",
    "Thread ID",
    "Gmail Thread",
    "Thread",
  ],
  SENT_MESSAGE_ID: [
    "Sent Message Id",
    "Sent Message ID",
    "Message ID",
    "Gmail Message",
  ],
  REPLY_STATUS: ["Reply Status", "Response Status", "Acknowledged", "Replied"],
  REPLIED_AT: ["Replied At", "Reply Date", "Response Date", "Replied On"],
  REPLY_KEYWORD: ["Reply Keyword", "Keyword", "Ack Keyword"],
  OPEN_TOKEN: ["Open Tracking Token", "Open Token", "Tracking Token", "Token"],
  FIRST_OPENED_AT: ["First Opened At", "First Open At", "First Viewed"],
  LAST_OPENED_AT: ["Last Opened At", "Last Open At", "Last Viewed"],
  OPEN_COUNT: ["Open Count", "View Count", "Times Opened"],
  FINAL_NOTICE_SENT_AT: [
    "Final Notice Sent At",
    "Final Notice Date",
    "Final Sent At",
  ],
  FINAL_NOTICE_THREAD_ID: [
    "Final Notice Thread Id",
    "Final Notice Thread ID",
    "Final Thread ID",
  ],
  FINAL_NOTICE_MESSAGE_ID: [
    "Final Notice Message Id",
    "Final Notice Message ID",
    "Final Message ID",
  ],
};
