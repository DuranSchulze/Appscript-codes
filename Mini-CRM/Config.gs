/**
 * AI & Automation Configuration
 */
const AI_CONFIG = {
  // Default model – will be overwritten by script property
  DEFAULT_GEMINI_MODEL: 'gemini-1.5-pro',
  // Gemini REST API base URL
  GEMINI_API_BASE_URL: 'https://generativelanguage.googleapis.com/v1beta',
  // FAQ matching confidence threshold (0-100)
  FAQ_CONFIDENCE_THRESHOLD: 95,
  // AI sales qualification thresholds (0-100)
  QUALIFIED_LEAD_THRESHOLD: 75,
  MANUAL_REVIEW_THRESHOLD: 45,
  // Shared Gemini safeguards. Script Properties with matching names can
  // override these conservative defaults without editing source code.
  MAX_AI_REQUESTS_PER_MINUTE: 4,
  MAX_AI_REQUESTS_PER_DAY: 150,
  AI_QUOTA_TIMEZONE: 'America/Los_Angeles',
  MAX_AI_QUALIFICATION_REQUESTS_PER_DAY: 100,
  MAX_AI_AUTO_REPLY_REQUESTS_PER_DAY: 40,
  MAX_AI_CONFIGURATION_REQUESTS_PER_DAY: 10,
  MAX_AUTO_DRAFTS_PER_RUN: 3,
  MAX_AI_PENDING_THREADS_PER_RUN: 20,
  MAX_AI_DRAFT_BODY_CHARS: 8000,
  // Limit synchronous classifications to protect Apps Script runtime/quota.
  MAX_QUALIFICATIONS_PER_SYNC: 10,
  MAX_PENDING_QUALIFICATIONS_PER_RUN: 10,
  MAX_QUALIFICATION_BODY_CHARS: 6000,
  // Follow-up delay in days
  FOLLOW_UP_DELAY_DAYS: 3,
  // Gmail label to watch for AI processing
  AI_PENDING_LABEL: 'AI-Pending',
  // FAQ Sheet name
  FAQ_SHEET_NAME: 'FAQ Responses',
  // TidyCal consultation links
  CONSULTATION_LINKS: [
    { lawyer: 'Marie Christine', url: 'https://tidycal.com/mariechristine' },
    { lawyer: 'Atty. Mary Wendy', url: 'https://tidycal.com/attymarywendy' }
  ]
};

/**
 * Get Gemini API key from script properties (set by user)
 */
function getGeminiApiKey() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
}

/**
 * Get selected Gemini model from script properties
 */
function getGeminiModel() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || AI_CONFIG.DEFAULT_GEMINI_MODEL;
}

/**
 * Get the optional fallback Gemini model. An empty value disables fallback.
 */
function getGeminiFallbackModel() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_FALLBACK_MODEL') || '';
}

/**
 * Normalizes values returned by the Gemini models endpoint (for example,
 * "models/gemini-2.5-flash") into the ID used by generateContent.
 */
function normalizeGeminiModelId(modelName) {
  return String(modelName || '').trim().replace(/^models\//, '');
}
