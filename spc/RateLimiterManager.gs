/**
 * ═══════════════════════════════════════════════════════════════
 * RATE LIMITER MANAGER
 * Version: 1.0 (Integrated with PCF v3.0)
 * Purpose: Track API quota usage and prevent 429 errors
 * ═══════════════════════════════════════════════════════════════
 */

/**
 * Global Rate Limiter for Gemini API
 * Tracks requests per minute, tokens per minute, and daily requests
 */
const RateLimiterManager = {
  /**
   * Get rate limiter state from PropertiesService (persistent across executions)
   */
  getState() {
    const props = PropertiesService.getScriptProperties();
    const now = Date.now();

    // Get stored state or initialize
    const stateJson = props.getProperty("RATE_LIMITER_STATE");
    let state = stateJson ? JSON.parse(stateJson) : null;

    // Initialize or reset if needed
    if (!state || !state.lastResetDate) {
      state = {
        requestTimes: [],
        tokenEvents: [],
        tokenCount: 0,
        dailyRequestCount: 0,
        lastResetDate: new Date().toDateString(),
        lastRequestTime: 0,
      };
    }

    if (!state.tokenEvents) {
      state.tokenEvents = [];
    }

    // Reset daily counter if new day
    const currentDate = new Date().toDateString();
    if (currentDate !== state.lastResetDate) {
      state.dailyRequestCount = 0;
      state.lastResetDate = currentDate;
      console.log("✓ Daily quota reset");
    }

    // Clean old request times (older than 1 minute)
    const oneMinuteAgo = now - 60000;
    state.requestTimes = state.requestTimes.filter(
      (time) => time > oneMinuteAgo,
    );
    state.tokenEvents = state.tokenEvents.filter(
      (event) => event.time > oneMinuteAgo,
    );
    state.tokenCount = state.tokenEvents.reduce(
      (total, event) => total + (event.tokens || 0),
      0,
    );

    return state;
  },

  /**
   * Save rate limiter state
   */
  saveState(state) {
    const props = PropertiesService.getScriptProperties();
    props.setProperty("RATE_LIMITER_STATE", JSON.stringify(state));
  },

  /**
   * Check if we can make a request without exceeding limits
   * @param {number} estimatedTokens - Estimated tokens for this request
   * @returns {Object} - {canProceed: boolean, waitTime: number, reason: string}
   */
  checkLimit(estimatedTokens = 0) {
    const state = this.getState();
    const now = Date.now();

    // Gemini 2.0 Flash Free Tier limits
    const MAX_RPM = 15; // Requests per minute
    const MAX_TPM = 1000000; // Tokens per minute
    const MAX_RPD = 200; // Requests per day

    // Check daily limit
    if (state.dailyRequestCount >= MAX_RPD) {
      const waitTime = this.getTimeUntilMidnightPST();
      console.log(
        `✗ Daily request limit reached: ${state.dailyRequestCount}/${MAX_RPD}`,
      );
      return {
        canProceed: false,
        waitTime,
        reason: `DAILY_LIMIT_REACHED (${state.dailyRequestCount}/${MAX_RPD})`,
      };
    }

    // Check per-minute request limit
    if (state.requestTimes.length >= MAX_RPM) {
      const oldestRequest = Math.min(...state.requestTimes);
      const waitTime = 60000 - (now - oldestRequest) + 2000; // Add 2s buffer
      console.log(
        `✗ Per-minute request limit: ${state.requestTimes.length}/${MAX_RPM}, wait ${(waitTime / 1000).toFixed(1)}s`,
      );
      return {
        canProceed: false,
        waitTime,
        reason: `RPM_LIMIT (${state.requestTimes.length}/${MAX_RPM})`,
      };
    }

    // Check token limit (approximate)
    if (state.tokenCount + estimatedTokens > MAX_TPM) {
      const waitTime = 60000; // Wait full minute for token reset
      console.log(
        `✗ Token limit would be exceeded: ${state.tokenCount + estimatedTokens}/${MAX_TPM}`,
      );
      return {
        canProceed: false,
        waitTime,
        reason: `TOKEN_LIMIT (${state.tokenCount}/${MAX_TPM})`,
      };
    }

    return { canProceed: true, waitTime: 0, reason: "OK" };
  },

  /**
   * Record a successful request
   * @param {number} tokensUsed - Actual tokens consumed by this request
   */
  recordRequest(tokensUsed = 0) {
    const state = this.getState();
    const now = Date.now();

    state.requestTimes.push(now);
    state.tokenEvents.push({ time: now, tokens: tokensUsed });
    state.tokenCount += tokensUsed;
    state.dailyRequestCount++;
    state.lastRequestTime = now;

    this.saveState(state);

    console.log(
      `📊 Rate Stats: RPM=${state.requestTimes.length}/15, Daily=${state.dailyRequestCount}/200, Tokens≈${state.tokenCount}`,
    );
  },

  /**
   * Calculate milliseconds until midnight PST (when daily quota resets)
   */
  getTimeUntilMidnightPST() {
    const now = new Date();
    const pstOffset = -8 * 60; // PST is UTC-8
    const nowPST = new Date(
      now.getTime() + (now.getTimezoneOffset() + pstOffset) * 60000,
    );

    const midnightPST = new Date(nowPST);
    midnightPST.setHours(24, 0, 0, 0);

    return midnightPST.getTime() - nowPST.getTime();
  },

  /**
   * Reset rate limiter (for testing purposes)
   */
  reset() {
    PropertiesService.getScriptProperties().deleteProperty(
      "RATE_LIMITER_STATE",
    );
    console.log("✓ Rate limiter reset");
  },

  /**
   * Get current usage statistics
   * @returns {Object} Current usage stats
   */
  getUsageStats() {
    const state = this.getState();
    return {
      requestsThisMinute: state.requestTimes.length,
      requestsToday: state.dailyRequestCount,
      tokensThisMinute: state.tokenCount,
      maxRPM: 15,
      maxRPD: 200,
      maxTPM: 1000000,
    };
  },
};

/**
 * Helper function: Sleep with logging
 * @param {number} ms - Milliseconds to sleep
 * @param {string} reason - Reason for sleeping
 */
function rateLimiterSleep(ms, reason = "Rate limiting") {
  if (ms <= 0) return;

  const seconds = (ms / 1000).toFixed(1);
  console.log(`⏳ ${reason}: Waiting ${seconds}s...`);
  Logger.log(`${reason}: Waiting ${seconds}s`, "INFO");
  Utilities.sleep(ms);
  console.log(`✓ Wait complete`);
}

/**
 * Test function - Check current rate limiter status
 */
function testRateLimiter() {
  const stats = RateLimiterManager.getUsageStats();
  const message =
    `📊 Rate Limiter Status:\n\n` +
    `Requests this minute: ${stats.requestsThisMinute}/${stats.maxRPM}\n` +
    `Requests today: ${stats.requestsToday}/${stats.maxRPD}\n` +
    `Tokens this minute: ${stats.tokensThisMinute}/${stats.maxTPM}\n\n` +
    `Status: ${stats.requestsThisMinute < stats.maxRPM && stats.requestsToday < stats.maxRPD ? "✅ Ready" : "⚠️ Approaching limit"}`;

  SpreadsheetApp.getUi().alert(
    "Rate Limiter Status",
    message,
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
  console.log(message);
}

/**
 * Test function - Reset rate limiter
 */
function resetRateLimiter() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Reset Rate Limiter",
    "This will reset all rate limit counters.\n\nContinue?",
    ui.ButtonSet.YES_NO,
  );

  if (response === ui.Button.YES) {
    RateLimiterManager.reset();
    ui.alert("✅ Success", "Rate limiter has been reset.", ui.ButtonSet.OK);
  }
}
