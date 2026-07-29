/**
 * Version: 3.0 (Production Ready)
 * Data Validator
 * Validates and cleans extracted data
 */

const DataValidator = {
  /**
   * Validate and clean extracted data
   */
  validateAndClean(rawData) {
    const cleanData = {
      name: this.validateName(rawData.name),
      errandDate: this.validateDate(rawData.errandDate),
      voucherNo: this.validateVoucherNo(rawData.voucherNo),
      company: this.validateText(rawData.company, "Company"),
      errandBy: this.validateText(rawData.errandBy, "Staff"),
      service: this.validateText(rawData.service, "Service"),
      details: this.validateText(rawData.details, "Details"),
      expenseClassification: this.validateText(
        rawData.expenseClassification,
        "Expense Classification",
      ),
      total: this.validateAmount(rawData.total),
      mainLocation: this.validateText(rawData.mainLocation, "Location"),
    };

    return cleanData;
  },
  /**
   * Validate name field
   */
  validateName(name) {
    if (!name || typeof name !== "string") {
      return "";
    }
    return name.trim().substring(0, 100);
  },
  /**
   * Validate date field
   */
  validateDate(dateStr) {
    if (!dateStr) return "";

    const value = String(dateStr).trim();
    let match;

    // Treat voucher dates as calendar dates, not moments in time. Constructing
    // `new Date("YYYY-MM-DD")` interprets the value through a timezone and can
    // move it to the previous day in some Apps Script/spreadsheet settings.
    match = value.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:\D.*)?$/);
    if (match) {
      return this.formatDateParts(match[1], match[2], match[3]);
    }

    match = value.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
    if (match) {
      const first = Number(match[1]);
      const second = Number(match[2]);

      // Use the unambiguous component when possible. For dates such as
      // 07/29/2026, this is MM/DD/YYYY. Ambiguous numeric dates follow the
      // spreadsheet's expected US-style entry format.
      const month = first > 12 ? second : first;
      const day = first > 12 ? first : second;
      return this.formatDateParts(match[3], month, day);
    }

    return "";
  },
  /**
   * Validate date components and return a timezone-free ISO calendar date.
   */
  formatDateParts(yearValue, monthValue, dayValue) {
    const year = Number(yearValue);
    const month = Number(monthValue);
    const day = Number(dayValue);

    if (
      !Number.isInteger(year) ||
      !Number.isInteger(month) ||
      !Number.isInteger(day) ||
      year < 1900 ||
      year > 2100 ||
      month < 1 ||
      month > 12 ||
      day < 1
    ) {
      return "";
    }

    const daysInMonth = new Date(Date.UTC(year, month, 0)).getUTCDate();
    if (day > daysInMonth) {
      return "";
    }

    return `${String(year).padStart(4, "0")}-${String(month).padStart(
      2,
      "0",
    )}-${String(day).padStart(2, "0")}`;
  },
  /**
   * Validate voucher number format
   */
  validateVoucherNo(voucherNo) {
    if (!voucherNo || typeof voucherNo !== "string") {
      return "";
    }

    const cleaned = voucherNo.trim();

    // Check if it matches expected format: YYYY-number-letter
    const voucherPattern = /^\d{4}-\d+-[A-Za-z]$/;
    if (voucherPattern.test(cleaned)) {
      return cleaned;
    }

    // Return as-is if doesn't match expected format
    return cleaned.substring(0, 50);
  },
  /**
   * Validate text fields
   */
  validateText(text, fieldName) {
    if (!text || typeof text !== "string") {
      return "";
    }

    const cleaned = text.trim();

    // Limit length based on field
    const maxLength = this.getMaxLength(fieldName);
    return cleaned.substring(0, maxLength);
  },
  /**
   * Get maximum length for different fields
   */
  getMaxLength(fieldName) {
    const lengths = {
      Company: 200,
      Staff: 100,
      Service: 150,
      Details: 2000,
      "Expense Classification": 100,
      Location: 200,
    };

    return lengths[fieldName] || 100;
  },
  /**
   * Validate amount field
   */
  validateAmount(amount) {
    if (!amount) return 0;

    // If string, try to parse
    if (typeof amount === "string") {
      // Remove currency symbols and commas
      const cleaned = amount.replace(/[₱$,\s]/g, "");
      const parsed = parseFloat(cleaned);
      return isNaN(parsed) ? 0 : Math.abs(parsed);
    }

    // If number, validate
    if (typeof amount === "number") {
      return isNaN(amount) ? 0 : Math.abs(amount);
    }

    return 0;
  },
};
