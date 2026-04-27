// 90 Shared Utils


function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || "").trim());
}


function normalizeEmailAddress(value) {
  var email = String(value || "").trim();
  if (!email) return "";

  var angleMatch = email.match(/<([^<>]+)>/);
  if (angleMatch && angleMatch[1]) {
    email = String(angleMatch[1]).trim();
  }

  return email.replace(/^[\s"'`<]+|[\s"'`>]+$/g, "").trim();
}


function normalizeEmailList(value) {
  var rawValues = [];

  if (Array.isArray(value)) {
    rawValues = value;
  } else {
    var text = String(value || "");
    if (!text.trim()) return [];
    rawValues = text.split(/[,;\n]/);
  }

  var seen = {};
  var result = [];

  for (var i = 0; i < rawValues.length; i++) {
    var email = normalizeEmailAddress(rawValues[i]);
    if (!email) continue;

    var key = email.toLowerCase();
    if (seen[key]) continue;
    seen[key] = true;
    result.push(email);
  }

  return result;
}


function parseClientEmails(rawValue) {
  var list = normalizeEmailList(rawValue);
  return list.length > 0 ? list : [];
}


function validateEmailList(emails) {
  var invalid = [];
  for (var i = 0; i < emails.length; i++) {
    if (!isValidEmail(emails[i])) invalid.push(emails[i]);
  }
  return invalid;
}


function mergeUniqueEmails() {
  var merged = [];
  var seen = {};

  for (var i = 0; i < arguments.length; i++) {
    var source = normalizeEmailList(arguments[i]);
    for (var j = 0; j < source.length; j++) {
      var email = source[j];
      var key = email.toLowerCase();
      if (seen[key]) continue;
      seen[key] = true;
      merged.push(email);
    }
  }

  return merged;
}

function getMidnight(date) {
  var d = new Date(date);
  d.setHours(0, 0, 0, 0);
  return d;
}

function isSameDay(dateA, dateB) {
  var a = getMidnight(dateA);
  var b = getMidnight(dateB);
  return (
    a.getFullYear() === b.getFullYear() &&
    a.getMonth() === b.getMonth() &&
    a.getDate() === b.getDate()
  );
}

function isTargetDateDue(targetDate, referenceDate) {
  return (
    getMidnight(targetDate).getTime() <= getMidnight(referenceDate).getTime()
  );
}


function getSupportedNoticeDateHint() {
  return (
    'Supported formats: "N days/weeks/months/years before", ' +
    '"N days/weeks/months/years", numeric day count (e.g. "7"), ' +
    '"On expiry date", or an explicit date (e.g. "2026-12-31").'
  );
}


function parseIsoDateString(value) {
  var text = String(value || "").trim();
  var match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!match) return null;

  var year = parseInt(match[1], 10);
  var month = parseInt(match[2], 10);
  var day = parseInt(match[3], 10);
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;

  var parsed = new Date(year, month - 1, day);
  if (
    parsed.getFullYear() !== year ||
    parsed.getMonth() !== month - 1 ||
    parsed.getDate() !== day
  ) {
    return null;
  }

  return parsed;
}

function parseNoticeOffset(noticeStr) {
  var s = String(noticeStr || "")
    .toLowerCase()
    .trim();
  if (!s) return null;

  if (/^on(\s+the)?\s+expiry\s+date$/.test(s)) {
    return { unit: "days", value: 0 };
  }

  if (/^\d+$/.test(s)) {
    return { unit: "days", value: parseInt(s, 10) };
  }

  var relative = s.match(
    /^(\d+)\s*(day|days|week|weeks|month|months|year|years|yr|yrs)(?:\s+before)?$/,
  );
  if (relative) {
    var value = parseInt(relative[1], 10);
    var unitRaw = relative[2];

    if (unitRaw === "week" || unitRaw === "weeks") {
      return { unit: "days", value: value * 7 };
    }
    if (unitRaw === "month" || unitRaw === "months") {
      return { unit: "months", value: value };
    }
    if (
      unitRaw === "year" ||
      unitRaw === "years" ||
      unitRaw === "yr" ||
      unitRaw === "yrs"
    ) {
      return { unit: "years", value: value };
    }

    return { unit: "days", value: value };
  }

  var isoDate = parseIsoDateString(s);
  if (isoDate) {
    return { unit: "absolute_date", value: getMidnight(isoDate) };
  }

  var directDate = new Date(s);
  if (!isNaN(directDate.getTime())) {
    return { unit: "absolute_date", value: getMidnight(directDate) };
  }

  return null;
}

function computeTargetDate(expiryDate, offset) {
  if (offset.unit === "absolute_date") {
    return getMidnight(offset.value);
  }

  var target = new Date(expiryDate);
  if (offset.unit === "days") {
    target.setDate(target.getDate() - offset.value);
  } else if (offset.unit === "months") {
    target.setMonth(target.getMonth() - offset.value);
  } else if (offset.unit === "years") {
    target.setFullYear(target.getFullYear() - offset.value);
  }
  return getMidnight(target);
}

function formatDate(date) {
  var months = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
  ];
  return (
    date.getDate() + " " + months[date.getMonth()] + " " + date.getFullYear()
  );
}
