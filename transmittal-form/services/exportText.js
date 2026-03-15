"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.formatExportText = void 0;
var ZERO_WIDTH_SPACE = "\u200B";
var LONG_TOKEN_THRESHOLD = 12;
var CHUNK_SIZE = 8;
var SOFT_BREAK_AFTER_PATTERN = /([/_\\|])/g;
var SOFT_BREAK_AROUND_PATTERN = /([\-.,;:])/g;
var splitLongToken = function (token) {
    if (token.length <= LONG_TOKEN_THRESHOLD) {
        return token;
    }
    var parts = [];
    for (var index = 0; index < token.length; index += CHUNK_SIZE) {
        parts.push(token.slice(index, index + CHUNK_SIZE));
    }
    return parts.join(ZERO_WIDTH_SPACE);
};
var formatExportText = function (value) {
    var normalized = String(value || "")
        .replace(/\r\n/g, "\n")
        .replace(/\r/g, "\n")
        .replace(SOFT_BREAK_AFTER_PATTERN, "$1".concat(ZERO_WIDTH_SPACE))
        .replace(SOFT_BREAK_AROUND_PATTERN, "".concat(ZERO_WIDTH_SPACE, "$1").concat(ZERO_WIDTH_SPACE));
    return normalized
        .split(/(\s+)/)
        .map(function (token) { return (/^\s+$/.test(token) ? token : splitLongToken(token)); })
        .join("");
};
exports.formatExportText = formatExportText;
