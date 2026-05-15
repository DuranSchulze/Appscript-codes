let text1 = "2025-06-14 2025 06 14 2025 06 14 Compliance [bunao]";
let text2 = "2025-06-14 Compliance 2025-06-14";
let text3 = "**2025-06-14** Compliance";

const regex = /^(\s*(?:\d{4}[-/\s]\d{1,2}[-/\s]\d{1,2}|\d{1,2}[-/\s]\d{1,2}[-/\s]\d{4})\s*)+/i;

function test(cleanedTitle) {
  cleanedTitle = cleanedTitle.replace(/\.[^/.]+$/, "");
  cleanedTitle = cleanedTitle.replace(/^[`"'“”‘’\s]+|[`"'“”‘’\s]+$/g, "");
  cleanedTitle = cleanedTitle.replace(regex, "");
  cleanedTitle = cleanedTitle.replace(/[\\/:*?"<>|]+/g, " ");
  cleanedTitle = cleanedTitle.replace(/[_-]+/g, " ");
  cleanedTitle = cleanedTitle.replace(/\s+/g, " ").trim();
  return cleanedTitle;
}

console.log(test(text1));
console.log(test(text2));
console.log(test(text3));
