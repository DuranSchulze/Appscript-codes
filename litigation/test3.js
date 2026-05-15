let originalName = "2025-06-14 2025 06 14 2025 06 14 Compliance [bunao]";

function sanitizePleadingTitle(rawTitle) {
  if (!rawTitle) return "";
  let cleanedTitle = String(rawTitle);
  cleanedTitle = cleanedTitle.replace(/\.[^/.]+$/, "");
  cleanedTitle = cleanedTitle.replace(/^[`"'“”‘’\s]+|[`"'“”‘’\s]+$/g, "");
  
  console.log("Before date strip:", cleanedTitle);
  cleanedTitle = cleanedTitle.replace(
    /^(\s*(?:\d{4}[-/\s]\d{1,2}[-/\s]\d{1,2}|\d{1,2}[-/\s]\d{1,2}[-/\s]\d{4})\s*)+/i,
    "",
  );
  console.log("After date strip:", cleanedTitle);
  
  cleanedTitle = cleanedTitle.replace(/[\\/:*?"<>|]+/g, " ");
  cleanedTitle = cleanedTitle.replace(/[_-]+/g, " ");
  cleanedTitle = cleanedTitle.replace(/\s+/g, " ").trim();
  return cleanedTitle;
}

console.log(sanitizePleadingTitle(originalName));
