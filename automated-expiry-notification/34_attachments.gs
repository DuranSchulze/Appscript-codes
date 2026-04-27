// 34 Attachments

function extractDriveFileId(urlOrId) {
  if (!urlOrId) return null;
  var s = urlOrId.trim();

  // /file/d/FILE_ID/
  var m = s.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];

  // ?id=FILE_ID or &id=FILE_ID
  m = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) return m[1];

  // Assume raw ID if it looks like one (alphanumeric + _ -)
  if (/^[a-zA-Z0-9_-]{20,}$/.test(s)) return s;

  return null;
}

function splitAttachmentEntries(rawField) {
  if (!rawField || rawField.trim() === "") return [];
  // Normalize: treat newlines as delimiters alongside commas
  var normalized = String(rawField).replace(/[\r\n]+/g, ",");
  return normalized
    .split(",")
    .map(function (e) {
      return e.trim();
    })
    .filter(function (e) {
      return !!e;
    });
}

function resolveAttachments(rawField) {
  if (!rawField || rawField.trim() === "") {
    return { blobs: [], warnings: [], failedLinks: [], fatalError: null };
  }

  var entries = splitAttachmentEntries(rawField);
  if (entries.length === 0) {
    return { blobs: [], warnings: [], failedLinks: [], fatalError: null };
  }

  var blobs = [];
  var warnings = [];
  var failedLinks = [];

  for (var i = 0; i < entries.length; i++) {
    var entry = entries[i];
    var fileId = extractDriveFileId(entry);

    if (!fileId) {
      warnings.push('Cannot parse Drive file ID from: "' + entry + '"');
      failedLinks.push({ label: entry, url: null });
      continue;
    }

    try {
      var file = DriveApp.getFileById(fileId);
      blobs.push({
        blob: file.getBlob(),
        name: file.getName(),
        fileId: fileId,
      });
    } catch (e) {
      warnings.push(
        "File not found or inaccessible (ID: " + fileId + "): " + e.message,
      );
      var originalUrl =
        entry.indexOf("drive.google.com") >= 0
          ? entry
          : "https://drive.google.com/file/d/" + fileId + "/view";
      failedLinks.push({ label: fileId, url: originalUrl });
    }
  }

  return {
    blobs: blobs,
    warnings: warnings,
    failedLinks: failedLinks,
    fatalError: null,
  };
}


function buildFallbackLinksHtml(failedLinks) {
  if (!failedLinks || failedLinks.length === 0) return "";

  var items = [];
  for (var i = 0; i < failedLinks.length; i++) {
    var fl = failedLinks[i];
    if (fl.url) {
      items.push(
        '<li><a href="' +
          sanitizeHtmlAttribute(fl.url) +
          '" target="_blank" rel="noopener noreferrer">' +
          sanitizeHtmlContent(fl.label) +
          "</a></li>",
      );
    } else {
      items.push("<li>" + sanitizeHtmlContent(fl.label) + "</li>");
    }
  }

  return (
    '<div style="margin-top:16px;padding:12px;background:#fff8e1;border-left:3px solid #f9a825;font-size:13px;">' +
    '<p style="margin:0 0 8px 0;font-weight:bold;color:#7a5c00;">Some files could not be attached. You can access them via the links below:</p>' +
    '<ul style="margin:0;padding-left:20px;">' +
    items.join("") +
    "</ul></div>"
  );
}
