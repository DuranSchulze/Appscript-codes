function archiveOldSheetsAndTrim() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const currentYear = getCurrentArchiveYear_();
  const CHUNK_SIZE = 500; // rows processed per batch

  // --- Step 1: Create or open archive spreadsheet ---
  const archiveName = `CRM Archive (pre-${currentYear})`;
  const archiveFiles = DriveApp.getFilesByName(archiveName);
  let archive;
  if (archiveFiles.hasNext()) {
    archive = SpreadsheetApp.openById(archiveFiles.next().getId());
  } else {
    archive = SpreadsheetApp.create(archiveName);
  }

  // Handle the default "Sheet1" – never delete if it's the only sheet
  const allArchiveSheets = archive.getSheets();
  if (allArchiveSheets.length === 1 && allArchiveSheets[0].getName() === 'Sheet1') {
    allArchiveSheets[0].setName('Archive Placeholder');
  } else {
    const defaultSheet = archive.getSheetByName('Sheet1');
    if (defaultSheet) archive.deleteSheet(defaultSheet);
  }

  // --- Step 2: Move monthly sheets from years before the current year ---
  const monthlyPattern = /^([A-Za-z]{3})-(\d{4})$/;
  let movedCount = 0;
  const sheetsToMove = [];
  for (const sheet of sheets) {
    const name = sheet.getName();
    const match = name.match(monthlyPattern);
    if (match) {
      const year = parseInt(match[2]);
      if (year < currentYear) {
        sheetsToMove.push(sheet);
      }
    }
  }

  for (const sheet of sheetsToMove) {
    const sheetName = sheet.getName();
    sheet.copyTo(archive).setName(sheetName);
    ss.deleteSheet(sheet);
    movedCount++;
  }

  // Remove placeholder if real sheets were moved
  const finalArchiveSheets = archive.getSheets();
  if (finalArchiveSheets.length > 1) {
    const placeholder = archive.getSheetByName('Archive Placeholder');
    if (placeholder) archive.deleteSheet(placeholder);
  }

  ui.alert(`Moved ${movedCount} monthly sheets to archive.`);

  // --- Step 3: Trim Engagement Information in chunks ---
  const infoSheet = ss.getSheetByName(CONFIG.ENGAGE_SHEET_NAME);
  if (infoSheet) {
    let lastRow = infoSheet.getLastRow();
    if (lastRow > 1) {
      const headers = infoSheet.getRange(1, 1, 1, infoSheet.getLastColumn()).getValues()[0];
      const sourceMonthCol = headers.indexOf('Source Month') + 1 || 30;
      let totalDeleted = 0;
      let scanEndRow = lastRow;

      // Process from bottom to top in chunks
      while (scanEndRow > 1) {
        const startRow = Math.max(2, scanEndRow - CHUNK_SIZE + 1);
        const numRows = scanEndRow - startRow + 1;
        const chunkData = infoSheet.getRange(startRow, 1, numRows, infoSheet.getLastColumn()).getValues();

        // Identify rows to delete (in reverse order within chunk)
        for (let i = chunkData.length - 1; i >= 0; i--) {
          const row = chunkData[i];
          const contactDate = row[0];
          const sourceMonth = row[sourceMonthCol - 1];
          let year = null;
          if (contactDate instanceof Date && !isNaN(contactDate)) {
            year = getYearInCrmTimezone_(contactDate);
          } else if (sourceMonth && typeof sourceMonth === 'string') {
            const parts = sourceMonth.split('-');
            if (parts.length === 2) year = parseInt(parts[1]);
          }
          if (year && year < currentYear) {
            infoSheet.deleteRow(startRow + i);
            totalDeleted++;
          }
        }

        // Continue with rows above this chunk. Using getLastRow() here can loop
        // forever when a chunk contains no rows eligible for deletion.
        scanEndRow = startRow - 1;
      }
      ui.alert(`Trimmed Engagement sheet, deleted ${totalDeleted} old rows.`);
    }
  }

  // --- Step 4: Trim Conversion Tracking in chunks ---
  const conversionSheet = ss.getSheetByName(CONFIG.CONVERSION_SHEET_NAME);
  if (conversionSheet) {
    let lastRow = conversionSheet.getLastRow();
    if (lastRow > 1) {
      let totalDeleted = 0;
      let scanEndRow = lastRow;

      while (scanEndRow > 1) {
        const startRow = Math.max(2, scanEndRow - CHUNK_SIZE + 1);
        const numRows = scanEndRow - startRow + 1;
        const chunkData = conversionSheet.getRange(startRow, 1, numRows, conversionSheet.getLastColumn()).getValues();

        for (let i = chunkData.length - 1; i >= 0; i--) {
          const firstContact = chunkData[i][3]; // Column D
          if (firstContact instanceof Date && !isNaN(firstContact) && getYearInCrmTimezone_(firstContact) < currentYear) {
            conversionSheet.deleteRow(startRow + i);
            totalDeleted++;
          }
        }

        scanEndRow = startRow - 1;
      }
      ui.alert(`Trimmed Conversion Tracking, deleted ${totalDeleted} old rows.`);
    }
  }

  ui.alert(`✅ Archive complete! All data before ${currentYear} is in the archive spreadsheet.`);
}

function getCurrentArchiveYear_() {
  return getYearInCrmTimezone_(new Date());
}

function getYearInCrmTimezone_(date) {
  return Number(Utilities.formatDate(date, CONFIG.TIMEZONE, 'yyyy'));
}
