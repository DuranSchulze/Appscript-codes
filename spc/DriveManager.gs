/**
 * Google Drive File Manager
 * Version: 3.0 (Production Ready) Fixed for dynamic CONFIG
 */

const DriveManager = {
  /**
   * Get config dynamically
   */
  getConfig() {
    const scriptProps = PropertiesService.getScriptProperties();
    return {
      DRIVE_FOLDER_ID: scriptProps.getProperty("DRIVE_FOLDER_ID") || "",
      SUPPORTED_FORMATS: [
        "pdf",
        "jpg",
        "jpeg",
        "png",
        "doc",
        "docx",
        "xls",
        "xlsx",
      ],
      MAX_PDF_SIZE_BYTES: 50 * 1024 * 1024,
      MAX_PDF_PAGES: 1000,
    };
  },
  createOrGetDateSubfolder(processedDate) {
    try {
      const CONFIG = this.getConfig();
      const rootFolder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);

      const monthNames = [
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
      ];
      const month = String(processedDate.getMonth() + 1).padStart(2, "0");
      const monthName = monthNames[processedDate.getMonth()];
      const folderName = `${month} ${monthName}`;

      console.log(`Looking for month folder: ${folderName}`);

      const existingFolders = rootFolder.getFoldersByName(folderName);

      if (existingFolders.hasNext()) {
        const existingFolder = existingFolders.next();
        console.log(`✓ Using existing month subfolder: ${folderName}`);
        Logger.log(`Using existing month subfolder: ${folderName}`, "INFO");
        return existingFolder;
      }

      const newFolder = rootFolder.createFolder(folderName);
      console.log(`✓ Created new month subfolder: ${folderName}`);
      Logger.log(`Created new month subfolder: ${folderName}`, "INFO");

      return newFolder;
    } catch (error) {
      console.error("Failed to create/get month subfolder:", error);
      throw new Error(`Month subfolder creation failed: ${error.message}`);
    }
  },
  moveFileToSubfolder(file, targetFolder) {
    try {
      const fileId = file.getId();
      const targetFolderId = targetFolder.getId();
      const fileName = file.getName();

      console.log(`Attempting to move file: ${fileName}`);
      console.log(`Target folder: ${targetFolder.getName()}`);

      const parents = file.getParents();
      const parentIds = [];

      while (parents.hasNext()) {
        const parent = parents.next();
        parentIds.push(parent.getId());
      }

      if (parentIds.length === 0) {
        throw new Error("File has no parent folders");
      }

      try {
        Drive.Files.update({}, fileId, null, {
          addParents: targetFolderId,
          removeParents: parentIds.join(","),
          supportsAllDrives: true,
        });

        console.log(`✓ Moved file using Drive API v3: ${fileName}`);
        Logger.log(
          `Moved file ${fileName} to ${targetFolder.getName()}`,
          "INFO",
        );

        return DriveApp.getFileById(fileId);
      } catch (apiError) {
        console.warn(`Drive API v3 failed: ${apiError.message}`);

        targetFolder.addFile(file);

        const parentsToRemove = file.getParents();
        while (parentsToRemove.hasNext()) {
          const parent = parentsToRemove.next();
          if (parent.getId() !== targetFolderId) {
            parent.removeFile(file);
          }
        }

        console.log(`✓ Moved file using DriveApp fallback: ${fileName}`);
        return file;
      }
    } catch (error) {
      console.error(`CRITICAL: Failed to move file ${file.getName()}:`, error);
      throw new Error(`File move failed: ${error.message}`);
    }
  },
  getUnprocessedFiles() {
    try {
      const CONFIG = this.getConfig();
      const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
      const files = [];
      const fileIterator = folder.getFiles();

      while (fileIterator.hasNext()) {
        const file = fileIterator.next();
        const fileName = file.getName().toLowerCase();

        if (fileName.includes("[processed]")) {
          continue;
        }

        const extension = fileName.split(".").pop();
        if (CONFIG.SUPPORTED_FORMATS.includes(extension)) {
          files.push(file);
        }
      }

      console.log(`Found ${files.length} unprocessed files`);
      return files;
    } catch (error) {
      throw new Error(`Failed to get unprocessed files: ${error.message}`);
    }
  },
  getFileExtension(file) {
    const fileName = file.getName().toLowerCase();
    const parts = fileName.split(".");
    return parts.length > 1 ? parts.pop() : "";
  },
  isPdfFile(file) {
    return (
      this.getFileExtension(file) === "pdf" ||
      file.getMimeType() === MimeType.PDF ||
      file.getMimeType() === "application/pdf"
    );
  },
  countPdfPages(file) {
    try {
      if (!this.isPdfFile(file)) {
        return null;
      }

      const pdfText = file.getBlob().getDataAsString("ISO-8859-1");
      const matches = pdfText.match(/\/Type\s*\/Page\b/g);
      return matches ? matches.length : null;
    } catch (error) {
      console.warn(`Failed to count PDF pages for ${file.getName()}:`, error);
      return null;
    }
  },
  getFileAuditInfo(file, includePageCount = true) {
    const CONFIG = this.getConfig();
    const sizeBytes = file.getSize();
    const isPdf = this.isPdfFile(file);
    const pageCount =
      isPdf && includePageCount && sizeBytes <= CONFIG.MAX_PDF_SIZE_BYTES
        ? this.countPdfPages(file)
        : null;
    const issues = [];

    if (isPdf && sizeBytes > CONFIG.MAX_PDF_SIZE_BYTES) {
      issues.push(
        `PDF is ${(sizeBytes / (1024 * 1024)).toFixed(1)} MB; Gemini PDF limit is 50 MB`,
      );
    }

    if (isPdf && pageCount !== null && pageCount > CONFIG.MAX_PDF_PAGES) {
      issues.push(
        `PDF has ${pageCount} pages; Gemini PDF limit is 1000 pages`,
      );
    }

    return {
      name: file.getName(),
      url: this.getFileUrl(file),
      extension: this.getFileExtension(file),
      mimeType: file.getMimeType(),
      sizeBytes,
      sizeMb: sizeBytes / (1024 * 1024),
      isPdf,
      pageCount,
      canProcess: issues.length === 0,
      issues,
    };
  },
  validateFileForGemini(file) {
    return this.getFileAuditInfo(file, true);
  },
  auditUnprocessedFiles() {
    const files = this.getUnprocessedFiles();
    const auditRows = [];
    let totalPages = 0;
    let countedPdfFiles = 0;
    let filesOverLimit = 0;

    files.forEach((file) => {
      const audit = this.getFileAuditInfo(file, true);
      auditRows.push(audit);

      if (audit.pageCount !== null) {
        totalPages += audit.pageCount;
        countedPdfFiles++;
      }

      if (!audit.canProcess) {
        filesOverLimit++;
      }
    });

    return {
      totalFiles: files.length,
      totalPages,
      countedPdfFiles,
      filesOverLimit,
      rows: auditRows,
    };
  },
  renameProcessedFile(file, voucherData, addProcessedTag = false) {
    try {
      const voucherNo = voucherData.voucherNo || "NO_VOUCHER";
      const company = voucherData.company || "NO_COMPANY";
      const date = voucherData.errandDate || "NO_DATE";

      const cleanVoucherNo = voucherNo.replace(/[^a-zA-Z0-9-]/g, "_");
      const cleanCompany = company
        .replace(/[^a-zA-Z0-9-]/g, "_")
        .substring(0, 20);
      const cleanDate = date.replace(/[^0-9-]/g, "");

      const originalName = file.getName();
      const extension = originalName.split(".").pop();

      let newName;
      if (addProcessedTag) {
        newName = `${cleanVoucherNo}_${cleanCompany}_${cleanDate}_[PROCESSED].${extension}`;
      } else {
        newName = `${cleanVoucherNo}_${cleanCompany}_${cleanDate}.${extension}`;
      }

      file.setName(newName);
      console.log(`✓ Renamed: ${originalName} → ${newName}`);

      return newName;
    } catch (error) {
      console.warn("Failed to rename file:", error.message);
      return file.getName();
    }
  },
  markFileAsProcessed(file) {
    try {
      const currentName = file.getName();

      if (currentName.toLowerCase().includes("[processed]")) {
        return;
      }

      const nameParts = currentName.split(".");
      const extension = nameParts.pop();
      const baseName = nameParts.join(".");
      const newName = `${baseName}_[PROCESSED].${extension}`;

      file.setName(newName);
      console.log(`✓ Marked as processed: ${newName}`);
    } catch (error) {
      console.warn("Failed to mark as processed:", error.message);
    }
  },
  getFolderUrl(folder) {
    try {
      return folder.getUrl();
    } catch (error) {
      console.warn("Failed to get folder URL:", error.message);
      return "";
    }
  },
  getFileUrl(file) {
    try {
      return file.getUrl();
    } catch (error) {
      console.warn("Failed to get file URL:", error.message);
      return "";
    }
  },
  getFileContent(file) {
    try {
      const mimeType = file.getMimeType();
      const blob = file.getBlob();

      let contentType = "document";
      if (mimeType.includes("image")) {
        contentType = "image";
      } else if (mimeType.includes("pdf")) {
        contentType = "pdf";
      }

      return {
        type: contentType,
        data: blob,
        mimeType: mimeType,
        sizeBytes: file.getSize(),
      };
    } catch (error) {
      throw new Error(`Failed to get file content: ${error.message}`);
    }
  },
};
