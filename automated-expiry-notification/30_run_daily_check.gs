// 30 Run Daily Check

function manualRunNow() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    "Run Manual Check",
    "This will scan all Active/blank rows and send emails for any row whose target date is due (today or earlier), subject to Send Mode rules.\n\nProceed?",
    ui.ButtonSet.YES_NO,
  );
  if (confirm !== ui.Button.YES) return;

  try {
    runDailyCheck();
    ui.alert(
      "Done",
      "Manual check complete. Check the LOGS sheet for details.",
      ui.ButtonSet.OK,
    );
  } catch (e) {
    ui.alert("Error", "Manual check failed: " + e.message, ui.ButtonSet.OK);
  }
}

function runDailyCheck() {
  var ss = getAutomationSpreadsheet();
  var logsSheet = ensureLogsSheet(ss);
  var sheetConfigs = resolveAutomationSheets(ss);
  var senderEmail = getSenderAccountEmail();
  var trackingEnabled = !!getOpenTrackingBaseUrl();
  var today = getMidnight(new Date());

  var totalAllTabs = 0,
    sentAllTabs = 0,
    errorsAllTabs = 0;

  for (var t = 0; t < sheetConfigs.length; t++) {
    var sheetConfig = sheetConfigs[t];
    var tabName = sheetConfig.sheetName;
    var visaSheet = sheetConfig.sheet;

    if (!visaSheet) {
      appendLog(
        logsSheet,
        tabName,
        "",
        "ERROR",
        'Sheet "' +
          tabName +
          '" not found. Use "Configure Automation Sheet(s)" to fix.',
      );
      continue;
    }

    // Get per-tab configuration
    var headerRow = getTabHeaderRow(tabName);
    var dataStartRow = getTabDataStartRow(tabName);

    var colMap = buildColumnMap(visaSheet, tabName);
    var mapError = validateColumnMap(colMap, headerRow);
    if (mapError) {
      appendLog(
        logsSheet,
        tabName,
        "",
        "ERROR",
        "[" + tabName + "] " + mapError,
      );
      continue;
    }

    var lastRow = visaSheet.getLastRow();
    if (lastRow < dataStartRow) {
      appendLog(
        logsSheet,
        tabName,
        "",
        "INFO",
        "[" + tabName + "] No data rows found. Skipping.",
      );
      continue;
    }

    var numDataRows = lastRow - dataStartRow + 1;
    var numCols = visaSheet.getLastColumn();
    if (numCols === 0) {
      appendLog(
        logsSheet,
        tabName,
        "",
        "INFO",
        "[" + tabName + "] Tab has no columns. Skipping.",
      );
      continue;
    }
    var data = visaSheet
      .getRange(dataStartRow, 1, numDataRows, numCols)
      .getValues();
    var totalRows = data.length;
    var processed = 0,
      sent = 0,
      errors = 0,
      autoActivated = 0,
      skippedMode = 0,
      skippedFuture = 0;

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var rowIndex = dataStartRow + i;

      var clientName = getCellStr(row, colMap.CLIENT_NAME);
      var clientEmailRaw = getCellStr(row, colMap.CLIENT_EMAIL);
      var clientEmailList = parseClientEmails(clientEmailRaw);
      var staffEmail = getCellStr(row, colMap.STAFF_EMAIL);
      var docType = getCellStr(row, colMap.DOC_TYPE);
      var expiryRaw = colMap.EXPIRY_DATE ? row[colMap.EXPIRY_DATE - 1] : "";
      var noticeStr = getCellStr(row, colMap.NOTICE_DATE);
      var remarks = getCellStr(row, colMap.REMARKS);
      var attachRaw = getCellStr(row, colMap.ATTACHMENTS);
      var status = getCellStr(row, colMap.STATUS);
      var replyStatus = getCellStr(row, colMap.REPLY_STATUS);
      var sendMode = getRowSendMode(row, colMap);
      var templateContext = buildRowTemplateContext(visaSheet, tabName, row);

      if (!isProcessableStatus(status)) continue;

      if (isStatusBlank(status)) {
        if (colMap.STATUS) {
          setResolvedStatus(
            visaSheet,
            rowIndex,
            colMap,
            tabName,
            STATUS.ACTIVE,
          );
          status = resolveStatusValueForTab(
            visaSheet,
            tabName,
            colMap,
            STATUS.ACTIVE,
          );
          autoActivated++;
          appendLog(
            logsSheet,
            tabName,
            clientName,
            "INFO",
            "Blank Status auto-set to Active for processing.",
          );
        } else {
          status = STATUS.ACTIVE;
        }
      }

      var modeSkipReason = getSendModeSkipReason(sendMode);
      if (modeSkipReason) {
        appendLog(logsSheet, tabName, clientName, "SKIPPED", modeSkipReason);
        skippedMode++;
        continue;
      }

      processed++;

      var missing = [];
      if (!clientName) missing.push("Client Name");
      if (clientEmailList.length === 0) missing.push("Client Email");
      if (!expiryRaw) missing.push("Expiry Date");
      if (!noticeStr) missing.push("Notice Date");
      if (missing.length > 0) {
        var errMsg = "Missing required field(s): " + missing.join(", ");
        setResolvedStatus(visaSheet, rowIndex, colMap, tabName, STATUS.ERROR);
        appendLog(logsSheet, tabName, clientName, "ERROR", errMsg);
        errors++;
        continue;
      }

      var expiryDate =
        expiryRaw instanceof Date ? expiryRaw : new Date(expiryRaw);
      if (isNaN(expiryDate.getTime())) {
        setResolvedStatus(visaSheet, rowIndex, colMap, tabName, STATUS.ERROR);
        appendLog(
          logsSheet,
          tabName,
          clientName,
          "ERROR",
          "Invalid Expiry Date: " + expiryRaw,
        );
        errors++;
        continue;
      }

      var offset = parseNoticeOffset(noticeStr);
      if (offset === null) {
        setResolvedStatus(visaSheet, rowIndex, colMap, tabName, STATUS.ERROR);
        appendLog(
          logsSheet,
          tabName,
          clientName,
          "ERROR",
          'Cannot parse Notice Date: "' +
            noticeStr +
            '". ' +
            getSupportedNoticeDateHint(),
        );
        errors++;
        continue;
      }

      var targetDate = computeTargetDate(expiryDate, offset);
      var isNoticeDue = isTargetDateDue(targetDate, today);
      var isExpiryDay = isSameDay(expiryDate, today);
      var isPastExpiry =
        getMidnight(today).getTime() > getMidnight(expiryDate).getTime();
      var sameDayFinal = isSameDay(targetDate, expiryDate);
      var statusAllowsNotice = isStatusBlank(status) || isStatusActive(status);
      var statusAllowsFinal = statusAllowsNotice || isStatusNoticeSent(status);

      var shouldSendNotice = isNoticeDue && statusAllowsNotice && !isPastExpiry;
      var shouldSendFinal = isExpiryDay && statusAllowsFinal && !isPastExpiry;

      if (isPastExpiry) {
        appendLog(
          logsSheet,
          tabName,
          clientName,
          "SKIPPED",
          "Expiry date has already passed. No further reminder emails will be sent for this row.",
        );
      }

      if (isReplyStatusReplied(replyStatus)) {
        shouldSendNotice = false;
        if (!shouldSendFinal) {
          appendLog(
            logsSheet,
            tabName,
            clientName,
            "SKIPPED",
            "Reply keyword already received. Future reminder emails are suppressed for this row until the final expiry-day email.",
          );
        }
      }

      if (!shouldSendNotice && !shouldSendFinal) {
        skippedFuture++;
        continue;
      }

      var attachResult = resolveAttachments(attachRaw);
      if (attachResult.fatalError) {
        setResolvedStatus(visaSheet, rowIndex, colMap, tabName, STATUS.ERROR);
        appendLog(
          logsSheet,
          tabName,
          clientName,
          "ERROR",
          attachResult.fatalError,
        );
        errors++;
        continue;
      }
      if (attachResult.warnings && attachResult.warnings.length > 0) {
        appendLog(
          logsSheet,
          tabName,
          clientName,
          "INFO",
          "Attachment warning(s): " + attachResult.warnings.join("; "),
        );
      }
      var ccEmails = resolveCcEmails(clientEmailList, staffEmail);
      if (trackingEnabled) {
        colMap = ensureOpenTrackingColumns(visaSheet, tabName, colMap);
      }
      var baseSubject = buildEmailSubject(docType, clientName, expiryDate);
      var clientEmailDisplay = clientEmailList.join(", ");

      try {
        var sentThisRow = false;
        var displayName = getSenderDisplayName(senderEmail);

        if (shouldSendNotice) {
          var noticeToken = trackingEnabled ? generateOpenTrackingToken() : "";
          var noticeContent = buildStageEmailContent(
            remarks,
            clientName,
            expiryDate,
            docType,
            noticeToken,
            templateContext,
            "notice",
          );
          var noticeSubject = buildStageSubject(baseSubject, "notice");
          var noticeFallbackHtml = buildFallbackLinksHtml(
            attachResult.failedLinks,
          );
          var noticeHtmlBody = noticeFallbackHtml
            ? noticeContent.htmlBody + noticeFallbackHtml
            : noticeContent.htmlBody;
          var noticeSendResult = sendReminderEmails(
            clientEmailList,
            ccEmails,
            noticeSubject,
            noticeHtmlBody,
            attachResult.blobs,
            displayName,
          );
          var noticeMeta = noticeSendResult.meta;

          colMap = ensureReplyStatusColumn(visaSheet, tabName, colMap);
          writePostSendMetadata(visaSheet, rowIndex, colMap, {
            sentAt: new Date(),
            senderEmail: senderEmail,
            openToken: noticeToken,
            threadId: noticeMeta.threadId,
            messageId: noticeMeta.messageId,
          });

          if (sameDayFinal && shouldSendFinal) {
            colMap = ensureFinalNoticeColumns(visaSheet, tabName, colMap);
            writeFinalNoticeMetadata(visaSheet, rowIndex, colMap, {
              sentAt: new Date(),
              threadId: noticeMeta.threadId,
              messageId: noticeMeta.messageId,
            });
            shouldSendFinal = false;
            setResolvedStatus(
              visaSheet,
              rowIndex,
              colMap,
              tabName,
              STATUS.SENT,
            );
          } else {
            setResolvedStatus(
              visaSheet,
              rowIndex,
              colMap,
              tabName,
              STATUS.NOTICE_SENT,
            );
          }

          setStaffEmail(visaSheet, rowIndex, colMap.STAFF_EMAIL, senderEmail);
          appendLog(
            logsSheet,
            tabName,
            clientName,
            sameDayFinal ? "SENT_NOTICE_FINAL" : "SENT_NOTICE",
            "Email sent to " +
              noticeSendResult.success.join(", ") +
              (ccEmails.length > 0
                ? " (CC: " + ccEmails.join(", ") + ")"
                : "") +
              " | Stage: " +
              (sameDayFinal ? "notice+final" : "notice") +
              " | Mode: " +
              sendMode +
              " | Body: " +
              noticeContent.source +
              (attachResult.blobs.length > 0
                ? " | Attachments: " + attachResult.blobs.length
                : "") +
              (noticeToken ? " | Tracking: enabled" : "") +
              (noticeSendResult.failed.length > 0
                ? " | Failed Recipients: " +
                  noticeSendResult.failed
                    .map(function (item) {
                      return item.email + " (" + item.error + ")";
                    })
                    .join("; ")
                : "") +
              (senderEmail ? " | Sender: " + senderEmail : ""),
          );
          sent++;
          sentThisRow = true;
        }

        if (shouldSendFinal) {
          colMap = ensureFinalNoticeColumns(visaSheet, tabName, colMap);
          var finalToken = trackingEnabled ? generateOpenTrackingToken() : "";
          var finalContent = buildStageEmailContent(
            remarks,
            clientName,
            expiryDate,
            docType,
            finalToken,
            templateContext,
            "final",
          );
          var finalSubject = buildStageSubject(baseSubject, "final");
          var finalFallbackHtml = buildFallbackLinksHtml(
            attachResult.failedLinks,
          );
          var finalHtmlBody = finalFallbackHtml
            ? finalContent.htmlBody + finalFallbackHtml
            : finalContent.htmlBody;
          var finalSendResult = sendReminderEmails(
            clientEmailList,
            ccEmails,
            finalSubject,
            finalHtmlBody,
            attachResult.blobs,
            displayName,
          );
          var finalMeta = finalSendResult.meta;

          writeFinalNoticeMetadata(visaSheet, rowIndex, colMap, {
            sentAt: new Date(),
            threadId: finalMeta.threadId,
            messageId: finalMeta.messageId,
          });

          if (!isStatusNoticeSent(status) && !shouldSendNotice) {
            colMap = ensureReplyStatusColumn(visaSheet, tabName, colMap);
            writePostSendMetadata(visaSheet, rowIndex, colMap, {
              sentAt: new Date(),
              senderEmail: senderEmail,
              openToken: finalToken,
              threadId: finalMeta.threadId,
              messageId: finalMeta.messageId,
            });
          } else if (finalToken) {
            setCellValueIfColumn(
              visaSheet,
              rowIndex,
              colMap.OPEN_TOKEN,
              finalToken,
            );
          }

          setResolvedStatus(visaSheet, rowIndex, colMap, tabName, STATUS.SENT);
          setStaffEmail(visaSheet, rowIndex, colMap.STAFF_EMAIL, senderEmail);
          appendLog(
            logsSheet,
            tabName,
            clientName,
            "SENT_FINAL",
            "Email sent to " +
              finalSendResult.success.join(", ") +
              (ccEmails.length > 0
                ? " (CC: " + ccEmails.join(", ") + ")"
                : "") +
              " | Stage: final" +
              " | Mode: " +
              sendMode +
              " | Body: " +
              finalContent.source +
              (attachResult.blobs.length > 0
                ? " | Attachments: " + attachResult.blobs.length
                : "") +
              (finalToken ? " | Tracking: enabled" : "") +
              (finalSendResult.failed.length > 0
                ? " | Failed Recipients: " +
                  finalSendResult.failed
                    .map(function (item) {
                      return item.email + " (" + item.error + ")";
                    })
                    .join("; ")
                : "") +
              (senderEmail ? " | Sender: " + senderEmail : ""),
          );
          sent++;
          sentThisRow = true;
        }

        if (!sentThisRow) {
          skippedFuture++;
        }
      } catch (e) {
        setResolvedStatus(visaSheet, rowIndex, colMap, tabName, STATUS.ERROR);
        appendLog(
          logsSheet,
          tabName,
          clientName,
          "ERROR",
          "Send failed: " + e.message,
        );
        errors++;
      }
    }

    appendLog(
      logsSheet,
      tabName,
      "",
      "SUMMARY",
      "[" +
        tabName +
        "] Total: " +
        totalRows +
        " | Eligible: " +
        processed +
        " | Auto-Activated: " +
        autoActivated +
        " | Skipped (Mode): " +
        skippedMode +
        " | Skipped (Future): " +
        skippedFuture +
        " | Sent: " +
        sent +
        " | Errors: " +
        errors,
    );

    totalAllTabs += totalRows;
    sentAllTabs += sent;
    errorsAllTabs += errors;
  }

  if (sheetConfigs.length > 1) {
    appendLog(
      logsSheet,
      "",
      "",
      "SUMMARY",
      "All tabs complete. Tabs processed: " +
        sheetConfigs.length +
        " | Total Rows: " +
        totalAllTabs +
        " | Total Sent: " +
        sentAllTabs +
        " | Total Errors: " +
        errorsAllTabs,
    );
  }
}
