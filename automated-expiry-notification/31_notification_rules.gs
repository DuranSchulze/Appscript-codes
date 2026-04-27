// 31 Notification Rules

function normalizeSendMode(modeValue) {
  var text = String(modeValue || "")
    .trim()
    .toLowerCase();
  if (!text) return SEND_MODE.AUTO;
  if (text === "auto") return SEND_MODE.AUTO;
  if (text === "hold") return SEND_MODE.HOLD;
  if (text === "manual only" || text === "manual-only" || text === "manual") {
    return SEND_MODE.MANUAL_ONLY;
  }
  return SEND_MODE.AUTO;
}

function getRowSendMode(row, colMap) {
  if (!colMap.SEND_MODE) return SEND_MODE.AUTO;
  return normalizeSendMode(getCellStr(row, colMap.SEND_MODE));
}

function getSendModeSkipReason(sendMode) {
  if (sendMode === SEND_MODE.HOLD) {
    return "Skipped by Send Mode: Hold";
  }
  if (sendMode === SEND_MODE.MANUAL_ONLY) {
    return "Skipped by Send Mode: Manual Only";
  }
  return "";
}
