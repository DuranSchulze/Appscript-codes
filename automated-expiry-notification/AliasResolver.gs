// 36 Send Alias Resolver
//
// Per-row outgoing emails set the From address to the row's Assigned Staff
// Email so the client sees the staff member as the sender. Gmail only allows
// this when the address is registered as a "Send mail as" alias on the
// script-runner's account. This module gates the send on that check.

var __ALIAS_CACHE = null;

function getVerifiedSenderAliases() {
  if (__ALIAS_CACHE) return __ALIAS_CACHE;

  var lookup = {};

  function add(email) {
    var normalized = String(email || "").trim().toLowerCase();
    if (normalized) lookup[normalized] = true;
  }

  try {
    add(Session.getEffectiveUser().getEmail());
  } catch (e) {}
  try {
    add(Session.getActiveUser().getEmail());
  } catch (e) {}

  try {
    var aliases = GmailApp.getAliases() || [];
    for (var i = 0; i < aliases.length; i++) add(aliases[i]);
  } catch (e) {}

  __ALIAS_CACHE = lookup;
  return lookup;
}

function canSendAs(email) {
  var normalized = String(email || "").trim().toLowerCase();
  if (!normalized) return false;
  var lookup = getVerifiedSenderAliases();
  return !!lookup[normalized];
}

function listVerifiedSenderAliases() {
  var lookup = getVerifiedSenderAliases();
  var list = [];
  for (var key in lookup) list.push(key);
  return list;
}

function resetVerifiedSenderAliasCache() {
  __ALIAS_CACHE = null;
}
