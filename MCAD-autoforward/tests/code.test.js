const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const source = fs.readFileSync(path.join(__dirname, "..", "code.gs"), "utf8");
const context = vm.createContext({
  console,
  Logger: {log() {}},
  Map,
  Set,
  Date,
  JSON,
  Math
});
vm.runInContext(source, context, {filename: "code.gs"});

function makeProperties(initial = {}) {
  const values = Object.assign({}, initial);
  return {
    deleteProperty(key) {
      delete values[key];
    },
    getProperties() {
      return Object.assign({}, values);
    },
    getProperty(key) {
      return Object.prototype.hasOwnProperty.call(values, key)
        ? values[key]
        : null;
    },
    setProperty(key, value) {
      values[key] = String(value);
    },
    values
  };
}

{
  const properties = makeProperties();
  context.setBackfillState_(properties, {
    phase: "unread",
    senderBatch: 2,
    start: 100
  });
  assert.deepEqual(
    JSON.parse(JSON.stringify(context.getBackfillState_(properties))),
    {phase: "unread", senderBatch: 2, start: 100}
  );
}

{
  context.getRules_ = () => Array.from({length: 120}, (_, index) => ({
    sender: `sender${index}@example.com`
  }));
  const batches = context.getSenderBatches_();
  assert.equal(batches.length, 3);
  assert.equal((batches[0].match(/from:/g) || []).length, 50);
  assert.equal((batches[2].match(/from:/g) || []).length, 20);
}

{
  context.getRules_ = () => [{sender: "monitored@example.com"}];
  const makeMessage = (id, unread) => ({
    getDate() {
      return new Date(unread ? 100 : 200);
    },
    getId() {
      return id;
    },
    isDraft() {
      return false;
    },
    isInInbox() {
      return true;
    },
    isInTrash() {
      return false;
    },
    isUnread() {
      return unread;
    }
  });
  const unreadMessage = makeMessage("unread", true);
  const readMessage = makeMessage("read", false);
  const makeThread = message => ({
    getMessages() {
      return [message];
    },
    isInSpam() {
      return false;
    }
  });
  const queries = [];
  context.GmailApp = {
    search(query, start) {
      queries.push(query);
      if (start > 0 || query.includes("in:spam")) {
        return [];
      }
      return query.includes("is:unread")
        ? [makeThread(unreadMessage)]
        : [makeThread(readMessage)];
    }
  };
  const candidates = context.getRecentCandidateMessages_();
  assert.deepEqual(
    Array.from(candidates, message => message.getId()),
    ["unread", "read"]
  );
  assert.ok(queries.every(query => !query.includes("-label:")));
  assert.ok(queries.some(query => query.includes("from:monitored@example.com")));
}

{
  const properties = makeProperties({
    AF_BACKFILL_STATE: JSON.stringify({
      phase: "unread",
      senderBatch: 0,
      start: 0
    })
  });
  context.getUserProperties_ = () => properties;
  assert.equal(context.getRequiredTriggerHandlers_().length, 5);
  context.setBackfillState_(properties, {
    phase: "complete",
    senderBatch: 0,
    start: 0
  });
  assert.equal(context.getRequiredTriggerHandlers_().length, 4);
}

{
  const properties = makeProperties();
  context.advanceBackfillPhase_(
    properties,
    {phase: "unread", senderBatch: 3, start: 100},
    null
  );
  assert.equal(context.getBackfillState_(properties).phase, "read");
  context.advanceBackfillPhase_(
    properties,
    {phase: "read", senderBatch: 3, start: 100},
    null
  );
  assert.equal(context.getBackfillState_(properties).phase, "spam");
}

{
  const properties = makeProperties();
  const labels = [];
  const forwarded = [];
  const thread = {
    addLabel(label) {
      labels.push(label);
    },
    getId() {
      return "thread-1";
    }
  };
  const messages = Array.from({length: 30}, (_, index) => ({
    forward() {
      forwarded.push(index);
    },
    getDate() {
      return new Date(1_000_000 + index);
    },
    getId() {
      return `message-${index}`;
    },
    getSubject() {
      return `Subject ${index}`;
    },
    getThread() {
      return thread;
    }
  }));

  context.getOrCreateLabel_ = name => name;
  context.findMatchingRule_ = () => ({
    sender: "sender@example.com",
    recipients: ["recipient@example.com"]
  });
  context.getMatchedKeywords_ = () => [];
  context.recordForwardForDailySummary_ = () => {};
  context.syncThreadFailureLabels_ = () => {};

  const result = context.processCandidateMessages_(messages, properties, {
    retryOnly: false
  });
  assert.deepEqual(JSON.parse(JSON.stringify(result)), {
    detected: 30,
    forwarded: 25,
    failed: 0,
    deferred: 5
  });
  assert.equal(forwarded.length, 25);
  assert.equal(labels.filter(label => label === "AutoForward/Detected").length, 1);
}

{
  const properties = makeProperties();
  let shouldFail = true;
  const thread = {
    addLabel() {},
    getId() {
      return "retry-thread";
    }
  };
  const message = {
    forward() {
      if (shouldFail) {
        throw new Error("temporary forwarding failure");
      }
    },
    getDate() {
      return new Date();
    },
    getId() {
      return "retry-message";
    },
    getSubject() {
      return "Retry subject";
    },
    getThread() {
      return thread;
    }
  };

  const first = context.processCandidateMessages_([message], properties, {
    retryOnly: false
  });
  assert.equal(first.failed, 1);
  assert.equal(
    JSON.parse(properties.values["AF_FAILED_RETRY_retry-message"]).count,
    1
  );

  const normalWorker = context.processCandidateMessages_([message], properties, {
    retryOnly: false
  });
  assert.equal(normalWorker.detected, 0);

  const retryState = JSON.parse(
    properties.values["AF_FAILED_RETRY_retry-message"]
  );
  retryState.lastAttempt = 0;
  properties.setProperty(
    "AF_FAILED_RETRY_retry-message",
    JSON.stringify(retryState)
  );
  shouldFail = false;
  const retryWorker = context.processCandidateMessages_([message], properties, {
    retryOnly: true
  });
  assert.equal(retryWorker.forwarded, 1);
  assert.ok(properties.values["AF_FORWARDED_retry-message"]);
  assert.equal(properties.values["AF_FAILED_RETRY_retry-message"], undefined);
}

console.log("code.gs tests passed");
