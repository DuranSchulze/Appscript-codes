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
  Math,
  Utilities: {sleep() {}}
});
vm.runInContext(source, context, {filename: "code.gs"});

// Legacy blocks below exercise the native-forward delivery path.
context.getDeliveryMode_ = () => "forward";

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

{
  // A property-store failure after Gmail already accepted the forward must
  // not be recorded as a forward failure, which would duplicate the send.
  const properties = makeProperties();
  const failingProperties = {
    deleteProperty(key) {
      properties.deleteProperty(key);
    },
    getProperties() {
      return properties.getProperties();
    },
    getProperty(key) {
      return properties.getProperty(key);
    },
    setProperty(key, value) {
      if (key.startsWith("AF_FORWARDED_")) {
        throw new Error("store unavailable");
      }
      properties.setProperty(key, value);
    }
  };
  const thread = {
    addLabel() {},
    getId() {
      return "bk-thread";
    }
  };
  const message = {
    forward() {},
    getDate() {
      return new Date();
    },
    getId() {
      return "bk-message";
    },
    getSubject() {
      return "Bookkeeping subject";
    },
    getThread() {
      return thread;
    }
  };

  const result = context.processCandidateMessages_(
    [message],
    failingProperties,
    {retryOnly: false}
  );
  assert.equal(result.forwarded, 1);
  assert.equal(result.failed, 0);
  assert.equal(properties.values["AF_FAILED_RETRY_bk-message"], undefined);
}

{
  // An unreadable body after a rule already matched must not mark the
  // message as failed; it is still forwarded.
  const properties = makeProperties();
  const thread = {
    addLabel() {},
    getId() {
      return "kw-thread";
    }
  };
  const message = {
    forward() {},
    getDate() {
      return new Date();
    },
    getId() {
      return "kw-message";
    },
    getSubject() {
      return "Keyword subject";
    },
    getThread() {
      return thread;
    }
  };
  context.getMatchedKeywords_ = () => {
    throw new Error("body unavailable");
  };

  const result = context.processCandidateMessages_([message], properties, {
    retryOnly: false
  });
  assert.equal(result.forwarded, 1);
  assert.equal(result.failed, 0);
  assert.ok(properties.values["AF_FORWARDED_kw-message"]);
  context.getMatchedKeywords_ = () => [];
}

{
  // One unreadable message must not abort processing of the others.
  const properties = makeProperties();
  const thread = {
    addLabel() {},
    getId() {
      return "poison-thread";
    }
  };
  const makeMessage = id => ({
    forward() {},
    getDate() {
      return new Date();
    },
    getId() {
      return id;
    },
    getSubject() {
      return `Subject ${id}`;
    },
    getThread() {
      return thread;
    }
  });
  const previousMatcher = context.findMatchingRule_;
  context.findMatchingRule_ = message => {
    if (message.getId() === "poison") {
      throw new Error("getPlainBody failed");
    }
    return {
      sender: "sender@example.com",
      recipients: ["recipient@example.com"]
    };
  };

  const result = context.processCandidateMessages_(
    [makeMessage("poison"), makeMessage("healthy")],
    properties,
    {retryOnly: false}
  );
  assert.equal(result.detected, 1);
  assert.equal(result.forwarded, 1);
  assert.equal(result.failed, 0);
  assert.ok(properties.values["AF_FORWARDED_healthy"]);
  context.findMatchingRule_ = previousMatcher;
}

{
  // An expired soft deadline defers everything instead of crashing.
  const properties = makeProperties();
  const result = context.processCandidateMessages_([], properties, {
    retryOnly: false,
    deadlineTimestamp: Date.now() - 1
  });
  assert.deepEqual(JSON.parse(JSON.stringify(result)), {
    detected: 0,
    forwarded: 0,
    failed: 0,
    deferred: 0
  });
}

{
  // The reconciliation sweep targets Detected conversations and heals
  // orphaned messages through the same record-based pipeline.
  const properties = makeProperties();
  const queries = [];
  const thread = {
    addLabel() {},
    getId() {
      return "reconcile-thread";
    }
  };
  const message = {
    forward() {},
    getDate() {
      return new Date();
    },
    getId() {
      return "reconcile-message";
    },
    getSubject() {
      return "Reconcile subject";
    },
    getThread() {
      return thread;
    },
    isDraft() {
      return false;
    },
    isInTrash() {
      return false;
    }
  };
  context.GmailApp = {
    search(query) {
      queries.push(query);
      return [{
        getMessages() {
          return [message];
        }
      }];
    }
  };

  const result = context.reconcileDetectedEmails_(
    properties,
    Date.now() + 60000
  );
  assert.ok(
    queries.some(query => query.includes("label:AutoForward/Detected"))
  );
  assert.equal(result.forwarded, 1);
  assert.ok(properties.values["AF_FORWARDED_reconcile-message"]);
}

{
  // A transient Gmail search error is retried once before failing.
  let calls = 0;
  context.GmailApp = {
    search() {
      calls++;
      if (calls === 1) {
        throw new Error("backend error");
      }
      return [];
    }
  };
  assert.deepEqual(
    context.searchGmailPage_("q", 0, 10, "transient"),
    []
  );
  assert.equal(calls, 2);
}

{
  // Notify delivery mode re-sends an authenticated copy from the forwarding
  // account instead of Gmail's native forward.
  const properties = makeProperties();
  const sent = [];
  const thread = {
    addLabel() {},
    getId() {
      return "notify-thread";
    }
  };
  const message = {
    forward() {
      throw new Error("native forward must not be used in notify mode");
    },
    getAttachments() {
      return [{getSize() { return 1024; }}];
    },
    getDate() {
      return new Date("2026-08-28T10:00:00Z");
    },
    getFrom() {
      return "Bank <no-reply@bank.com>";
    },
    getId() {
      return "notify-message";
    },
    getPlainBody() {
      return "Your security code is 123456.";
    },
    getSubject() {
      return "Security alert";
    },
    getThread() {
      return thread;
    }
  };
  context.GmailApp = {
    sendEmail(recipients, subject, body, options) {
      sent.push({recipients, subject, body, options});
    }
  };
  context.getDeliveryMode_ = () => "notify";
  context.findMatchingRule_ = () => ({
    sender: "no-reply@bank.com",
    recipients: ["recipient@example.com", "second@example.com"]
  });

  const result = context.processCandidateMessages_([message], properties, {
    retryOnly: false
  });
  assert.equal(result.forwarded, 1);
  assert.equal(result.failed, 0);
  assert.equal(sent.length, 1);
  assert.equal(
    sent[0].recipients,
    "recipient@example.com,second@example.com"
  );
  assert.equal(sent[0].subject, "Fwd: Security alert");
  assert.equal(sent[0].options.replyTo, "no-reply@bank.com");
  assert.equal(sent[0].options.attachments.length, 1);
  assert.ok(sent[0].body.includes("no-reply@bank.com"));
  assert.ok(sent[0].body.includes("Your security code is 123456."));
  assert.ok(properties.values["AF_FORWARDED_notify-message"]);

  context.getDeliveryMode_ = () => "forward";
}

{
  // Forward delivery mode keeps using Gmail's native forward.
  const properties = makeProperties();
  let nativeForwards = 0;
  const thread = {
    addLabel() {},
    getId() {
      return "native-thread";
    }
  };
  const message = {
    forward() {
      nativeForwards++;
    },
    getDate() {
      return new Date();
    },
    getId() {
      return "native-message";
    },
    getSubject() {
      return "Native subject";
    },
    getThread() {
      return thread;
    }
  };
  context.GmailApp = {
    sendEmail() {
      throw new Error("sendEmail must not be used in forward mode");
    }
  };
  context.getDeliveryMode_ = () => "forward";

  const result = context.processCandidateMessages_([message], properties, {
    retryOnly: false
  });
  assert.equal(result.forwarded, 1);
  assert.equal(nativeForwards, 1);
}

console.log("code.gs tests passed");
