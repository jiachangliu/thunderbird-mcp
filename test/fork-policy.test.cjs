"use strict";

const assert = require("node:assert/strict");
const { readFileSync } = require("node:fs");
const path = require("node:path");
const { test } = require("node:test");
const vm = require("node:vm");
const { loadPrivacyRuntime } = require("./helpers/privacy-access.cjs");

const folderPath = "imap://fixture/Inbox";
const messageId = "fixture@example.test";

test("authenticated HTTP refuses all delivery paths before dispatch", async () => {
  const runtime = loadPrivacyRuntime();
  const blocked = [
    ...["sendMail", "sendMailNow", "forwardMessage", "applyFilters", "createFilter", "updateFilter",
      "createEvent", "updateEvent", "deleteEvent", "createTask", "updateTask"].map(name => [name, {}]),
    ["saveDraft", { sendAt: "2027-02-01T10:00:00Z" }],
    ["replyToMessage", { skipReview: true }],
    ["replyToMessage", { skipReview: "true" }],
  ];
  for (const [name, args] of blocked) {
    const response = await runtime.http("tools/call", { name, arguments: args });
    assert.equal(response.result.isError, true, name);
    assert.match(response.result.content[0].text, /draft-only/);
  }
  assert.deepEqual(runtime.calls, []);
});

test("authenticated HTTP forces a reply into Drafts even if the caller opts out", async () => {
  const runtime = loadPrivacyRuntime();
  const response = await runtime.http("tools/call", { name: "replyToMessage",
    arguments: { folderPath, messageId, body: "Reply", saveAsDraft: false },
  });
  assert.equal(response.error, undefined);
  assert.equal(runtime.calls.length, 1);
  assert.equal(runtime.calls[0].name, "replyToMessage");
  assert.deepEqual(runtime.calls[0].args.slice(-2), [false, true]);
});

test("fork tools keep unread ordering, text fallback, and explicit read-state writes", async () => {
  const source = readFileSync(path.join(__dirname, "../extension/mcp_server/api.js"), "utf8");
  const start = source.indexOf("// BEGIN FORK COMPATIBILITY");
  const end = source.indexOf("// END TOOL CALL DISPATCH", start);
  assert.ok(start !== -1 && end > start);
  const headers = [
    { messageId: "old", date: 10, isRead: false },
    { messageId: "read", date: 40, isRead: true },
    { messageId: "new", date: 30, isRead: false },
  ];
  const inbox = { URI: folderPath };
  const account = { key: "allowed", incomingServer: { rootFolder: { getFolderWithFlags: () => inbox } } };
  const writes = [];
  const runtime = {
    MailServices: { accounts: { defaultAccount: { key: "restricted" } } },
    getAccessibleAccounts: () => [account],
    Ci: { nsMsgFolderFlags: { Inbox: 4096 } },
    openFolder: uri => uri === folderPath ? { folder: inbox, db: { enumerateMessages: () => headers } } : { error: "restricted" },
    getMessage: (id, uri, saveAttachments, format) => {
      assert.equal(uri, folderPath);
      assert.equal(saveAttachments, false);
      assert.equal(format, "text");
      return { id, body: "decoded body", bodyIsHtml: false };
    },
    updateMessage: (...args) => { writes.push(args); return { success: true }; },
    MESSAGE_READ_TIMEOUT_MS: 1000,
    MESSAGE_READ_OUTPUT_MAX_BYTES: 50000,
    TextEncoder,
    truncateHtmlForParsing: html => ({ html, truncated: false }),
  };
  vm.createContext(runtime);
  vm.runInContext(source.slice(start, end) + "\nthis.invoke = dispatchTool;", runtime);
  assert.equal((await runtime.invoke("getLatestUnread", {})).id, "new");
  const batch = await runtime.invoke("getLatestUnreadBatch", { limit: 2 });
  assert.deepEqual(Array.from(batch.items, item => [item.id, item.bodyText, item.bodyType]),
    [["new", "decoded body", "text/plain"], ["old", "decoded body", "text/plain"]]);
  assert.equal(writes.length, 0);
  assert.equal((await runtime.invoke("getLatestUnread", { folderPath: "restricted" })).error, "restricted");
  const marked = await runtime.invoke("setMessageRead", { folderPath, messageId, read: false });
  assert.equal(marked.success, true);
  assert.deepEqual(writes, [[messageId, undefined, folderPath, false]]);
});
