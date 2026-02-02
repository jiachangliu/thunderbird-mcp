/* global ExtensionCommon, ChromeUtils, Services, Cc, Ci */
"use strict";

/**
 * Thunderbird MCP Server Extension
 * Exposes email, calendar, and contacts via MCP protocol over HTTP.
 *
 * Architecture: MCP Client <-> mcp-bridge.cjs (stdio<->HTTP) <-> This extension (port 8765)
 *
 * Key quirks documented inline:
 * - MIME header decoding (mime2Decoded* properties)
 * - HTML body charset handling (emojis require HTML entity encoding)
 * - Compose window body preservation (must use New type, not Reply)
 * - IMAP folder sync (msgDatabase may be stale)
 */

const resProto = Cc[
  "@mozilla.org/network/protocol;1?name=resource"
].getService(Ci.nsISubstitutingProtocolHandler);

const MCP_PORT = 8765;
const MAX_SEARCH_RESULTS = 50;

var mcpServer = class extends ExtensionCommon.ExtensionAPI {
  getAPI(context) {
    const extensionRoot = context.extension.rootURI;
    const resourceName = "thunderbird-mcp";

    // Access WebExtension APIs (MV2/MV3) from this experiment script.
    // Depending on Thunderbird version/context, the WebExtension global may be reachable from different paths.
    let extBrowser = null;
    try {
      if (context && context.cloneScope && context.cloneScope.browser) {
        extBrowser = context.cloneScope.browser;
      }
    } catch {}
    try {
      if (!extBrowser && context && context.extension && context.extension.apiManager && context.extension.apiManager.global && context.extension.apiManager.global.browser) {
        extBrowser = context.extension.apiManager.global.browser;
      }
    } catch {}
    try {
      if (!extBrowser && context && context.extension && context.extension.views) {
        const views = context.extension.views;
        const iter = (views && typeof views[Symbol.iterator] === "function") ? views : [];
        for (const v of iter) {
          try {
            if (v && v.viewType === "background" && v.xulBrowser && v.xulBrowser.contentWindow && v.xulBrowser.contentWindow.browser) {
              extBrowser = v.xulBrowser.contentWindow.browser;
              break;
            }
          } catch {}
        }
      }
    } catch {}

    resProto.setSubstitutionWithFlags(
      resourceName,
      extensionRoot,
      resProto.ALLOW_CONTENT_ACCESS
    );

    const tools = [
      {
        name: "searchMessages",
        title: "Search Mail",
        description: "Find messages using Thunderbird's search index",
        inputSchema: {
          type: "object",
          properties: {
            query: { type: "string", description: "Text to search for in messages (searches subject, body, author)" }
          },
          required: ["query"],
        },
      },
      {
        name: "getLatestUnread",
        title: "Get Latest Unread",
        description: "Return the most recent unread message from a folder (defaults to Inbox)",
        inputSchema: {
          type: "object",
          properties: {
            folderPath: { type: "string", description: "Folder URI (e.g., imap://.../INBOX). Defaults to the account Inbox." }
          },
          required: [],
        },
      },
      {
        name: "getLatestUnreadBatch",
        title: "Get Latest Unread (Batch)",
        description: "Return up to N most recent unread messages from a folder WITHOUT changing read/unread state",
        inputSchema: {
          type: "object",
          properties: {
            folderPath: { type: "string", description: "Folder URI (e.g., imap://.../INBOX). Defaults to the account Inbox." },
            limit: { type: "number", description: "Max number of unread messages to return (default: 10, max: 50)" }
          },
          required: [],
        },
      },
      {
        name: "getMessage",
        title: "Get Message",
        description: "Read the full content of an email message by its ID",
        inputSchema: {
          type: "object",
          properties: {
            messageId: { type: "string", description: "The message ID (from searchMessages results)" },
            folderPath: { type: "string", description: "The folder URI path (from searchMessages results)" }
          },
          required: ["messageId", "folderPath"],
        },
      },
      {
        name: "sendMail",
        title: "Compose Mail",
        description: "Open a compose window with pre-filled recipient, subject, and body for user review before sending",
        inputSchema: {
          type: "object",
          properties: {
            to: { type: "string", description: "Recipient email address" },
            subject: { type: "string", description: "Email subject line" },
            body: { type: "string", description: "Email body text" },
            cc: { type: "string", description: "CC recipient (optional)" },
            isHtml: { type: "boolean", description: "Set to true if body contains HTML markup (default: false)" }
          },
          required: ["to", "subject", "body"]
        }
      },
      {
        name: "saveDraft",
        title: "Save Draft",
        description: "Create a draft message (saved to the account Drafts folder for cloud sync). Supports idempotency to prevent duplicates.",
        inputSchema: {
          type: "object",
          properties: {
            to: { type: "string", description: "Recipient email address" },
            subject: { type: "string", description: "Email subject line" },
            body: { type: "string", description: "Email body text" },
            cc: { type: "string", description: "CC recipient (optional)" },
            isHtml: { type: "boolean", description: "Set to true if body contains HTML markup (default: false)" },
            idempotencyKey: { type: "string", description: "Optional stable key to avoid creating duplicate drafts on retries" }
          },
          required: ["to", "subject", "body"]
        }
      },
      {
        name: "listCalendars",
        title: "List Calendars",
        description: "Return the user's calendars",
        inputSchema: { type: "object", properties: {}, required: [] },
      },
      {
        name: "searchContacts",
        title: "Search Contacts",
        description: "Find contacts the user interacted with",
        inputSchema: {
          type: "object",
          properties: {
            query: { type: "string", description: "Email address or name to search for" }
          },
          required: ["query"],
        },
      },
      {
        name: "replyToMessage",
        title: "Reply to Message",
        description: "Open a reply compose window for a specific message with proper threading",
        inputSchema: {
          type: "object",
          properties: {
            messageId: { type: "string", description: "The message ID to reply to (from searchMessages results)" },
            folderPath: { type: "string", description: "The folder URI path (from searchMessages results)" },
            body: { type: "string", description: "Reply body text" },
            replyAll: { type: "boolean", description: "Reply to all recipients (default: false)" },
            isHtml: { type: "boolean", description: "Set to true if body contains HTML markup (default: false)" }
          },
          required: ["messageId", "folderPath", "body"]
        }
      },
      {
        name: "openNativeReplyCompose",
        title: "Open Native Reply Compose",
        description: "Open a native Thunderbird Reply/ReplyAll compose window (with quoted original) for a message. Use UI automation to paste your reply at top.",
        inputSchema: {
          type: "object",
          properties: {
            messageId: { type: "string" },
            folderPath: { type: "string" },
            replyAll: { type: "boolean" }
          },
          required: ["messageId", "folderPath"]
        }
      },
      {
        name: "replyToMessageDraft",
        title: "Reply to Message (Save Draft)",
        description: "Create a reply draft saved to the account Drafts folder (for Outlook cloud sync). Includes quoted original message by default; supports idempotency to prevent duplicates.",
        inputSchema: {
          type: "object",
          properties: {
            messageId: { type: "string", description: "The message ID to reply to (from searchMessages results)" },
            folderPath: { type: "string", description: "The folder URI path (from searchMessages results)" },
            body: { type: "string", description: "Reply body text" },
            replyAll: { type: "boolean", description: "Reply to all recipients (default: false)" },
            isHtml: { type: "boolean", description: "Set to true if body contains HTML markup (default: false)" },
            idempotencyKey: { type: "string", description: "Optional stable key to avoid creating duplicate reply drafts on retries" },
            includeQuotedOriginal: { type: "boolean", description: "Whether to include quoted original message at bottom (default: true)" },
            useClosePromptSave: { type: "boolean", description: "Use Thunderbird's close-window prompt (auto-click Save) to produce an Outlook-sendable Draft (default: false)" }
          },
          required: ["messageId", "folderPath", "body"]
        }
      },
      {
        name: "replyToMessageDraftComposeApi",
        title: "Reply to Message (Save Draft via compose API)",
        description: "Create a reply draft using Thunderbird's WebExtension compose API (messages.query + compose.beginReply + compose.setComposeDetails + compose.saveMessage). This should preserve native reply quoting and produce a real Draft without OS-level UI automation.",
        inputSchema: {
          type: "object",
          properties: {
            messageId: { type: "string", description: "RFC822 Message-ID header (same as other tools; e.g. 4f23...@...). Will be resolved via messages.query({headerMessageId})." },
            replyAll: { type: "boolean", description: "Reply to all recipients (default: false)" },
            plainTextBody: { type: "string", description: "Reply body as plain text to insert ABOVE the quoted original. Newlines will be preserved." },
            htmlBody: { type: "string", description: "Reply body as HTML to insert ABOVE the quoted original." },
            includeQuotedOriginal: { type: "boolean", description: "Whether to keep the quoted original at bottom (default: true)." },
            closeAfterSave: { type: "boolean", description: "Close the compose tab/window after saving (default: true)." }
          },
          required: ["messageId"]
        }
      },
      {
        name: "debugContext",
        title: "Debug Context",
        description: "Return debugging info about the experiment context and any reachable WebExtension browser globals.",
        inputSchema: { type: "object", properties: {}, required: [] }
      },
      {
        name: "listLatestMessages",
        title: "List Latest Messages (by folder)",
        description: "List the most recent messages in a folder WITHOUT changing any state. Useful for Drafts verification.",
        inputSchema: {
          type: "object",
          properties: {
            folderPath: { type: "string", description: "Folder URI (e.g., imap://.../Drafts)" },
            limit: { type: "number", description: "Max messages to return (default 20, max 100)" }
          },
          required: ["folderPath"]
        }
      },
      {
        name: "setMessageRead",
        title: "Mark Message Read/Unread",
        description: "Mark a specific message as read or unread",
        inputSchema: {
          type: "object",
          properties: {
            messageId: { type: "string", description: "The message ID (from searchMessages results)" },
            folderPath: { type: "string", description: "The folder URI path (from searchMessages results)" },
            read: { type: "boolean", description: "true to mark read, false to mark unread" }
          },
          required: ["messageId", "folderPath", "read"],
        },
      },
      {
        name: "deleteMessages",
        title: "Delete Messages",
        description: "Delete specific messages by message-id from a folder (state-changing; use with care)",
        inputSchema: {
          type: "object",
          properties: {
            folderPath: { type: "string", description: "Folder URI" },
            messageIds: { type: "array", items: { type: "string" }, description: "List of message-id values to delete" }
          },
          required: ["folderPath", "messageIds"],
        },
      },
      {
        name: "getRawMessage",
        title: "Get Raw Message Source",
        description: "Fetch raw RFC822 source (headers+body) for a message. Useful for comparing Draft formatting.",
        inputSchema: {
          type: "object",
          properties: {
            messageId: { type: "string", description: "Message-ID" },
            folderPath: { type: "string", description: "Folder URI" }
          },
          required: ["messageId", "folderPath"]
        }
      },
      {
        name: "moveMessages",
        title: "Move Messages",
        description: "Move messages from one folder to another (e.g., Drafts -> Deleted Items)",
        inputSchema: {
          type: "object",
          properties: {
            fromFolderPath: { type: "string", description: "Source folder URI" },
            toFolderPath: { type: "string", description: "Destination folder URI" },
            messageIds: { type: "array", items: { type: "string" }, description: "List of message-id values" }
          },
          required: ["fromFolderPath", "toFolderPath", "messageIds"]
        }
      },
    ];

    return {
      mcpServer: {
        start: async function() {
          try {
            const { HttpServer } = ChromeUtils.importESModule(
              "resource://thunderbird-mcp/httpd.sys.mjs?" + Date.now()
            );
            const { NetUtil } = ChromeUtils.importESModule(
              "resource://gre/modules/NetUtil.sys.mjs"
            );
            const { MailServices } = ChromeUtils.importESModule(
              "resource:///modules/MailServices.sys.mjs"
            );

            let cal = null;
            try {
              const calModule = ChromeUtils.importESModule(
                "resource:///modules/calendar/calUtils.sys.mjs"
              );
              cal = calModule.cal;
            } catch {
              // Calendar not available
            }

            /**
             * CRITICAL: Must specify { charset: "UTF-8" } or emojis/special chars
             * will be corrupted. NetUtil defaults to Latin-1.
             */
            function readRequestBody(request) {
              const stream = request.bodyInputStream;
              return NetUtil.readInputStreamToString(stream, stream.available(), { charset: "UTF-8" });
            }

            /**
             * Email bodies may contain control characters (BEL, etc.) that break
             * JSON.stringify. Remove them but preserve \n, \r, \t.
             */
            function sanitizeForJson(text) {
              if (!text) return text;
              return text.replace(/[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]/g, '');
            }

            function searchMessages(query) {
              const results = [];
              const lowerQuery = query.toLowerCase();

              function searchFolder(folder) {
                if (results.length >= MAX_SEARCH_RESULTS) return;

                try {
                  // Attempt to refresh IMAP folders. This is async and may not
                  // complete before we read, but helps with stale data.
                  if (folder.server && folder.server.type === "imap") {
                    try {
                      folder.updateFolder(null);
                    } catch {
                      // updateFolder may fail, continue anyway
                    }
                  }

                  const db = folder.msgDatabase;
                  if (!db) return;

                  for (const msgHdr of db.enumerateMessages()) {
                    if (results.length >= MAX_SEARCH_RESULTS) break;

                    // IMPORTANT: Use mime2Decoded* properties for searching.
                    // Raw headers contain MIME encoding like "=?UTF-8?Q?...?="
                    // which won't match plain text searches.
                    const subject = (msgHdr.mime2DecodedSubject || msgHdr.subject || "").toLowerCase();
                    const author = (msgHdr.mime2DecodedAuthor || msgHdr.author || "").toLowerCase();
                    const recipients = (msgHdr.mime2DecodedRecipients || msgHdr.recipients || "").toLowerCase();

                    if (subject.includes(lowerQuery) ||
                        author.includes(lowerQuery) ||
                        recipients.includes(lowerQuery)) {
                      results.push({
                        id: msgHdr.messageId,
                        subject: msgHdr.mime2DecodedSubject || msgHdr.subject,
                        author: msgHdr.mime2DecodedAuthor || msgHdr.author,
                        recipients: msgHdr.mime2DecodedRecipients || msgHdr.recipients,
                        date: msgHdr.date ? new Date(msgHdr.date / 1000).toISOString() : null,
                        folder: folder.prettyName,
                        folderPath: folder.URI,
                        read: msgHdr.isRead,
                        flagged: msgHdr.isFlagged
                      });
                    }
                  }
                } catch {
                  // Skip inaccessible folders
                }

                if (folder.hasSubFolders) {
                  for (const subfolder of folder.subFolders) {
                    if (results.length >= MAX_SEARCH_RESULTS) break;
                    searchFolder(subfolder);
                  }
                }
              }

              for (const account of MailServices.accounts.accounts) {
                if (results.length >= MAX_SEARCH_RESULTS) break;
                searchFolder(account.incomingServer.rootFolder);
              }

              return results;
            }

            function searchContacts(query) {
              const results = [];
              const lowerQuery = query.toLowerCase();

              for (const book of MailServices.ab.directories) {
                for (const card of book.childCards) {
                  if (card.isMailList) continue;

                  const email = (card.primaryEmail || "").toLowerCase();
                  const displayName = (card.displayName || "").toLowerCase();
                  const firstName = (card.firstName || "").toLowerCase();
                  const lastName = (card.lastName || "").toLowerCase();

                  if (email.includes(lowerQuery) ||
                      displayName.includes(lowerQuery) ||
                      firstName.includes(lowerQuery) ||
                      lastName.includes(lowerQuery)) {
                    results.push({
                      id: card.UID,
                      displayName: card.displayName,
                      email: card.primaryEmail,
                      firstName: card.firstName,
                      lastName: card.lastName,
                      addressBook: book.dirName
                    });
                  }

                  if (results.length >= MAX_SEARCH_RESULTS) break;
                }
                if (results.length >= MAX_SEARCH_RESULTS) break;
              }

              return results;
            }

            function listCalendars() {
              if (!cal) {
                return { error: "Calendar not available" };
              }
              try {
                return cal.manager.getCalendars().map(c => ({
                  id: c.id,
                  name: c.name,
                  type: c.type,
                  readOnly: c.readOnly
                }));
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function getFolderOrInbox(folderPath) {
              if (folderPath) {
                const f = MailServices.folderLookup.getFolderForURL(folderPath);
                return f || null;
              }
              try {
                const defaultAccount = MailServices.accounts.defaultAccount;
                if (!defaultAccount) return null;
                const root = defaultAccount.incomingServer.rootFolder;
                // Prefer folder flagged as Inbox if available.
                try {
                  const inbox = root.getFolderWithFlags(Ci.nsMsgFolderFlags.Inbox);
                  if (inbox) return inbox;
                } catch {
                  // fallthrough
                }
                // Fallback: first child named INBOX
                for (const sub of root.subFolders) {
                  if ((sub.prettyName || "").toUpperCase() === "INBOX" || (sub.name || "").toUpperCase() === "INBOX") {
                    return sub;
                  }
                }
                return root;
              } catch {
                return null;
              }
            }

            function getLatestUnread(folderPath) {
              return new Promise((resolve) => {
                try {
                  const folder = getFolderOrInbox(folderPath);
                  if (!folder) {
                    resolve({ error: folderPath ? `Folder not found: ${folderPath}` : "Inbox folder not found" });
                    return;
                  }

                  // Nudge IMAP sync.
                  try {
                    if (folder.server && folder.server.type === "imap") folder.updateFolder(null);
                  } catch {}

                  const db = folder.msgDatabase;
                  if (!db) {
                    resolve({ error: "Could not access folder database" });
                    return;
                  }

                  let latest = null;
                  for (const hdr of db.enumerateMessages()) {
                    if (!hdr.isRead) {
                      if (!latest || (hdr.date || 0) > (latest.date || 0)) {
                        latest = hdr;
                      }
                    }
                  }

                  if (!latest) {
                    resolve({ ok: true, message: "No unread messages found in folder", folderPath: folder.URI });
                    return;
                  }

                  // Reuse getMessage to extract body
                  getMessage(latest.messageId, folder.URI).then((msg) => {
                    resolve(msg);
                  });
                } catch (e) {
                  resolve({ error: e.toString() });
                }
              });
            }

            function getLatestUnreadBatch(folderPath, limit) {
              return new Promise((resolve) => {
                try {
                  const folder = getFolderOrInbox(folderPath);
                  if (!folder) {
                    resolve({ error: folderPath ? `Folder not found: ${folderPath}` : "Inbox folder not found" });
                    return;
                  }

                  // Nudge IMAP sync.
                  try {
                    if (folder.server && folder.server.type === "imap") folder.updateFolder(null);
                  } catch {}

                  const db = folder.msgDatabase;
                  if (!db) {
                    resolve({ error: "Could not access folder database" });
                    return;
                  }

                  const n = Math.max(1, Math.min(50, Number(limit || 10)));

                  // Collect unread headers, then sort by date descending.
                  const unread = [];
                  for (const hdr of db.enumerateMessages()) {
                    if (!hdr.isRead) unread.push(hdr);
                  }

                  unread.sort((a, b) => (b.date || 0) - (a.date || 0));
                  const selected = unread.slice(0, n);

                  if (selected.length === 0) {
                    resolve({ ok: true, message: "No unread messages found in folder", folderPath: folder.URI, items: [] });
                    return;
                  }

                  Promise.all(selected.map(h => getMessage(h.messageId, folder.URI)))
                    .then((items) => resolve({ ok: true, folderPath: folder.URI, count: items.length, items }))
                    .catch((e) => resolve({ error: e.toString() }));
                } catch (e) {
                  resolve({ error: e.toString() });
                }
              });
            }

            function getMessage(messageId, folderPath) {
              return new Promise((resolve) => {
                try {
                  const folder = MailServices.folderLookup.getFolderForURL(folderPath);
                  if (!folder) {
                    resolve({ error: `Folder not found: ${folderPath}` });
                    return;
                  }

                  const db = folder.msgDatabase;
                  if (!db) {
                    resolve({ error: "Could not access folder database" });
                    return;
                  }

                  let msgHdr = null;
                  for (const hdr of db.enumerateMessages()) {
                    if (hdr.messageId === messageId) {
                      msgHdr = hdr;
                      break;
                    }
                  }

                  if (!msgHdr) {
                    resolve({ error: `Message not found: ${messageId}` });
                    return;
                  }

                  const { MsgHdrToMimeMessage } = ChromeUtils.importESModule(
                    "resource:///modules/gloda/MimeMessage.sys.mjs"
                  );

                  function htmlToText(html) {
                    if (!html || typeof html !== "string") return "";
                    try {
                      // DOMParser is available in the extension context.
                      const doc = new DOMParser().parseFromString(html, "text/html");
                      const text = (doc && doc.body && doc.body.textContent) ? doc.body.textContent : "";
                      return text
                        .replace(/\r\n/g, "\n")
                        .replace(/\n{3,}/g, "\n\n")
                        .trim();
                    } catch {
                      // Fallback: strip tags naively
                      return html
                        .replace(/<style[\s\S]*?<\/style>/gi, " ")
                        .replace(/<script[\s\S]*?<\/script>/gi, " ")
                        .replace(/<[^>]+>/g, " ")
                        .replace(/\s+/g, " ")
                        .trim();
                    }
                  }

                  function extractFromParts(part) {
                    if (!part) return { textPlain: "", textHtml: "" };

                    // Some parts are leaf nodes with body; others have sub-parts.
                    const ct = (part.contentType || "").toLowerCase();
                    const body = typeof part.body === "string" ? part.body : "";

                    let textPlain = "";
                    let textHtml = "";

                    if (ct.startsWith("text/plain") && body) {
                      textPlain = body;
                    } else if (ct.startsWith("text/html") && body) {
                      textHtml = body;
                    }

                    // Recurse
                    if (Array.isArray(part.parts)) {
                      for (const sub of part.parts) {
                        const subRes = extractFromParts(sub);
                        if (!textPlain && subRes.textPlain) textPlain = subRes.textPlain;
                        if (!textHtml && subRes.textHtml) textHtml = subRes.textHtml;
                        if (textPlain && textHtml) break;
                      }
                    }

                    return { textPlain, textHtml };
                  }

                  MsgHdrToMimeMessage(msgHdr, null, (aMsgHdr, aMimeMsg) => {
                    if (!aMimeMsg) {
                      resolve({ error: "Could not parse message" });
                      return;
                    }

                    let body = "";
                    let bodyHtml = "";
                    let bodyText = "";
                    let bodyType = "unknown";

                    // 1) Preferred: Thunderbird's coercion to plaintext
                    try {
                      const coerced = aMimeMsg.coerceBodyToPlaintext();
                      if (coerced && typeof coerced === "string") {
                        body = sanitizeForJson(coerced);
                        bodyText = body;
                        bodyType = "text/plain";
                      }
                    } catch {
                      // fallthrough
                    }

                    // 2) Fallback: traverse MIME parts and pick first text/plain or text/html
                    if (!body) {
                      try {
                        const extracted = extractFromParts(aMimeMsg);
                        if (extracted.textPlain) {
                          body = sanitizeForJson(extracted.textPlain);
                          bodyText = body;
                          bodyType = "text/plain";
                        } else if (extracted.textHtml) {
                          bodyHtml = sanitizeForJson(extracted.textHtml);
                          bodyText = sanitizeForJson(htmlToText(bodyHtml));
                          bodyType = "text/html";
                          body = "(HTML body available in bodyHtml)";
                        }
                      } catch {
                        // fallthrough
                      }
                    }

                    if (!body && !bodyHtml) {
                      body = "(Could not extract body text)";
                    }

                    resolve({
                      id: msgHdr.messageId,
                      subject: msgHdr.subject,
                      author: msgHdr.author,
                      recipients: msgHdr.recipients,
                      ccList: msgHdr.ccList,
                      date: msgHdr.date ? new Date(msgHdr.date / 1000).toISOString() : null,
                      body,
                      bodyType,
                      bodyText,
                      bodyHtml
                    });
                  }, true, { examineEncryptedParts: true });

                } catch (e) {
                  resolve({ error: e.toString() });
                }
              });
            }

            /**
             * Opens a compose window with pre-filled fields.
             *
             * HTML body handling quirks:
             * 1. Strip newlines from HTML - Thunderbird adds <br> for each \n
             * 2. Encode non-ASCII as HTML entities - compose window has charset issues
             *    with emojis/unicode even with <meta charset="UTF-8">
             */
            function composeMail(to, subject, body, cc, isHtml) {
              try {
                const msgComposeService = Cc["@mozilla.org/messengercompose;1"]
                  .getService(Ci.nsIMsgComposeService);

                const msgComposeParams = Cc["@mozilla.org/messengercompose/composeparams;1"]
                  .createInstance(Ci.nsIMsgComposeParams);

                const composeFields = Cc["@mozilla.org/messengercompose/composefields;1"]
                  .createInstance(Ci.nsIMsgCompFields);

                composeFields.to = to || "";
                composeFields.cc = cc || "";
                composeFields.subject = subject || "";

                if (isHtml) {
                  let bodyText = (body || "").replace(/\n/g, '');
                  // Convert non-ASCII to HTML entities (handles emojis > U+FFFF)
                  bodyText = [...bodyText].map(c => c.codePointAt(0) > 127 ? `&#${c.codePointAt(0)};` : c).join('');
                  composeFields.body = bodyText.includes('<html')
                    ? bodyText
                    : `<html><head><meta charset="UTF-8"></head><body>${bodyText}</body></html>`;
                } else {
                  const htmlBody = (body || "")
                    .replace(/&/g, '&amp;')
                    .replace(/</g, '&lt;')
                    .replace(/>/g, '&gt;')
                    .replace(/\n/g, '<br>');
                  composeFields.body = `<html><body>${htmlBody}</body></html>`;
                }

                msgComposeParams.type = Ci.nsIMsgCompType.New;
                msgComposeParams.format = Ci.nsIMsgCompFormat.HTML;
                msgComposeParams.composeFields = composeFields;

                const defaultAccount = MailServices.accounts.defaultAccount;
                if (defaultAccount) {
                  msgComposeParams.identity = defaultAccount.defaultIdentity;
                }

                msgComposeService.OpenComposeWindowWithParams(null, msgComposeParams);

                return { success: true, message: "Compose window opened" };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function _getDraftsFolderURIForIdentity(identity) {
              try {
                if (!identity || !identity.key) return "";
                return Services.prefs.getCharPref(`mail.identity.${identity.key}.draft_folder`, "");
              } catch {
                return "";
              }
            }

            function _formatFrom(identity) {
              const email = (identity && identity.email) ? identity.email : "";
              const name = (identity && identity.fullName) ? identity.fullName : "";
              if (name && email) return `"${name}" <${email}>`;
              return email || "";
            }

            function _normalizeCRLF(text) {
              return String(text || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").replace(/\n/g, "\r\n");
            }

            function _quotePlainText(text) {
              const t = String(text || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
              return t
                .split("\n")
                .map((line) => `> ${line}`)
                .join("\r\n");
            }

            function _formatReplyHeaderBlock(msg) {
              // Mimic Outlook/Thunderbird-style header block (like the Cornell Tech Housing draft).
              const lines = [];
              lines.push("-----Original Message-----");

              const from = msg.author || "";
              const sent = msg.date ? new Date(msg.date).toLocaleString("en-US") : "";
              const to = msg.recipients || "";
              const cc = msg.ccList || "";
              const subject = msg.subject || "";

              if (from) lines.push(`From: ${from}`);
              if (sent) lines.push(`Sent: ${sent}`);
              if (to) lines.push(`To: ${to}`);
              if (cc) lines.push(`Cc: ${cc}`);
              if (subject) lines.push(`Subject: ${subject}`);

              return lines.join("\r\n");
            }

            function _formatReplyQuote(msg) {
              const headerBlock = _formatReplyHeaderBlock(msg);
              const original = msg.bodyText || msg.body || "";
              return headerBlock + "\r\n\r\n" + _normalizeCRLF(original);
            }

            function _simpleHash32Hex(str) {
              // Deterministic non-crypto hash for idempotency.
              let h = 2166136261;
              const s = String(str || "");
              for (let i = 0; i < s.length; i++) {
                h ^= s.charCodeAt(i);
                h = Math.imul(h, 16777619);
              }
              // >>> 0 to unsigned
              return (h >>> 0).toString(16);
            }

            function _sanitizeMessageIdToken(token) {
              return String(token || "")
                .toLowerCase()
                .replace(/[^a-z0-9._-]+/g, "-")
                .replace(/-+/g, "-")
                .replace(/^-|-$/g, "")
                .slice(0, 120);
            }

            function _getEmailDomain(identity) {
              try {
                const email = identity && identity.email ? String(identity.email) : "";
                const parts = email.split("@");
                return parts.length === 2 ? parts[1] : "local";
              } catch {
                return "local";
              }
            }

            function _makeRfc822Draft({ from, to, cc, subject, body, inReplyTo, references, idempotencyKey, identityKey, fcc, isHtml }) {
              const headers = [];

              // Thunderbird-created drafts include these (helps mimic behavior).
              headers.push("X-Mozilla-Status: 0001");
              headers.push("X-Mozilla-Status2: 00000000");

              if (from) headers.push(`From: ${from}`);
              if (to) headers.push(`To: ${to}`);
              if (cc) headers.push(`Cc: ${cc}`);
              if (subject !== undefined) headers.push(`Subject: ${subject || ""}`);
              headers.push(`Date: ${new Date().toString()}`);
              headers.push("User-Agent: Mozilla Thunderbird");
              headers.push("Content-Language: en-US");

              const token = idempotencyKey ? _sanitizeMessageIdToken(idempotencyKey) : null;
              const domain = _getEmailDomain({ email: (from || "").match(/<([^>]+)>/)?.[1] || "" }) || "local";
              const messageId = token
                ? `tb-mcp-draft-${token}@${domain}`
                : `tb-mcp-draft-${Date.now()}@${domain}`;
              headers.push(`Message-ID: <${messageId}>`);

              if (references) headers.push(`References: ${references}`);
              if (inReplyTo) headers.push(`In-Reply-To: <${inReplyTo.replace(/[<>]/g, "")}>`);

              headers.push("MIME-Version: 1.0");

              // Draft metadata similar to TB
              headers.push("X-Mozilla-Draft-Info: internal/draft; vcard=0; receipt=0; DSN=0; uuencode=0; attachmentreminder=0; deliveryformat=0");
              if (identityKey) headers.push(`X-Identity-Key: ${identityKey}`);
              if (fcc) headers.push(`Fcc: ${fcc}`);

              if (isHtml) {
                headers.push('Content-Type: text/html; charset="UTF-8"');
              } else {
                headers.push("Content-Type: text/plain; charset=UTF-8");
              }
              headers.push("Content-Transfer-Encoding: 8bit");

              const payload = isHtml
                ? String(body || "")
                : _normalizeCRLF(body || "");

              return { rfc822: headers.join("\r\n") + "\r\n\r\n" + payload + "\r\n", messageId };
            }

            function _writeStringToTempFileUtf8(filenamePrefix, content) {
              const tmp = Services.dirsvc.get("TmpD", Ci.nsIFile);
              tmp.append(`${filenamePrefix}-${Date.now()}.eml`);
              tmp.createUnique(Ci.nsIFile.NORMAL_FILE_TYPE, 0o600);

              const foStream = Cc["@mozilla.org/network/file-output-stream;1"].createInstance(Ci.nsIFileOutputStream);
              foStream.init(tmp, 0x02 | 0x08 | 0x20, 0o600, 0); // write | create | truncate

              const conv = Cc["@mozilla.org/intl/converter-output-stream;1"].createInstance(Ci.nsIConverterOutputStream);
              conv.init(foStream, "UTF-8");
              conv.writeString(String(content));
              conv.close();
              foStream.close();

              return tmp;
            }

            const _pendingCopyListeners = new Set();
            const _pendingTimers = new Set();
            const _pendingDraftMessageIds = new Set();

            function _copyFileToFolderAsDraft(file, folder, timeoutMs = 20000) {
              return new Promise((resolve, reject) => {
                try {
                  const copyService = MailServices.copy;

                  const timer = Cc["@mozilla.org/timer;1"].createInstance(Ci.nsITimer);
                  let done = false;

                  const finish = (err) => {
                    if (done) return;
                    done = true;
                    try { timer.cancel(); } catch {}
                    try { _pendingCopyListeners.delete(listener); } catch {}
                    if (err) reject(err);
                    else resolve(true);
                  };

                  const listener = {
                    QueryInterface: ChromeUtils.generateQI([Ci.nsIMsgCopyServiceListener]),
                    OnStartCopy() {},
                    OnProgress() {},
                    SetMessageKey() {},
                    GetMessageId() {},
                    OnStopCopy(status) {
                      try {
                        if (status && !Components.isSuccessCode(status)) {
                          finish(new Error(`Copy failed: ${status}`));
                          return;
                        }
                        finish(null);
                      } catch (e) {
                        finish(e);
                      }
                    },
                  };

                  _pendingCopyListeners.add(listener);

                  timer.init(
                    {
                      notify: () => finish(new Error("Copy timed out (no OnStopCopy)")),
                    },
                    timeoutMs,
                    Ci.nsITimer.TYPE_ONE_SHOT
                  );

                  let flags = 0;
                  try {
                    flags = Ci.nsMsgMessageFlags.Draft;
                  } catch {
                    flags = 0;
                  }

                  const copyFn = copyService.CopyFileMessage || copyService.copyFileMessage;
                  if (typeof copyFn !== "function") {
                    finish(new Error("Copy service missing CopyFileMessage/copyFileMessage"));
                    return;
                  }

                  copyFn.call(copyService, file, folder, null, false, flags, "", listener, null);
                } catch (e) {
                  reject(e);
                }
              });
            }

            const _pendingSendListeners = new Set();

            function _setClipboardText(text) {
              try {
                const clipboard = Cc["@mozilla.org/widget/clipboard;1"].getService(Ci.nsIClipboard);
                const trans = Cc["@mozilla.org/widget/transferable;1"].createInstance(Ci.nsITransferable);
                trans.init(null);
                trans.addDataFlavor("text/unicode");
                const str = Cc["@mozilla.org/supports-string;1"].createInstance(Ci.nsISupportsString);
                str.data = String(text || "");
                trans.setTransferData("text/unicode", str, str.data.length * 2);
                clipboard.setData(trans, null, Ci.nsIClipboard.kGlobalClipboard);
                return true;
              } catch {
                return false;
              }
            }

            function _sendCtrlKey(win, keyCode) {
              try {
                const wu = win.windowUtils || win
                  .QueryInterface(Ci.nsIInterfaceRequestor)
                  .getInterface(Ci.nsIDOMWindowUtils);
                if (!wu || typeof wu.sendKeyEvent !== "function") return false;
                const CTRL = 1; // KEYEVENT_CTRLDOWN
                wu.sendKeyEvent("keydown", keyCode, 0, CTRL);
                wu.sendKeyEvent("keyup", keyCode, 0, CTRL);
                return true;
              } catch {
                return false;
              }
            }

            function _saveDraftViaComposeWindow({ identity, to, cc, subject, bodyHtml, timeoutMs = 45000 }) {
              // Uses Thunderbird's native Save-as-Draft via nsIMsgCompose.SendMsg,
              // so Exchange/OWA sees a real Draft ("[Draft]" + Send enabled).
              return new Promise((resolve) => {
                try {
                  const msgComposeService = Cc["@mozilla.org/messengercompose;1"].getService(Ci.nsIMsgComposeService);

                  const msgComposeParams = Cc["@mozilla.org/messengercompose/composeparams;1"].createInstance(Ci.nsIMsgComposeParams);
                  const composeFields = Cc["@mozilla.org/messengercompose/composefields;1"].createInstance(Ci.nsIMsgCompFields);

                  composeFields.to = to || "";
                  composeFields.cc = cc || "";
                  composeFields.subject = subject || "";
                  composeFields.body = bodyHtml || "";

                  msgComposeParams.type = Ci.nsIMsgCompType.New;
                  msgComposeParams.format = Ci.nsIMsgCompFormat.HTML;
                  msgComposeParams.composeFields = composeFields;
                  if (identity) msgComposeParams.identity = identity;

                  let finished = false;
                  const finish = (ok, reason, details) => {
                    if (finished) return;
                    finished = true;
                    try { timer.cancel(); } catch {}
                    try { Services.ww.unregisterNotification(observer); } catch {}
                    resolve({ ok: !!ok, reason: reason || "", ...(details || {}) });
                  };

                  const timer = Cc["@mozilla.org/timer;1"].createInstance(Ci.nsITimer);
                  timer.init({ notify: () => { try { Services.ww.unregisterNotification(observer); } catch {} finish(false, "timeout"); } }, timeoutMs, Ci.nsITimer.TYPE_ONE_SHOT);

                  const observer = {
                    observe(subjectWin, topic) {
                      if (topic !== "domwindowopened") return;
                      const win = subjectWin;
                      win.addEventListener(
                        "load",
                        () => {
                          try {
                            const url = String(win.location);
                            if (!url.includes("messengercompose")) return;

                            // Delay one tick so the editor is really ready.
                            const runnable = {
                              run: () => {
                                try {
                                  let comp = null;
                                  try {
                                    if (typeof msgComposeService.GetMsgComposeForWindow === "function") {
                                      comp = msgComposeService.GetMsgComposeForWindow(win);
                                    }
                                  } catch {}
                                  if (!comp) {
                                    try {
                                      // Compose window usually exposes gMsgCompose.
                                      comp = win.gMsgCompose || null;
                                    } catch {}
                                  }
                                  if (!comp) {
                                    finish(false, "no-compose-instance");
                                    try { win.close(); } catch {}
                                    return;
                                  }

                                  const sendFn = comp.SendMsg || comp.sendMsg;
                                  if (typeof sendFn !== "function") {
                                    finish(false, "no-SendMsg");
                                    try { win.close(); } catch {}
                                    return;
                                  }

                                  // Listen for save completion.
                                  let gotStop = false;
                                  const sendListener = {
                                    QueryInterface: ChromeUtils.generateQI([Ci.nsIMsgSendListener]),
                                    onStartSending() {},
                                    onProgress() {},
                                    onStatus() {},
                                    onGetDraftFolderURI() {},
                                    onStopSending(msgID, status, msg, returnFile) {
                                      gotStop = true;
                                      _pendingSendListeners.delete(sendListener);
                                      const ok = !status || Components.isSuccessCode(status);
                                      try { win.close(); } catch {}
                                      finish(ok, ok ? "saved" : "send-failed", { msgID, status: String(status || "") });
                                    },
                                  };
                                  _pendingSendListeners.add(sendListener);

                                  // Start save.
                                  sendFn.call(comp, Ci.nsIMsgSend.nsMsgSaveAsDraft, identity || null, null, null, sendListener);

                                  // Do NOT close early. We must wait for onStopSending (or timeout) so the draft actually lands in Drafts.
                                } catch (e) {
                                  try { win.close(); } catch {}
                                  finish(false, "exception", { error: String(e) });
                                }
                              },
                            };
                            Services.tm.dispatchToMainThread(runnable);
                          } catch (e) {
                            finish(false, "exception", { error: String(e) });
                          }
                        },
                        { once: true }
                      );
                    },
                  };

                  Services.ww.registerNotification(observer);
                  msgComposeService.OpenComposeWindowWithParams(null, msgComposeParams);
                } catch (e) {
                  resolve({ ok: false, reason: "exception", error: String(e) });
                }
              });
            }

            function _findDraftByMessageId(folder, messageId) {
              try {
                const db = folder.msgDatabase;
                if (!db) return null;
                for (const hdr of db.enumerateMessages()) {
                  if (hdr.messageId === messageId) return hdr;
                }
                return null;
              } catch {
                return null;
              }
            }

            async function saveDraft(to, subject, body, cc, isHtml, idempotencyKey) {
              try {
                if (isHtml) {
                  return { error: "saveDraft currently supports isHtml=false only (plain text)" };
                }

                const defaultAccount = MailServices.accounts.defaultAccount;
                const identity = defaultAccount ? defaultAccount.defaultIdentity : null;
                if (!identity) {
                  return { error: "No default identity found" };
                }

                const draftsURI = _getDraftsFolderURIForIdentity(identity);
                if (!draftsURI) {
                  return { error: "Could not determine Drafts folder URI for identity" };
                }

                const draftsFolder = MailServices.folderLookup.getFolderForURL(draftsURI);
                if (!draftsFolder) {
                  return { error: `Drafts folder not found: ${draftsURI}` };
                }

                const key = idempotencyKey || `new-${_simpleHash32Hex(`${to}|${cc}|${subject}|${body}`)}`;
                const { rfc822, messageId } = _makeRfc822Draft({
                  from: _formatFrom(identity),
                  to: to || "",
                  cc: cc || "",
                  subject: subject || "",
                  body: body || "",
                  idempotencyKey: key,
                  identityKey: identity.key,
                  fcc: "",
                  isHtml: false,
                });

                // Duplicate prevention: prevent immediate retries from creating multiple drafts.
                if (_pendingDraftMessageIds.has(messageId)) {
                  return { success: true, message: "Draft already pending (idempotent)", draftsFolder: draftsURI, messageId };
                }

                // If a draft with this deterministic Message-ID already exists, do nothing.
                const existing = _findDraftByMessageId(draftsFolder, messageId);
                if (existing) {
                  return { success: true, message: "Draft already exists (idempotent)", draftsFolder: draftsURI, messageId };
                }

                const file = _writeStringToTempFileUtf8("tb-mcp-draft", rfc822);
                _pendingDraftMessageIds.add(messageId);

                // Schedule copy async so the HTTP handler returns immediately.
                // Use dispatchToMainThread instead of nsITimer (timers can be unreliable in this add-on context).
                const runnable = {
                  run: () => {
                    _copyFileToFolderAsDraft(file, draftsFolder, 30000)
                      .then(() => {
                        try { draftsFolder.updateFolder(null); } catch {}
                        try { Services.console.logStringMessage(`thunderbird-mcp: draft saved to ${draftsURI}`); } catch {}
                      })
                      .catch((e) => {
                        try { Services.console.logStringMessage(`thunderbird-mcp: draft save failed: ${e}`); } catch {}
                      })
                      .finally(() => {
                        try { _pendingDraftMessageIds.delete(messageId); } catch {}
                        try { _pendingTimers.delete(runnable); } catch {}
                      });
                  },
                };

                _pendingTimers.add(runnable);
                Services.tm.dispatchToMainThread(runnable);

                return { success: true, message: "Draft save scheduled (backend copy)", draftsFolder: draftsURI, messageId };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            /**
             * Opens a reply compose window for a message.
             *
             * IMPORTANT: Uses nsIMsgCompType.New instead of Reply/ReplyAll.
             * Using Reply type causes Thunderbird to overwrite our body with
             * the quoted original message. We manually set To, Subject, and
             * References headers for proper threading.
             *
             * Limitation: Does not include quoted original message text.
             */
            function setMessageRead(messageId, folderPath, read) {
              try {
                const folder = MailServices.folderLookup.getFolderForURL(folderPath);
                if (!folder) {
                  return { error: `Folder not found: ${folderPath}` };
                }

                const db = folder.msgDatabase;
                if (!db) {
                  return { error: "Could not access folder database" };
                }

                let msgHdr = null;
                for (const hdr of db.enumerateMessages()) {
                  if (hdr.messageId === messageId) {
                    msgHdr = hdr;
                    break;
                  }
                }

                if (!msgHdr) {
                  return { error: `Message not found: ${messageId}` };
                }

                const desiredRead = !!read;

                /**
                 * IMPORTANT:
                 * - msgHdr.markRead(...) updates the local message database.
                 * - For IMAP accounts, that may NOT propagate the \"\\Seen\" flag
                 *   back to the server (so other clients / Outlook Web won't update).
                 *
                 * To ensure server sync, prefer folder.markMessagesRead(...).
                 */
                let used = null;

                // Try to mark read via folder API (propagates to server for IMAP).
                // Thunderbird's JS folder helpers typically expect a plain JS array of nsIMsgDBHdr.
                try {
                  const hdrFolder = msgHdr.folder || folder;
                  if (hdrFolder && typeof hdrFolder.markMessagesRead === "function") {
                    hdrFolder.markMessagesRead([msgHdr], desiredRead);
                    used = "folder.markMessagesRead";
                  }
                } catch {
                  // Fall back below.
                }

                // Fallback: local-only header flag.
                if (!used) {
                  try {
                    msgHdr.markRead(desiredRead);
                    used = "msgHdr.markRead";
                  } catch (e) {
                    return { error: `Failed to mark read: ${e.toString()}` };
                  }
                }

                // Encourage committing changes to the local DB.
                try {
                  if (folder.msgDatabase && folder.msgDatabase.Commit) {
                    folder.msgDatabase.Commit(Ci.nsMsgDBCommitType.kLargeCommit);
                  }
                } catch {}

                return { success: true, messageId, folderPath, read: desiredRead, used };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function getIdentityForFolder(folder) {
              try {
                const account = MailServices.accounts.findAccountForServer(folder.server);
                if (account && account.defaultIdentity) return account.defaultIdentity;
              } catch {}
              try {
                const defaultAccount = MailServices.accounts.defaultAccount;
                if (defaultAccount && defaultAccount.defaultIdentity) return defaultAccount.defaultIdentity;
              } catch {}
              return null;
            }

            function replyToMessage(messageId, folderPath, body, replyAll, isHtml) {
              try {
                const folder = MailServices.folderLookup.getFolderForURL(folderPath);
                if (!folder) {
                  return { error: `Folder not found: ${folderPath}` };
                }

                const db = folder.msgDatabase;
                if (!db) {
                  return { error: "Could not access folder database" };
                }

                let msgHdr = null;
                for (const hdr of db.enumerateMessages()) {
                  if (hdr.messageId === messageId) {
                    msgHdr = hdr;
                    break;
                  }
                }

                if (!msgHdr) {
                  return { error: `Message not found: ${messageId}` };
                }

                const msgComposeService = Cc["@mozilla.org/messengercompose;1"]
                  .getService(Ci.nsIMsgComposeService);

                const msgComposeParams = Cc["@mozilla.org/messengercompose/composeparams;1"]
                  .createInstance(Ci.nsIMsgComposeParams);

                const composeFields = Cc["@mozilla.org/messengercompose/composefields;1"]
                  .createInstance(Ci.nsIMsgCompFields);

                if (replyAll) {
                  composeFields.to = msgHdr.author;
                  const otherRecipients = (msgHdr.recipients || "").split(",")
                    .map(r => r.trim())
                    .filter(r => r && !r.includes(folder.server.username));
                  if (otherRecipients.length > 0) {
                    composeFields.cc = otherRecipients.join(", ");
                  }
                } else {
                  composeFields.to = msgHdr.author;
                }

                const origSubject = msgHdr.subject || "";
                composeFields.subject = /^\s*re:/i.test(origSubject) ? origSubject : `Re: ${origSubject}`;

                // References header enables proper email threading
                composeFields.references = `<${messageId}>`;

                // Same HTML/charset handling as composeMail
                if (isHtml) {
                  let bodyText = (body || "").replace(/\n/g, '');
                  bodyText = [...bodyText].map(c => c.codePointAt(0) > 127 ? `&#${c.codePointAt(0)};` : c).join('');
                  composeFields.body = `<html><head><meta charset="UTF-8"></head><body>${bodyText}</body></html>`;
                } else {
                  const htmlBody = (body || "")
                    .replace(/&/g, '&amp;')
                    .replace(/</g, '&lt;')
                    .replace(/>/g, '&gt;')
                    .replace(/\n/g, '<br>');
                  composeFields.body = `<html><body>${htmlBody}</body></html>`;
                }

                // New type preserves our body; Reply type overwrites it
                msgComposeParams.type = Ci.nsIMsgCompType.New;
                msgComposeParams.format = Ci.nsIMsgCompFormat.HTML;
                msgComposeParams.composeFields = composeFields;
                msgComposeParams.identity = getIdentityForFolder(folder);

                msgComposeService.OpenComposeWindowWithParams(null, msgComposeParams);

                return { success: true, message: "Reply window opened" };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function openNativeReplyCompose(messageId, folderPath, replyAll) {
              try {
                const folder = MailServices.folderLookup.getFolderForURL(folderPath);
                if (!folder) {
                  return { error: `Folder not found: ${folderPath}` };
                }
                const db = folder.msgDatabase;
                if (!db) {
                  return { error: "Could not access folder database" };
                }
                let msgHdr = null;
                for (const hdr of db.enumerateMessages()) {
                  if (hdr.messageId === messageId) { msgHdr = hdr; break; }
                }
                if (!msgHdr) {
                  return { error: `Message not found: ${messageId}` };
                }

                const msgComposeService = Cc["@mozilla.org/messengercompose;1"].getService(Ci.nsIMsgComposeService);
                const msgComposeParams = Cc["@mozilla.org/messengercompose/composeparams;1"].createInstance(Ci.nsIMsgComposeParams);
                const composeFields = Cc["@mozilla.org/messengercompose/composefields;1"].createInstance(Ci.nsIMsgCompFields);

                msgComposeParams.type = replyAll ? Ci.nsIMsgCompType.ReplyAll : Ci.nsIMsgCompType.Reply;
                msgComposeParams.format = Ci.nsIMsgCompFormat.HTML;
                msgComposeParams.composeFields = composeFields;
                msgComposeParams.identity = getIdentityForFolder(folder);
                msgComposeParams.originalMsgURI = msgHdr.folder.getUriForMsg(msgHdr);

                msgComposeService.OpenComposeWindowWithParams(null, msgComposeParams);

                return { success: true, message: "Native reply compose opened" };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            async function replyToMessageDraft(messageId, folderPath, body, replyAll, isHtml, idempotencyKey, includeQuotedOriginal = true, useClosePromptSave = false) {
              try {
                const folder = MailServices.folderLookup.getFolderForURL(folderPath);
                if (!folder) {
                  return { error: `Folder not found: ${folderPath}` };
                }

                const db = folder.msgDatabase;
                if (!db) {
                  return { error: "Could not access folder database" };
                }

                let msgHdr = null;
                for (const hdr of db.enumerateMessages()) {
                  if (hdr.messageId === messageId) {
                    msgHdr = hdr;
                    break;
                  }
                }

                if (!msgHdr) {
                  return { error: `Message not found: ${messageId}` };
                }

                const msgComposeService = Cc["@mozilla.org/messengercompose;1"]
                  .getService(Ci.nsIMsgComposeService);

                const msgComposeParams = Cc["@mozilla.org/messengercompose/composeparams;1"]
                  .createInstance(Ci.nsIMsgComposeParams);

                const composeFields = Cc["@mozilla.org/messengercompose/composefields;1"]
                  .createInstance(Ci.nsIMsgCompFields);

                if (replyAll) {
                  composeFields.to = msgHdr.author;
                  const otherRecipients = (msgHdr.recipients || "").split(",")
                    .map(r => r.trim())
                    .filter(r => r && !r.includes(folder.server.username));
                  if (otherRecipients.length > 0) {
                    composeFields.cc = otherRecipients.join(", ");
                  }
                } else {
                  composeFields.to = msgHdr.author;
                }

                const origSubject = msgHdr.subject || "";
                composeFields.subject = /^\s*re:/i.test(origSubject) ? origSubject : `Re: ${origSubject}`;
                composeFields.references = `<${messageId}>`;

                if (isHtml) {
                  let bodyText = (body || "").replace(/\n/g, '');
                  bodyText = [...bodyText].map(c => c.codePointAt(0) > 127 ? `&#${c.codePointAt(0)};` : c).join('');
                  composeFields.body = `<html><head><meta charset="UTF-8"></head><body>${bodyText}</body></html>`;
                } else {
                  const htmlBody = (body || "")
                    .replace(/&/g, '&amp;')
                    .replace(/</g, '&lt;')
                    .replace(/>/g, '&gt;')
                    .replace(/\n/g, '<br>');
                  composeFields.body = `<html><body>${htmlBody}</body></html>`;
                }

                msgComposeParams.type = Ci.nsIMsgCompType.New;
                msgComposeParams.format = Ci.nsIMsgCompFormat.HTML;
                msgComposeParams.composeFields = composeFields;

                const identity = getIdentityForFolder(folder);
                if (identity) {
                  msgComposeParams.identity = identity;
                }

                const draftsURI = _getDraftsFolderURIForIdentity(identity);
                if (!draftsURI) {
                  return { error: "Could not determine Drafts folder URI for identity" };
                }

                const draftsFolder = MailServices.folderLookup.getFolderForURL(draftsURI);
                if (!draftsFolder) {
                  return { error: `Drafts folder not found: ${draftsURI}` };
                }

                // If requested, mimic human workflow:
                // 1) open real Reply compose window (so TB generates quote + proper metadata)
                // 2) insert our reply text at top
                // 3) close window and auto-click "Save" on the prompt
                if (useClosePromptSave) {
                  const key = idempotencyKey || `reply-${_simpleHash32Hex(`${msgHdr.messageId}|${body}`)}`;
                  const draftMessageId = `tb-mcp-draft-${_sanitizeMessageIdToken(key)}@${_getEmailDomain(identity)}`;

                  if (_pendingDraftMessageIds.has(draftMessageId)) {
                    return { success: true, message: "Reply draft already pending (idempotent)", messageId, folderPath, draftsFolder: draftsURI, draftMessageId };
                  }
                  const existing = _findDraftByMessageId(draftsFolder, draftMessageId);
                  if (existing) {
                    return { success: true, message: "Reply draft already exists (idempotent)", messageId, folderPath, draftsFolder: draftsURI, draftMessageId };
                  }
                  _pendingDraftMessageIds.add(draftMessageId);

                  // Watch for the save-draft confirmation dialog and accept it.
                  const dialogObserver = {
                    observe(subjectWin, topic) {
                      if (topic !== "domwindowopened") return;
                      const win = subjectWin;
                      win.addEventListener(
                        "load",
                        () => {
                          try {
                            // The save-on-close prompt is usually a commonDialog, but can vary.
                            // We just look for a <dialog> with a Save button and click it.

                            // Click the button whose label contains "Save" OR whose accessKey is "S".
                            // (The prompt shows Discard/Cancel/Save with S underlined.)
                            try {
                              const dlg = win.document.querySelector("dialog");
                              const tryButtons = ["accept", "extra1", "extra2", "cancel"];
                              let clicked = false;

                              const clickBtnIfSave = (btn) => {
                                if (!btn) return false;
                                const label = (btn.getAttribute("label") || btn.label || btn.textContent || "").trim();
                                const access = (btn.getAttribute("accesskey") || btn.accessKey || "").trim();
                                if (/^s$/i.test(access) || /save/i.test(label)) {
                                  btn.click();
                                  return true;
                                }
                                return false;
                              };

                              if (dlg && typeof dlg.getButton === "function") {
                                for (const which of tryButtons) {
                                  const btn = dlg.getButton(which);
                                  if (clickBtnIfSave(btn)) {
                                    clicked = true;
                                    break;
                                  }
                                }
                              }

                              if (!clicked) {
                                // Fallback: scan all buttons
                                const buttons = Array.from(win.document.querySelectorAll("button"));
                                for (const btn of buttons) {
                                  if (clickBtnIfSave(btn)) {
                                    clicked = true;
                                    break;
                                  }
                                }
                              }

                              if (!clicked) {
                                // Last resort: synthesize the 'S' key (Save access key).
                                try {
                                  win.focus();
                                } catch {}
                                try {
                                  const wu = win.windowUtils || win
                                    .QueryInterface(Ci.nsIInterfaceRequestor)
                                    .getInterface(Ci.nsIDOMWindowUtils);
                                  if (wu && typeof wu.sendKeyEvent === "function") {
                                    // DOM_VK_S = 83
                                    wu.sendKeyEvent("keydown", 83, 0, 0);
                                    wu.sendKeyEvent("keyup", 83, 0, 0);
                                    clicked = true;
                                  }
                                } catch {}
                              }
                            } catch {
                              // Do nothing; better to not accidentally discard.
                            }
                          } catch {}
                        },
                        { once: true }
                      );
                    },
                  };
                  Services.ww.registerNotification(dialogObserver);

                  const msgComposeService = Cc["@mozilla.org/messengercompose;1"].getService(Ci.nsIMsgComposeService);
                  const msgComposeParams = Cc["@mozilla.org/messengercompose/composeparams;1"].createInstance(Ci.nsIMsgComposeParams);
                  const cf = Cc["@mozilla.org/messengercompose/composefields;1"].createInstance(Ci.nsIMsgCompFields);

                  // Pre-fill body BEFORE opening so the compose window starts with our text.
                  // Thunderbird should still generate the quoted original below (like manual replies).
                  try {
                    const safe = String(body || "")
                      .replace(/&/g, "&amp;")
                      .replace(/</g, "&lt;")
                      .replace(/>/g, "&gt;")
                      .replace(/\n/g, "<br>");
                    cf.body = `<!DOCTYPE html><html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\"></head><body><p>${safe}</p><p><br></p></body></html>`;
                  } catch {}

                  // Let Thunderbird generate recipients/quote by using Reply/ReplyAll.
                  msgComposeParams.type = replyAll ? Ci.nsIMsgCompType.ReplyAll : Ci.nsIMsgCompType.Reply;
                  msgComposeParams.format = Ci.nsIMsgCompFormat.HTML;
                  msgComposeParams.composeFields = cf;
                  if (identity) msgComposeParams.identity = identity;

                  // Set the original msgHdr so TB knows what we're replying to.
                  msgComposeParams.originalMsgURI = msgHdr.folder.getUriForMsg(msgHdr);

                  const composeObserver = {
                    observe(subjectWin, topic) {
                      if (topic !== "domwindowopened") return;
                      const win = subjectWin;

                      let didRun = false;
                      let attempts = 0;

                      const schedule = (ms, fn) => {
                        const t = Cc["@mozilla.org/timer;1"].createInstance(Ci.nsITimer);
                        _pendingTimers.add(t);
                        t.init({ notify: () => { try { _pendingTimers.delete(t); } catch {} fn(); } }, ms, Ci.nsITimer.TYPE_ONE_SHOT);
                      };

                      const tryRun = () => {
                        if (didRun) return;
                        attempts += 1;

                        let isCompose = false;
                        try {
                          isCompose = String(win.location).includes("messengercompose");
                        } catch {}

                        if (!isCompose) {
                          if (attempts < 60) schedule(200, tryRun);
                          return;
                        }

                        // Now run exactly once.
                        didRun = true;

                        const runnable = {
                          run: () => {
                            const poll = {
                              tries: 0,
                              run: () => {
                                poll.tries++;
                                let quoteReady = false;
                                try {
                                  quoteReady = !!win.document.querySelector("blockquote[type='cite'], .moz-cite-prefix, #divRplyFwdMsg");
                                } catch {}

                                if (!quoteReady && poll.tries < 20) {
                                  schedule(300, () => Services.tm.dispatchToMainThread(poll));
                                  return;
                                }

                                // Paste + close using Thunderbird commands (more reliable than synthetic key events).
                                try {
                                  const fm = Cc["@mozilla.org/focus-manager;1"].getService(Ci.nsIFocusManager);
                                  try { win.focus(); } catch {}
                                  try { fm.activeWindow = win; } catch {}
                                } catch {}

                                _setClipboardText(String(body || "") + "\n\n");

                                // cmd_paste should paste into the currently-focused editor.
                                try {
                                  if (typeof win.goDoCommand === "function") {
                                    win.goDoCommand("cmd_paste");
                                  }
                                } catch {}

                                schedule(800, () => {
                                  try {
                                    if (typeof win.goDoCommand === "function") {
                                      win.goDoCommand("cmd_close");
                                    }
                                  } catch {}
                                  try { if (!win.closed) win.close(); } catch {}

                                  try { Services.ww.unregisterNotification(composeObserver); } catch {}

                                  // Keep dialog observer alive a bit longer.
                                  schedule(25000, () => {
                                    try { Services.ww.unregisterNotification(dialogObserver); } catch {}
                                    try { _pendingDraftMessageIds.delete(draftMessageId); } catch {}
                                  });
                                });
                              },
                            };

                            Services.tm.dispatchToMainThread(poll);
                          },
                        };

                        Services.tm.dispatchToMainThread(runnable);
                      };

                      // Try immediately + repeated attempts (independent of load events).
                      tryRun();
                      schedule(1000, tryRun);
                      schedule(3000, tryRun);
                      schedule(6000, tryRun);

                      // Also attach listeners (best-effort).
                      try { win.addEventListener("DOMContentLoaded", tryRun, { once: true }); } catch {}
                      try { win.addEventListener("load", tryRun, { once: true }); } catch {}
                    },
                  };

                  Services.ww.registerNotification(composeObserver);
                  msgComposeService.OpenComposeWindowWithParams(null, msgComposeParams);

                  return { success: true, message: "Reply compose opened; will auto-close and save draft via prompt", messageId, folderPath, draftsFolder: draftsURI, draftMessageId };
                }

                // Default behavior (existing path): build HTML quote ourselves and SaveAsDraft.
                let finalBody = body || "";
                let finalBodyHtml = "";

                if (includeQuotedOriginal) {
                  try {
                    const orig = await getMessage(messageId, folderPath);
                    if (orig && !orig.error) {
                      const dateStr = orig.date ? new Date(orig.date).toLocaleString("en-US") : "";
                      const author = orig.author || "";
                      const citeMid = orig.id ? `mid:${orig.id}` : "";
                      const origHtml = orig.bodyHtml || "";

                      const replyHtml = `<p>${String(body || "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/\n/g,"<br>")}</p>`;
                      const citePrefix = `<div class=\"moz-cite-prefix\">On ${dateStr}, ${author} wrote:<br></div>`;
                      const quoted = origHtml
                        ? `<blockquote type=\"cite\" cite=\"${citeMid}\">${origHtml}</blockquote>`
                        : `<blockquote type=\"cite\" cite=\"${citeMid}\"><pre>${String(orig.bodyText || orig.body || "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}</pre></blockquote>`;

                      finalBodyHtml = `<!DOCTYPE html><html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"></head><body>${replyHtml}${citePrefix}${quoted}</body></html>`;
                    }
                  } catch {}
                }

                const key = idempotencyKey || `reply-${_simpleHash32Hex(`${msgHdr.messageId}|${composeFields.to}|${composeFields.subject}|${finalBodyHtml || finalBody}`)}`;
                const draftMessageId = `tb-mcp-draft-${_sanitizeMessageIdToken(key)}@${_getEmailDomain(identity)}`;

                if (_pendingDraftMessageIds.has(draftMessageId)) {
                  return { success: true, message: "Reply draft already pending (idempotent)", messageId, folderPath, draftsFolder: draftsURI, draftMessageId };
                }

                const existing = _findDraftByMessageId(draftsFolder, draftMessageId);
                if (existing) {
                  return { success: true, message: "Reply draft already exists (idempotent)", messageId, folderPath, draftsFolder: draftsURI, draftMessageId };
                }

                _pendingDraftMessageIds.add(draftMessageId);
                const bodyHtml = finalBodyHtml || `<!DOCTYPE html><html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"></head><body><p>${String(body || "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/\n/g,"<br>")}</p></body></html>`;

                const saveRes = await _saveDraftViaComposeWindow({
                  identity,
                  to: composeFields.to || "",
                  cc: composeFields.cc || "",
                  subject: /^\s*re:/i.test(origSubject) ? origSubject : `Re: ${origSubject}`,
                  bodyHtml,
                });

                try { draftsFolder.updateFolder(null); } catch {}
                try { _pendingDraftMessageIds.delete(draftMessageId); } catch {}

                return { success: !!saveRes.ok, message: "Reply draft saved (native SaveAsDraft)", messageId, folderPath, draftsFolder: draftsURI, draftMessageId, saveRes };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function debugContext() {
              try {
                const info = {
                  hasExtBrowser: !!extBrowser,
                  extBrowserType: extBrowser ? typeof extBrowser : null,
                  contextKeys: Object.getOwnPropertyNames(context || {}),
                  messageManagerType: (context && context.messageManager) ? Object.prototype.toString.call(context.messageManager) : null,
                  messageManagerKeys: (context && context.messageManager) ? Object.getOwnPropertyNames(context.messageManager).slice(0,80) : [],
                  messageManagerProxyType: (context && context.messageManagerProxy) ? Object.prototype.toString.call(context.messageManagerProxy) : null,
                  messageManagerProxyKeys: (context && context.messageManagerProxy) ? Object.getOwnPropertyNames(context.messageManagerProxy).slice(0,80) : [],
                  mmFromProxyType: (context && context.messageManagerProxy && context.messageManagerProxy.messageManager) ? Object.prototype.toString.call(context.messageManagerProxy.messageManager) : null,
                  mmFromProxyKeys: (context && context.messageManagerProxy && context.messageManagerProxy.messageManager) ? Object.getOwnPropertyNames(context.messageManagerProxy.messageManager).slice(0,80) : [],
                  hasExtension: !!(context && context.extension),
                  extensionKeys: context && context.extension ? Object.getOwnPropertyNames(context.extension) : [],
                  hasApiManager: !!(context && context.extension && context.extension.apiManager),
                  apiManagerKeys: context && context.extension && context.extension.apiManager ? Object.getOwnPropertyNames(context.extension.apiManager) : [],
                  hasBgFrameLoader: !!(context && context.extension && context.extension._backgroundPageFrameLoader),
                  bgFrameLoaderKeys: (context && context.extension && context.extension._backgroundPageFrameLoader) ? Object.getOwnPropertyNames(context.extension._backgroundPageFrameLoader).slice(0,50) : [],
                  viewsType: (context && context.extension && context.extension.views) ? Object.prototype.toString.call(context.extension.views) : null,
                  viewsCtor: (context && context.extension && context.extension.views && context.extension.views.constructor) ? context.extension.views.constructor.name : null,
                  viewsIsArray: !!(context && context.extension && Array.isArray(context.extension.views)),
                  viewsCount: (context && context.extension && Array.isArray(context.extension.views)) ? context.extension.views.length : null,
                  viewsKeys: (context && context.extension && context.extension.views) ? Object.getOwnPropertyNames(context.extension.views).slice(0,50) : [],
                };

                // Probe a few known paths.
                const probes = {};
                const paths = [
                  "context.cloneScope.browser",
                  "context.extension.apiManager.global.browser",
                  "context.extension.backgroundPage",
                  "context.extension.views",
                ];

                // Inspect extension.views (Set of ExtensionView instances)
                try {
                  const views = (context && context.extension) ? context.extension.views : null;
                  if (views && typeof views[Symbol.iterator] === "function") {
                    const arr = [];
                    let i = 0;
                    for (const v of views) {
                      if (i++ >= 8) break;
                      arr.push({
                        viewType: v && v.viewType,
                        hasXulBrowser: !!(v && v.xulBrowser),
                        hasContentWindow: !!(v && v.xulBrowser && v.xulBrowser.contentWindow),
                        contentWindowHasBrowser: !!(v && v.xulBrowser && v.xulBrowser.contentWindow && v.xulBrowser.contentWindow.browser),
                        contentWindowKeys: (v && v.xulBrowser && v.xulBrowser.contentWindow) ? Object.getOwnPropertyNames(v.xulBrowser.contentWindow).slice(0, 20) : [],
                      });
                    }
                    info.viewsSample = arr;
                  }
                } catch (e) {
                  info.viewsSampleError = String(e);
                }

                for (const p of paths) {
                  try {
                    let v;
                    if (p === "context.cloneScope.browser") v = context.cloneScope && context.cloneScope.browser;
                    if (p === "context.extension.apiManager.global.browser") v = context.extension && context.extension.apiManager && context.extension.apiManager.global && context.extension.apiManager.global.browser;
                    if (p === "context.extension.backgroundPage") v = context.extension && context.extension.backgroundPage;
                    if (p === "context.extension.views") v = context.extension && context.extension.views;
                    probes[p] = v ? (typeof v) : null;
                  } catch (e) {
                    probes[p] = `ERR:${e}`;
                  }
                }
                info.probes = probes;

                // Try to see if context.apiCan can resolve WebExtension APIs.
                try {
                  info.hasApiCan = !!context.apiCan;
                  info.apiCanKeys = context.apiCan ? Object.getOwnPropertyNames(context.apiCan).slice(0,50) : [];
                  if (context.apiCan && typeof context.apiCan.findAPIPath === "function") {
                    // This may return an object for known namespaces.
                    info.apiCanComposeType = typeof context.apiCan.findAPIPath("compose");
                    info.apiCanMessagesType = typeof context.apiCan.findAPIPath("messages");
                  }
                } catch (e) {
                  info.apiCanError = String(e);
                }

                return { ok: true, info };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            async function replyToMessageDraftComposeApi(messageIdHeader, replyAll, plainTextBody, htmlBody, includeQuotedOriginal = true, closeAfterSave = true) {
              try {
                // Prefer using context.apiCan to access WebExtension namespaces from the experiment context.
                if (!context || !context.apiCan || typeof context.apiCan.findAPIPath !== "function") {
                  return { error: "WebExtension API container (context.apiCan) is not available" };
                }

                const messagesApi = context.apiCan.findAPIPath("messages");
                const composeApi = context.apiCan.findAPIPath("compose");
                const tabsApi = context.apiCan.findAPIPath("tabs");

                if (!messagesApi || !composeApi) {
                  return { error: "Could not resolve messages/compose API via context.apiCan" };
                }

                // Resolve the WebExtension numeric messageId from a RFC822 Message-ID header.
                // Workaround: some Thunderbird builds throw if fromDate is undefined.
                const queryInfo = { headerMessageId: messageIdHeader, fromDate: new Date(0), toDate: new Date() };
                const queryRes = await messagesApi.query(queryInfo);
                const messages = (queryRes && Array.isArray(queryRes.messages)) ? queryRes.messages : [];
                if (messages.length === 0) {
                  return { error: `messages.query() returned no results for headerMessageId=${messageIdHeader}` };
                }

                const msg = messages[0];
                const msgId = msg.id;

                const replyType = replyAll ? "replyToAll" : "replyToSender";
                const tab = await composeApi.beginReply(msgId, replyType);
                const tabId = tab && tab.id;

                const details = await composeApi.getComposeDetails(tabId);
                const existingBody = (details && typeof details.body === "string") ? details.body : "";

                let prefixHtml = "";
                if (typeof htmlBody === "string" && htmlBody.trim()) {
                  prefixHtml = htmlBody;
                } else if (typeof plainTextBody === "string" && plainTextBody.length) {
                  const esc = (s) => String(s)
                    .replace(/&/g, "&amp;")
                    .replace(/</g, "&lt;")
                    .replace(/>/g, "&gt;");
                  prefixHtml = `<p>${esc(plainTextBody).replace(/\r\n/g, "\n").replace(/\n/g, "<br>")}</p>`;
                }

                const newBody = includeQuotedOriginal ? (prefixHtml + existingBody) : prefixHtml;
                await composeApi.setComposeDetails(tabId, { body: newBody });

                await composeApi.saveMessage(tabId, { mode: "draft" });

                if (closeAfterSave && tabsApi && typeof tabsApi.remove === "function") {
                  try { await tabsApi.remove(tabId); } catch {}
                }

                return {
                  ok: true,
                  method: "compose-api (apiCan)",
                  headerMessageId: messageIdHeader,
                  resolvedMessageId: msgId,
                  tabId,
                  replyType,
                  includeQuotedOriginal,
                  closeAfterSave,
                };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function listLatestMessages(folderPath, limit) {
              try {
                const folder = MailServices.folderLookup.getFolderForURL(folderPath);
                if (!folder) {
                  return { error: `Folder not found: ${folderPath}` };
                }

                const db = folder.msgDatabase;
                if (!db) {
                  return { error: "Could not access folder database" };
                }

                const max = Math.min(Math.max(parseInt(limit || 20, 10) || 20, 1), 100);
                const msgs = [];

                // Collect all headers then sort by date descending.
                const all = [];
                for (const hdr of db.enumerateMessages()) {
                  all.push(hdr);
                }
                all.sort((a, b) => (b.date || 0) - (a.date || 0));

                for (const hdr of all.slice(0, max)) {
                  msgs.push({
                    id: hdr.messageId,
                    subject: hdr.mime2DecodedSubject || hdr.subject,
                    author: hdr.mime2DecodedAuthor || hdr.author,
                    recipients: hdr.mime2DecodedRecipients || hdr.recipients,
                    date: hdr.date ? new Date(hdr.date / 1000).toISOString() : null,
                    folder: folder.prettyName,
                    folderPath: folder.URI,
                    read: hdr.isRead,
                    flagged: hdr.isFlagged,
                  });
                }

                return { ok: true, folderPath: folder.URI, count: msgs.length, items: msgs };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function _getMessageHdrsByIds(folder, ids) {
              const db = folder.msgDatabase;
              if (!db) return [];
              const out = [];
              for (const id of ids) {
                let found = null;
                for (const hdr of db.enumerateMessages()) {
                  if (hdr.messageId === id) {
                    found = hdr;
                    break;
                  }
                }
                if (found) out.push(found);
              }
              return out;
            }

            function moveMessages(fromFolderPath, toFolderPath, messageIds) {
              return new Promise((resolve) => {
                try {
                  const fromFolder = MailServices.folderLookup.getFolderForURL(fromFolderPath);
                  const toFolder = MailServices.folderLookup.getFolderForURL(toFolderPath);
                  if (!fromFolder) {
                    resolve({ error: `Folder not found: ${fromFolderPath}` });
                    return;
                  }
                  if (!toFolder) {
                    resolve({ error: `Folder not found: ${toFolderPath}` });
                    return;
                  }

                  const ids = Array.isArray(messageIds) ? messageIds : [];
                  const hdrs = _getMessageHdrsByIds(fromFolder, ids);
                  if (hdrs.length === 0) {
                    resolve({ success: true, moved: 0, message: "No matching messages found" });
                    return;
                  }

                  const copyService = MailServices.copy;
                  const listener = {
                    QueryInterface: ChromeUtils.generateQI([Ci.nsIMsgCopyServiceListener]),
                    OnStartCopy() {},
                    OnProgress() {},
                    SetMessageKey() {},
                    GetMessageId() {},
                    OnStopCopy(status) {
                      if (status && !Components.isSuccessCode(status)) {
                        resolve({ error: `Move failed: ${status}` });
                        return;
                      }
                      try { toFolder.updateFolder(null); } catch {}
                      resolve({ success: true, moved: hdrs.length, fromFolderPath: fromFolder.URI, toFolderPath: toFolder.URI });
                    },
                  };

                  const copyFn = copyService.CopyMessages || copyService.copyMessages;
                  if (typeof copyFn !== "function") {
                    resolve({ error: "Copy service missing CopyMessages/copyMessages" });
                    return;
                  }

                  // isMove=true
                  copyFn.call(copyService, fromFolder, hdrs, toFolder, true, listener, null, false);
                } catch (e) {
                  resolve({ error: e.toString() });
                }
              });
            }

            function deleteMessages(folderPath, messageIds) {
              try {
                const folder = MailServices.folderLookup.getFolderForURL(folderPath);
                if (!folder) {
                  return { error: `Folder not found: ${folderPath}` };
                }

                const ids = Array.isArray(messageIds) ? messageIds : [];

                // Preference: Draft deletions should be a move to Deleted Items.
                if (/\/Drafts$/i.test(folder.URI) || /\/Drafts\b/i.test(folder.URI)) {
                  const deletedItemsURI = folder.URI.replace(/\/Drafts$/i, "/Deleted Items");
                  // Best-effort move; if the destination doesn't exist, fall back to delete.
                  try {
                    const deletedFolder = MailServices.folderLookup.getFolderForURL(deletedItemsURI);
                    if (deletedFolder) {
                      return moveMessages(folder.URI, deletedItemsURI, ids);
                    }
                  } catch {}
                }

                const hdrs = _getMessageHdrsByIds(folder, ids);
                if (hdrs.length === 0) {
                  return { success: true, deleted: 0, message: "No matching messages found" };
                }

                // Try common signatures across TB versions.
                try {
                  folder.deleteMessages(hdrs, null, true, false, null, false);
                } catch {
                  try {
                    folder.deleteMessages(hdrs, null, true, false, null);
                  } catch (e2) {
                    return { error: `deleteMessages failed: ${e2.toString()}` };
                  }
                }

                return { success: true, deleted: hdrs.length, folderPath: folder.URI };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function getRawMessage(messageId, folderPath) {
              return new Promise((resolve) => {
                try {
                  const folder = MailServices.folderLookup.getFolderForURL(folderPath);
                  if (!folder) {
                    resolve({ error: `Folder not found: ${folderPath}` });
                    return;
                  }

                  const db = folder.msgDatabase;
                  if (!db) {
                    resolve({ error: "Could not access folder database" });
                    return;
                  }

                  let msgHdr = null;
                  for (const hdr of db.enumerateMessages()) {
                    if (hdr.messageId === messageId) {
                      msgHdr = hdr;
                      break;
                    }
                  }

                  if (!msgHdr) {
                    resolve({ error: `Message not found: ${messageId}` });
                    return;
                  }

                  const msgService = MailServices.messageServiceFromURI(msgHdr.folder.getUriForMsg(msgHdr));
                  const uri = msgHdr.folder.getUriForMsg(msgHdr);

                  let chunks = [];
                  const listener = {
                    QueryInterface: ChromeUtils.generateQI([Ci.nsIStreamListener, Ci.nsIRequestObserver]),
                    onStartRequest() {},
                    onDataAvailable(request, inputStream, offset, count) {
                      try {
                        const sis = Cc["@mozilla.org/scriptableinputstream;1"].createInstance(Ci.nsIScriptableInputStream);
                        sis.init(inputStream);
                        chunks.push(sis.read(count));
                      } catch (e) {
                        // ignore
                      }
                    },
                    onStopRequest(request, status) {
                      try {
                        if (status && !Components.isSuccessCode(status)) {
                          resolve({ error: `streamMessage failed: ${status}` });
                          return;
                        }
                        resolve({ ok: true, messageId, folderPath, source: sanitizeForJson(chunks.join("")) });
                      } catch (e) {
                        resolve({ error: e.toString() });
                      }
                    },
                  };

                  msgService.streamMessage(uri, listener, null, null, false, "", false);
                } catch (e) {
                  resolve({ error: e.toString() });
                }
              });
            }

            async function callTool(name, args) {
              switch (name) {
                case "searchMessages":
                  return searchMessages(args.query || "");
                case "getMessage":
                  return await getMessage(args.messageId, args.folderPath);
                case "getLatestUnread":
                  return await getLatestUnread(args.folderPath);
                case "getLatestUnreadBatch":
                  return await getLatestUnreadBatch(args.folderPath, args.limit);
                case "setMessageRead":
                  return setMessageRead(args.messageId, args.folderPath, args.read);
                case "searchContacts":
                  return searchContacts(args.query || "");
                case "listCalendars":
                  return listCalendars();
                case "sendMail":
                  return composeMail(args.to, args.subject, args.body, args.cc, args.isHtml);
                case "saveDraft":
                  return saveDraft(args.to, args.subject, args.body, args.cc, args.isHtml, args.idempotencyKey);
                case "replyToMessage":
                  return replyToMessage(args.messageId, args.folderPath, args.body, args.replyAll, args.isHtml);
                case "openNativeReplyCompose":
                  return openNativeReplyCompose(args.messageId, args.folderPath, args.replyAll);
                case "replyToMessageDraft":
                  return replyToMessageDraft(args.messageId, args.folderPath, args.body, args.replyAll, args.isHtml, args.idempotencyKey, args.includeQuotedOriginal, args.useClosePromptSave);
                case "replyToMessageDraftComposeApi":
                  return await replyToMessageDraftComposeApi(args.messageId, args.replyAll, args.plainTextBody, args.htmlBody, args.includeQuotedOriginal, args.closeAfterSave);
                case "debugContext":
                  return debugContext();
                case "listLatestMessages":
                  return listLatestMessages(args.folderPath, args.limit);
                case "deleteMessages":
                  return await deleteMessages(args.folderPath, args.messageIds);
                case "moveMessages":
                  return await moveMessages(args.fromFolderPath, args.toFolderPath, args.messageIds);
                case "getRawMessage":
                  return await getRawMessage(args.messageId, args.folderPath);
                default:
                  throw new Error(`Unknown tool: ${name}`);
              }
            }

            const server = new HttpServer();

            server.registerPathHandler("/", (req, res) => {
              res.processAsync();

              if (req.method !== "POST") {
                res.setStatusLine("1.1", 405, "Method Not Allowed");
                res.write("POST only");
                res.finish();
                return;
              }

              let message;
              try {
                message = JSON.parse(readRequestBody(req));
              } catch {
                res.setStatusLine("1.1", 400, "Bad Request");
                res.write("Invalid JSON");
                res.finish();
                return;
              }

              const { id, method, params } = message;

              (async () => {
                try {
                  let result;
                  switch (method) {
                    case "tools/list":
                      result = { tools };
                      break;
                    case "tools/call":
                      if (!params?.name) {
                        throw new Error("Missing tool name");
                      }
                      result = {
                        content: [{
                          type: "text",
                          text: JSON.stringify(await callTool(params.name, params.arguments || {}), null, 2)
                        }]
                      };
                      break;
                    default:
                      res.setStatusLine("1.1", 404, "Not Found");
                      res.write(`Unknown method: ${method}`);
                      res.finish();
                      return;
                  }
                  res.setStatusLine("1.1", 200, "OK");
                  // charset=utf-8 is critical for proper emoji handling in responses
                  res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                  res.write(JSON.stringify({ jsonrpc: "2.0", id, result }));
                } catch (e) {
                  res.setStatusLine("1.1", 200, "OK");
                  res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                  res.write(JSON.stringify({
                    jsonrpc: "2.0",
                    id,
                    error: { code: -32000, message: e.toString() }
                  }));
                }
                res.finish();
              })();
            });

            server.start(MCP_PORT);
            console.log(`Thunderbird MCP server listening on port ${MCP_PORT}`);
            return { success: true, port: MCP_PORT };
          } catch (e) {
            console.error("Failed to start MCP server:", e);
            return { success: false, error: e.toString() };
          }
        }
      }
    };
  }

  onShutdown(isAppShutdown) {
    if (isAppShutdown) return;
    resProto.setSubstitution("thunderbird-mcp", null);
    Services.obs.notifyObservers(null, "startupcache-invalidate");
  }
};
