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
            isHtml: { type: "boolean", description: "Set to true if body contains HTML markup (default: false)" },
          },
          required: ["to", "subject", "body"],
        },
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
            isHtml: { type: "boolean", description: "Set to true if body contains HTML markup (default: false)" },
          },
          required: ["messageId", "folderPath", "body"],
        },
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

                // Thunderbird message header API
                try {
                  msgHdr.markRead(!!read);
                } catch (e) {
                  return { error: `Failed to mark read: ${e.toString()}` };
                }

                return { success: true, messageId, folderPath, read: !!read };
              } catch (e) {
                return { error: e.toString() };
              }
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
                composeFields.subject = origSubject.startsWith("Re:") ? origSubject : `Re: ${origSubject}`;

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

                const account = MailServices.accounts.findAccountForServer(folder.server);
                if (account) {
                  msgComposeParams.identity = account.defaultIdentity;
                }

                msgComposeService.OpenComposeWindowWithParams(null, msgComposeParams);

                return { success: true, message: "Reply window opened" };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            async function callTool(name, args) {
              switch (name) {
                case "searchMessages":
                  return searchMessages(args.query || "");
                case "getMessage":
                  return await getMessage(args.messageId, args.folderPath);
                case "getLatestUnread":
                  return await getLatestUnread(args.folderPath);
                case "setMessageRead":
                  return setMessageRead(args.messageId, args.folderPath, args.read);
                case "searchContacts":
                  return searchContacts(args.query || "");
                case "listCalendars":
                  return listCalendars();
                case "sendMail":
                  return composeMail(args.to, args.subject, args.body, args.cc, args.isHtml);
                case "replyToMessage":
                  return replyToMessage(args.messageId, args.folderPath, args.body, args.replyAll, args.isHtml);
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
