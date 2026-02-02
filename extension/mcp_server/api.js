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
const DEFAULT_MAX_RESULTS = 50;

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
        description: "Search for email messages in Thunderbird. Returns up to maxResults messages (default 50) sorted by date (newest first by default). Each result includes: id (message ID for use with getMessage), subject, author, recipients, date (ISO 8601), folder, folderPath (for use with getMessage), read status, and flagged status.",
        inputSchema: {
          type: "object",
          properties: {
            query: {
              type: "string",
              description: "Text to search for in message subject, author, or recipients. Use empty string to match all messages (useful with date filters)."
            },
            startDate: {
              type: "string",
              description: "Filter messages on or after this date. ISO 8601 format (e.g., '2024-01-15' or '2024-01-15T00:00:00Z'). If omitted, no start date filter is applied."
            },
            endDate: {
              type: "string",
              description: "Filter messages on or before this date. ISO 8601 format (e.g., '2024-01-31' or '2024-01-31T23:59:59Z'). If omitted, no end date filter is applied."
            },
            maxResults: {
              type: "number",
              description: "Maximum number of messages to return (default: 50, max: 200). Results are sorted by date before limiting."
            },
            sortOrder: {
              type: "string",
              enum: ["desc", "asc"],
              description: "Sort order by date: 'desc' for newest first (default), 'asc' for oldest first."
            }
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
        name: "listAccounts",
        title: "List Email Accounts",
        description: "List configured email accounts (key, name, type, email, username, hostName)",
        inputSchema: { type: "object", properties: {}, required: [] },
      },
      {
        name: "listFolders",
        title: "List Folders",
        description: "List folders for an account (or all accounts) with message/unread counts",
        inputSchema: {
          type: "object",
          properties: {
            accountKey: { type: "string", description: "Optional account key from listAccounts" }
          },
          required: [],
        },
      },
      {
        name: "getRecentMessages",
        title: "Get Recent Messages",
        description: "Get recent messages sorted by date (newest first), optionally unread-only and/or folder-scoped",
        inputSchema: {
          type: "object",
          properties: {
            folderPath: { type: "string", description: "Optional folder URI (imap://.../INBOX). If omitted, uses all Inbox folders." },
            limit: { type: "number", description: "Max messages (default 20, max 100)" },
            daysBack: { type: "number", description: "Days back to include (default 30)" },
            unreadOnly: { type: "boolean", description: "Only unread (default false)" }
          },
          required: [],
        },
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
        name: "replyToMessageDraftNativeEditor",
        title: "Reply to Message (Draft via native editor insert)",
        description: "Open a native Reply compose (tab or window), insert the reply body using Thunderbird's editor APIs (no OS-level UI automation), then save as draft.",
        inputSchema: {
          type: "object",
          properties: {
            messageId: { type: "string", description: "RFC822 Message-ID header" },
            folderPath: { type: "string", description: "Folder URI (imap://.../INBOX)" },
            replyAll: { type: "boolean", description: "Reply to all recipients (default false)" },
            plainTextBody: { type: "string", description: "Reply body to insert at top (plain text)." },
            closeAfterSave: { type: "boolean", description: "Close compose tab/window after saving (default true)" }
          },
          required: ["messageId", "folderPath", "plainTextBody"]
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
        name: "reviseDraftInPlaceNativeEditor",
        title: "Revise Draft In Place (native editor)",
        description: "Open an existing draft in a native compose tab/window and revise it using Thunderbird editor APIs, then save. Supports preserving the quoted original in reply drafts.",
        inputSchema: {
          type: "object",
          properties: {
            messageId: { type: "string", description: "RFC822 Message-ID header of the existing draft" },
            folderPath: { type: "string", description: "Folder URI where the draft lives (imap://.../Drafts)" },
            plainTextBody: { type: "string", description: "Replacement body text (plain text). If preserveQuotedOriginal=true, this replaces only the top text above the quote." },
            preserveQuotedOriginal: { type: "boolean", description: "If true, keep the existing quoted original (blockquote cite) and replace only the text above it (default true)." },
            closeAfterSave: { type: "boolean", description: "Close compose tab/window after saving (default true)" }
          },
          required: ["messageId", "folderPath", "plainTextBody"]
        }
      },
      {
        name: "replyToMessageDraftComposeApi",
        title: "Reply to Message (Save Draft via compose API)",
        description: "Create a reply draft using Thunderbird's WebExtension compose API (messages.query + compose.beginReply + compose.setComposeDetails + compose.saveMessage). This should preserve native reply quoting and produce a real Draft without OS-level UI automation.",
        inputSchema: {
          type: "object",
          properties: {
            messageId: { type: "string", description: "RFC822 Message-ID header (same as other tools; e.g. 4f23...@...)." },
            folderPath: { type: "string", description: "Folder URI where the message currently lives (e.g. imap://.../INBOX). Used to resolve the WebExtension numeric message id without relying on global search." },
            replyAll: { type: "boolean", description: "Reply to all recipients (default: false)" },
            plainTextBody: { type: "string", description: "Reply body as plain text to insert ABOVE the quoted original. Newlines will be preserved." },
            htmlBody: { type: "string", description: "Reply body as HTML to insert ABOVE the quoted original." },
            includeQuotedOriginal: { type: "boolean", description: "Whether to keep the quoted original at bottom (default: true)." },
            closeAfterSave: { type: "boolean", description: "Close the compose tab/window after saving (default: true)." }
          },
          required: ["messageId", "folderPath"]
        }
      },
      {
        name: "debugContext",
        title: "Debug Context",
        description: "Return debugging info about the experiment context and any reachable WebExtension browser globals.",
        inputSchema: { type: "object", properties: {}, required: [] }
      },
      {
        name: "debugMessagesList",
        title: "Debug Messages List",
        description: "Debug helper: resolve folderId from folderPath and list the first N messages via WebExtension messages.list (shows id + headerMessageId + subject).",
        inputSchema: {
          type: "object",
          properties: {
            folderPath: { type: "string" },
            limit: { type: "number" }
          },
          required: ["folderPath"]
        }
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
          // Guard against multiple concurrent start attempts (can happen on TB reloads / multiple entrypoints).
          // If we try to bind twice, we get NS_ERROR_SOCKET_ADDRESS_IN_USE and the caller may hang.
          try {
            if (globalThis.__tbMcpStartPromise) {
              return await globalThis.__tbMcpStartPromise;
            }
          } catch {}

          globalThis.__tbMcpStartPromise = (async () => {
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
             * Also handles UTF-8 encoding since Thunderbird's HTTP server writes raw bytes.
             */
            function sanitizeForJson(text) {
              if (!text) return text;
              // Loop-based approach (regex unreliable in Thunderbird's JS engine)
              let result = "";
              for (let i = 0; i < text.length; i++) {
                const c = text.charCodeAt(i);
                // Skip control chars (except \t, \n, \r), quotes, and backslash
                if ((c >= 0x00 && c <= 0x08) || c === 0x0B || c === 0x0C ||
                    (c >= 0x0E && c <= 0x1F) || c === 0x7F ||
                    c === 0x22 || c === 0x5C) {
                  continue;
                }
                // UTF-8 encode non-ASCII (Thunderbird HTTP server doesn't)
                if (c >= 128 && c <= 2047) {
                  result += String.fromCharCode(0xC0 | (c >> 6), 0x80 | (c & 0x3F));
                } else if (c >= 2048 && c <= 65535) {
                  result += String.fromCharCode(0xE0 | (c >> 12), 0x80 | ((c >> 6) & 0x3F), 0x80 | (c & 0x3F));
                } else {
                  result += text[i];
                }
              }
              return result;
            }

            /**
             * Search for messages with optional date filtering and sorting.
             *
             * @param {string} query - Text to search for (empty string matches all)
             * @param {string} startDate - ISO 8601 date string for start of range (optional)
             * @param {string} endDate - ISO 8601 date string for end of range (optional)
             * @param {number} maxResults - Maximum results to return (default: 50, max: 200)
             * @param {string} sortOrder - 'desc' for newest first (default), 'asc' for oldest first
             * @returns {Array} Array of message objects sorted by date
             */
            function searchMessages(query, startDate, endDate, maxResults, sortOrder) {
              const results = [];
              const lowerQuery = query ? query.toLowerCase() : "";
              const matchAll = !query || query.trim() === "";

              // Tokenized matching: if query has multiple tokens, require all tokens to match
              // across subject/author/recipients (order-insensitive). This makes searches like
              // "Jiawei Ge Princeton SDS Seminar" work even though the subject contains punctuation.
              const tokens = (lowerQuery || "")
                .split(/[^a-z0-9@._-]+/i)
                .map(t => t.trim())
                .filter(Boolean);
              const useTokenMatch = tokens.length >= 2;

              function tokenMatch(haystack) {
                if (!useTokenMatch) return false;
                const h = String(haystack || "");
                return tokens.every(t => h.includes(t));
              }
              
              // Parse date filters
              let startTimestamp = null;
              let endTimestamp = null;
              
              if (startDate) {
                const parsed = Date.parse(startDate);
                if (!isNaN(parsed)) {
                  // Thunderbird stores dates in microseconds
                  startTimestamp = parsed * 1000;
                }
              }
              
              if (endDate) {
                const parsed = Date.parse(endDate);
                if (!isNaN(parsed)) {
                  // Add 1 day to include the end date fully (end of day)
                  endTimestamp = (parsed + 86400000) * 1000;
                }
              }
              
              // Validate and cap maxResults
              const effectiveMaxResults = Math.min(
                Math.max(1, maxResults || DEFAULT_MAX_RESULTS),
                200
              );
              
              // Default to descending (newest first)
              const ascending = sortOrder === "asc";
              
              // Collect all matching messages first (we need to sort them)
              const allMatches = [];

              function searchFolder(folder) {
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
                    // Apply date filters first (most efficient)
                    const msgDate = msgHdr.date;
                    
                    if (startTimestamp !== null && msgDate < startTimestamp) {
                      continue;
                    }
                    
                    if (endTimestamp !== null && msgDate > endTimestamp) {
                      continue;
                    }

                    // IMPORTANT: Use mime2Decoded* properties for searching.
                    // Raw headers contain MIME encoding like "=?UTF-8?Q?...?="
                    // which won't match plain text searches.
                    const subject = (msgHdr.mime2DecodedSubject || msgHdr.subject || "").toLowerCase();
                    const author = (msgHdr.mime2DecodedAuthor || msgHdr.author || "").toLowerCase();
                    const recipients = (msgHdr.mime2DecodedRecipients || msgHdr.recipients || "").toLowerCase();

                    const combined = `${subject} ${author} ${recipients}`;

                    if (matchAll ||
                        (useTokenMatch ? tokenMatch(combined) : (
                          subject.includes(lowerQuery) ||
                          author.includes(lowerQuery) ||
                          recipients.includes(lowerQuery)
                        ))) {
                      allMatches.push({
                        id: msgHdr.messageId,
                        subject: msgHdr.mime2DecodedSubject || msgHdr.subject,
                        author: msgHdr.mime2DecodedAuthor || msgHdr.author,
                        recipients: msgHdr.mime2DecodedRecipients || msgHdr.recipients,
                        date: msgDate ? new Date(msgDate / 1000).toISOString() : null,
                        dateTimestamp: msgDate || 0,
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
                    searchFolder(subfolder);
                  }
                }
              }

              for (const account of MailServices.accounts.accounts) {
                searchFolder(account.incomingServer.rootFolder);
              }

              // Sort by date
              allMatches.sort((a, b) => {
                if (ascending) {
                  return a.dateTimestamp - b.dateTimestamp;
                } else {
                  return b.dateTimestamp - a.dateTimestamp;
                }
              });

              // Take top N results and remove internal timestamp field
              const finalResults = allMatches.slice(0, effectiveMaxResults).map(msg => {
                const { dateTimestamp, ...rest } = msg;
                return rest;
              });

              return finalResults;
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


            function listAccounts() {
              try {
                const accounts = [];
                for (const account of MailServices.accounts.accounts) {
                  const server = account.incomingServer;
                  accounts.push({
                    key: account.key,
                    name: account.name || (server && server.prettyName) || "",
                    type: (server && server.type) || "",
                    email: (account.defaultIdentity && account.defaultIdentity.email) ? account.defaultIdentity.email : "",
                    username: (server && server.username) || "",
                    hostName: (server && server.hostName) || "",
                  });
                }
                return accounts;
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function listFolders(accountKey) {
              try {
                const folders = [];

                function folderType(flags) {
                  return (flags & 0x00001000) ? "inbox" :
                         (flags & 0x00000200) ? "sent" :
                         (flags & 0x00000400) ? "drafts" :
                         (flags & 0x00000100) ? "trash" :
                         (flags & 0x00000800) ? "templates" : "folder";
                }

                function addFolderAndSubfolders(folder, acctInfo) {
                  folders.push({
                    name: folder.prettyName,
                    path: folder.URI,
                    type: folderType(folder.flags),
                    accountKey: acctInfo.key,
                    accountName: acctInfo.name,
                    totalMessages: folder.getTotalMessages(false),
                    unreadMessages: folder.getNumUnread(false)
                  });

                  if (folder.hasSubFolders) {
                    for (const sub of folder.subFolders) {
                      addFolderAndSubfolders(sub, acctInfo);
                    }
                  }
                }

                if (accountKey) {
                  const account = MailServices.accounts.getAccount(accountKey);
                  if (!account) return { error: `Account not found: ${accountKey}` };
                  addFolderAndSubfolders(account.incomingServer.rootFolder, { key: account.key, name: account.name });
                } else {
                  for (const account of MailServices.accounts.accounts) {
                    addFolderAndSubfolders(account.incomingServer.rootFolder, { key: account.key, name: account.name });
                  }
                }

                return folders;
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function getRecentMessages(folderPath, limit, daysBack, unreadOnly) {
              try {
                const maxLimit = Math.min(Math.max(parseInt(limit || 20, 10) || 20, 1), 100);
                const days = Math.max(parseInt(daysBack || 30, 10) || 30, 1);
                const cutoffMs = Date.now() - days * 24 * 60 * 60 * 1000;
                const all = [];

                function collect(folder) {
                  try {
                    try { if (folder.server && folder.server.type === "imap") folder.updateFolder(null); } catch {}
                    const db = folder.msgDatabase;
                    if (!db) return;
                    for (const msgHdr of db.enumerateMessages()) {
                      const msgDateMs = (msgHdr.date || 0) / 1000;
                      if (msgDateMs < cutoffMs) continue;
                      if (unreadOnly && msgHdr.isRead) continue;
                      all.push({
                        id: msgHdr.messageId,
                        subject: msgHdr.mime2DecodedSubject || msgHdr.subject,
                        author: msgHdr.mime2DecodedAuthor || msgHdr.author,
                        recipients: msgHdr.mime2DecodedRecipients || msgHdr.recipients,
                        date: msgHdr.date ? new Date(msgHdr.date / 1000).toISOString() : null,
                        _dateTs: msgHdr.date || 0,
                        folder: folder.prettyName,
                        folderPath: folder.URI,
                        read: msgHdr.isRead,
                        flagged: msgHdr.isFlagged,
                      });
                    }
                  } catch {}
                }

                if (folderPath) {
                  const folder = MailServices.folderLookup.getFolderForURL(folderPath);
                  if (!folder) return { error: `Folder not found: ${folderPath}` };
                  collect(folder);
                } else {
                  for (const account of MailServices.accounts.accounts) {
                    try {
                      const root = account.incomingServer.rootFolder;
                      const inbox = root.getFolderWithFlags(Ci.nsMsgFolderFlags.Inbox);
                      if (inbox) collect(inbox);
                    } catch {}
                  }
                }

                all.sort((a, b) => (b._dateTs || 0) - (a._dateTs || 0));
                const out = all.slice(0, maxLimit).map(({ _dateTs, ...rest }) => rest);
                return { ok: true, count: out.length, results: out };
              } catch (e) {
                return { error: e.toString() };
              }
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

                  // Combine original recipients and CC list, filter out own address.
                  const allRecipients = [
                    ...(msgHdr.recipients || "").split(","),
                    ...(msgHdr.ccList || "").split(",")
                  ]
                    .map(r => r.trim())
                    .filter(r => r && !r.includes(folder.server.username));

                  // Deduplicate by email address.
                  const seen = new Set();
                  const uniqueRecipients = allRecipients.filter(r => {
                    const email = r.match(/<([^>]+)>/)?.[1]?.toLowerCase() || r.toLowerCase();
                    if (seen.has(email)) return false;
                    seen.add(email);
                    return true;
                  });

                  if (uniqueRecipients.length > 0) {
                    composeFields.cc = uniqueRecipients.join(", ");
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

            function _insertTextAtTopOfCompose(win, text) {
              // Best-effort insertion at the beginning of the message body.
              // This is intentionally defensive across Thunderbird versions.
              try {
                if (!win) return false;
                const t = String(text || "");

                // Try to get the editor.
                let editor = null;
                try {
                  if (win.gMsgCompose && win.gMsgCompose.editor) {
                    editor = win.gMsgCompose.editor;
                  }
                } catch {}
                try {
                  if (!editor && typeof win.GetCurrentEditor === "function") {
                    editor = win.GetCurrentEditor();
                  }
                } catch {}

                if (!editor) return false;

                // Move caret to start and insert with line breaks.
                try {
                  if (typeof editor.beginningOfDocument === "function") {
                    editor.beginningOfDocument();
                  }
                } catch {}

                // Insert as text. In HTML editor, \n will typically become <br>.
                try {
                  if (typeof editor.insertText === "function") {
                    editor.insertText(t + "\n\n");
                    return true;
                  }
                } catch {}

                // Fallback: use nsIHTMLEditor / nsIPlaintextEditor interfaces.
                try {
                  const plain = editor.QueryInterface(Ci.nsIPlaintextEditor);
                  plain.insertText(t + "\n\n");
                  return true;
                } catch {}

                return false;
              } catch {
                return false;
              }
            }

            function _replaceBodyWithPlainText(win, text) {
              // Replace the entire compose body with the provided plain text.
              try {
                if (!win) return false;
                const t = String(text || "");

                let editor = null;
                try { if (win.gMsgCompose && win.gMsgCompose.editor) editor = win.gMsgCompose.editor; } catch {}
                try { if (!editor && typeof win.GetCurrentEditor === "function") editor = win.GetCurrentEditor(); } catch {}
                if (!editor) return false;

                // Select all + delete.
                try { if (typeof editor.selectAll === "function") editor.selectAll(); } catch {}
                try {
                  if (typeof editor.deleteSelection === "function") {
                    editor.deleteSelection("next", "strip");
                  }
                } catch {}

                // Insert new content.
                try {
                  if (typeof editor.insertText === "function") {
                    editor.insertText(t + "\n");
                    return true;
                  }
                } catch {}

                try {
                  const plain = editor.QueryInterface(Ci.nsIPlaintextEditor);
                  plain.insertText(t + "\n");
                  return true;
                } catch {}

                return false;
              } catch {
                return false;
              }
            }

            function _replaceTopTextPreserveQuote(win, text) {
              // Replace only the top part of the body, keeping the quoted original (first blockquote[type=cite]).
              // IMPORTANT: use editor APIs / selection transactions so Thunderbird persists the change to the saved draft.
              try {
                if (!win) return false;
                const t = String(text || "");

                let editor = null;
                try { if (win.gMsgCompose && win.gMsgCompose.editor) editor = win.gMsgCompose.editor; } catch {}
                try { if (!editor && typeof win.GetCurrentEditor === "function") editor = win.GetCurrentEditor(); } catch {}
                if (!editor) return false;

                const doc = editor.document || (win.document || null);
                if (!doc) return _replaceBodyWithPlainText(win, text);

                const body = doc.body || doc.documentElement;
                if (!body) return _replaceBodyWithPlainText(win, text);

                let quote = null;
                try {
                  quote = doc.querySelector && doc.querySelector('blockquote[type="cite"], blockquote[cite], blockquote');
                } catch {}
                if (!quote) return _replaceBodyWithPlainText(win, text);

                // Select everything from start of body up to (but not including) the quote.
                let sel = null;
                try { sel = editor.selection; } catch {}
                try { if (!sel && typeof win.getSelection === "function") sel = win.getSelection(); } catch {}
                if (!sel) return _replaceBodyWithPlainText(win, text);

                try { if (typeof editor.beginTransaction === "function") editor.beginTransaction(); } catch {}

                try {
                  const r = doc.createRange();
                  r.setStart(body, 0);
                  r.setEndBefore(quote);

                  try { sel.removeAllRanges(); } catch {}
                  try { sel.addRange(r); } catch {}

                  // Delete the selected top portion.
                  try {
                    if (typeof editor.deleteSelection === "function") {
                      editor.deleteSelection("next", "strip");
                    }
                  } catch {}

                  // Move caret to start of body.
                  try {
                    const r2 = doc.createRange();
                    r2.setStart(body, 0);
                    r2.collapse(true);
                    sel.removeAllRanges();
                    sel.addRange(r2);
                  } catch {}

                  // Insert replacement text. Newlines will generally map to <br> or paragraphs depending on compose mode.
                  try {
                    if (typeof editor.insertText === "function") {
                      editor.insertText(t + "\n\n");
                      return true;
                    }
                  } catch {}

                  try {
                    const plain = editor.QueryInterface(Ci.nsIPlaintextEditor);
                    plain.insertText(t + "\n\n");
                    return true;
                  } catch {}

                  return false;
                } finally {
                  try { if (typeof editor.endTransaction === "function") editor.endTransaction(); } catch {}
                }
              } catch {
                return false;
              }
            }

            // Keep timers strongly referenced; otherwise they can be GC'd before firing.
            const _activeTimers = new Set();
            function _sleep(ms) {
              return new Promise((resolve) => {
                try {
                  const t = Cc["@mozilla.org/timer;1"].createInstance(Ci.nsITimer);
                  _activeTimers.add(t);
                  t.initWithCallback({
                    notify: () => {
                      try { _activeTimers.delete(t); } catch {}
                      resolve();
                    }
                  }, ms, Ci.nsITimer.TYPE_ONE_SHOT);
                } catch {
                  Services.tm.dispatchToMainThread({ run: () => resolve() });
                }
              });
            }

            function _withTimeout(promise, timeoutMs, label) {
              return Promise.race([
                promise,
                (async () => {
                  await _sleep(timeoutMs);
                  throw new Error(`Timeout: ${label || "operation"} (${timeoutMs}ms)`);
                })(),
              ]);
            }

            function _snapshotComposeTabs() {
              // Returns a set of compose tab identifiers + light metadata.
              const out = [];
              try {
                const en = Services.wm.getEnumerator("mail:3pane");
                while (en.hasMoreElements()) {
                  const w = en.getNext();
                  const tabmail = w.document && w.document.getElementById && w.document.getElementById("tabmail");
                  if (!tabmail || !Array.isArray(tabmail.tabInfo)) continue;
                  for (const ti of tabmail.tabInfo) {
                    try {
                      const browser = ti.browser || ti.linkedBrowser;
                      const cw = browser && browser.contentWindow;
                      const href = cw && cw.location && cw.location.href;
                      if (!href || !String(href).includes("messengercompose")) continue;
                      const subject = (cw && cw.gMsgCompose && cw.gMsgCompose.compFields && cw.gMsgCompose.compFields.subject) || "";
                      out.push({ w, tabmail, ti, cw, href: String(href), subject: String(subject || "") });
                    } catch {}
                  }
                }
              } catch {}
              return out;
            }

            function _pickNewComposeTab(before, after, wantSubject) {
              // Heuristic: choose a compose tab that wasn't present in the 'before' snapshot.
              const beforeSet = new Set((before || []).map(x => String(x.href) + "|" + String(x.subject)));
              const candidates = (after || []).filter(x => !beforeSet.has(String(x.href) + "|" + String(x.subject)));
              const all = candidates.length ? candidates : (after || []);
              if (wantSubject) {
                const exact = all.find(x => (x.subject || "") === wantSubject);
                if (exact) return exact;
              }
              return all[0] || null;
            }

            async function reviseDraftInPlaceNativeEditor(messageId, folderPath, plainTextBody, preserveQuotedOriginal = true, closeAfterSave = true) {
              // Open an existing draft using the WebExtension compose API (beginNew(messageId)),
              // then overwrite the body using native editor APIs and save.
              // If TB creates a new draft item on save, delete the old one to avoid duplicates.
              try {
                if (!context || !context.apiCan || typeof context.apiCan.findAPIPath !== "function") {
                  return { error: "WebExtension API container (context.apiCan) is not available" };
                }
                const accountsApi = context.apiCan.findAPIPath("accounts");
                const messagesApi = context.apiCan.findAPIPath("messages");
                const composeApi = context.apiCan.findAPIPath("compose");
                const tabsApi = context.apiCan.findAPIPath("tabs");
                if (!accountsApi || !messagesApi || !composeApi) {
                  return { error: "Could not resolve accounts/messages/compose API via context.apiCan" };
                }

                function parseFolderUri(uri) {
                  const m = String(uri || "").match(/^imap:\/\/(.+?)@([^\/]+)\/(.+)$/i);
                  if (!m) return null;
                  const user = decodeURIComponent(m[1]);
                  const folderPathPart = "/" + m[3].replace(/^\/+/, "");
                  return { user, folderPathPart };
                }
                function findFolderByPath(folders, wantedPath) {
                  if (!Array.isArray(folders)) return null;
                  for (const f of folders) {
                    if (!f) continue;
                    if (f.path === wantedPath) return f;
                    const sub = findFolderByPath(f.subFolders || f.folders || f.subfolders, wantedPath);
                    if (sub) return sub;
                  }
                  return null;
                }

                const parsed = parseFolderUri(folderPath);
                if (!parsed) return { error: `Could not parse folderPath URI: ${folderPath}` };

                const accounts = await accountsApi.list();
                let folder = null;
                let desiredIdentityId = null;
                for (const acct of accounts || []) {
                  const match = (acct.identities || []).find(i => i && i.email && i.email.toLowerCase() === parsed.user.toLowerCase());
                  if (!match) continue;
                  desiredIdentityId = match.id || null;
                  folder = findFolderByPath(acct.folders, parsed.folderPathPart);
                  if (folder) break;
                }
                if (!folder || !folder.id) return { error: `Could not resolve folder for ${folderPath}` };

                // Resolve numeric message id for this draft.
                let msgId = null;
                let chunk = await messagesApi.list(folder.id);
                const listId = chunk && chunk.id;
                let safety = 0;
                while (chunk && safety++ < 500) {
                  for (const m of (chunk.messages || [])) {
                    if (m && m.headerMessageId === messageId) { msgId = m.id; break; }
                  }
                  if (msgId) break;
                  if (!chunk.id) break;
                  try { chunk = await messagesApi.continueList(chunk.id); } catch { break; }
                }
                try { if (listId) await messagesApi.abortList(listId); } catch {}
                if (!msgId) return { error: `Draft not found in folder by headerMessageId: ${messageId}` };

                const startedAt = Date.now();

                // Open the draft for editing.
                const tab = await _withTimeout(composeApi.beginNew(msgId), 15000, "compose.beginNew(draft)");
                const tabId = tab && tab.id;
                if (!tabId) return { error: "compose.beginNew did not return a tabId" };

                // Ensure identity.
                if (desiredIdentityId) {
                  try { await _withTimeout(composeApi.setComposeDetails(tabId, { identityId: desiredIdentityId }), 5000, "compose.setComposeDetails(identityId)"); } catch {}
                }

                // Find the compose window via TabManager, retrying a bit.
                let cw = null;
                for (let i = 0; i < 60; i++) {
                  try {
                    const tm = context && context.extension && context.extension.tabManager;
                    if (tm && typeof tm.get === "function") {
                      const extTab = tm.get(tabId);
                      if (extTab && extTab.nativeTab) cw = extTab.nativeTab;
                    }
                  } catch {}
                  if (cw && cw.gMsgCompose) break;
                  await _sleep(250);
                }

                // Fallback: scan mail:3pane compose tabs.
                if (!cw || !cw.gMsgCompose) {
                  for (let i = 0; i < 60; i++) {
                    const after = _snapshotComposeTabs();
                    const picked = (after || []).find(x => {
                      try { return x && x.cw && x.cw.gMsgCompose && x.cw.gMsgCompose.editor; } catch { return false; }
                    }) || null;
                    if (picked) { cw = picked.cw; break; }
                    await _sleep(250);
                  }
                }

                if (!cw || !cw.gMsgCompose) return { error: "Timeout locating compose editor for draft" };

                // Prefer WebExtension compose APIs for actually persisting the change into the draft.
                // For reply-drafts, preserve the quoted original by keeping the existing HTML quote block intact.
                let replaced = false;
                try {
                  if (preserveQuotedOriginal) {
                    const details = await _withTimeout(composeApi.getComposeDetails(tabId), 8000, "compose.getComposeDetails");
                    const existingBody = (details && typeof details.body === "string") ? details.body : "";

                    function escHtml(s) {
                      return String(s)
                        .replace(/&/g, "&amp;")
                        .replace(/</g, "&lt;")
                        .replace(/>/g, "&gt;")
                        .replace(/\"/g, "&quot;");
                    }
                    const lines = String(plainTextBody || "").split(/\r?\n/).map(escHtml);
                    const topHtml = `<p>${lines.join("<br>\n")}</p>\n<br>\n<br>\n`;

                    // Split at the moz-cite-prefix (preferred) or the first blockquote.
                    let idx = -1;
                    idx = existingBody.indexOf('<div class="moz-cite-prefix"');
                    if (idx < 0) idx = existingBody.search(/<blockquote[^>]*(type=\"cite\"|type='cite'|type=cite|cite=)/i);
                    if (idx < 0) idx = existingBody.search(/<blockquote/i);

                    const remainder = (idx >= 0) ? existingBody.slice(idx) : "";
                    const newBody = remainder ? (topHtml + remainder) : topHtml;
                    await _withTimeout(composeApi.setComposeDetails(tabId, { body: newBody }), 15000, "compose.setComposeDetails(body)");
                    replaced = true;
                  } else {
                    // Full replace (no quote preservation).
                    await _withTimeout(composeApi.setComposeDetails(tabId, { plainTextBody: String(plainTextBody || "") }), 15000, "compose.setComposeDetails(plainTextBody)");
                    replaced = true;
                  }
                } catch {
                  // Fallback to editor manipulation (best-effort).
                  replaced = preserveQuotedOriginal
                    ? _replaceTopTextPreserveQuote(cw, plainTextBody)
                    : _replaceBodyWithPlainText(cw, plainTextBody);
                }

                function _linesFromText(t) {
                  try {
                    const lines = String(t || "")
                      .split(/\r?\n/)
                      .map(x => String(x || "").trim())
                      .filter(Boolean);

                    // Always keep token-like lines even if short.
                    const tok = lines.filter(x => /\bTOK\d?\b|\bTOK[0-9-]/.test(x));
                    const normal = lines
                      // drop super-short filler lines ("Hi", etc.) unless token.
                      .filter(x => x.length >= 5)
                      .slice(0, 20);

                    const merged = [];
                    for (const x of tok.concat(normal)) {
                      if (!merged.includes(x)) merged.push(x);
                      if (merged.length >= 20) break;
                    }
                    return merged;
                  } catch {
                    return [];
                  }
                }

                // Snapshot Drafts BEFORE save (newest N by date) so we can diff deterministically.
                const wantLines = _linesFromText(plainTextBody);
                const wantHasToken = wantLines.some(x => /\bTOK\d?\b|\bTOK[0-9-]/.test(x));
                const before = { ids: [], items: [] };
                try {
                  const b = listLatestMessages(folderPath, 15);
                  before.items = (b && b.items) ? b.items.map(x => ({ id: x.id, date: x.date, subject: x.subject })) : [];
                  before.ids = before.items.map(x => x.id);
                } catch {}

                // Save via WebExtension compose API.
                // If available, use compose.onAfterSave to learn the saved message-id deterministically.
                let afterSaveInfo = null;
                let afterSaveErr = null;
                let removeAfterSave = null;
                try {
                  const evt = composeApi && composeApi.onAfterSave;
                  if (evt && typeof evt.addListener === "function") {
                    const handler = (savedTab, info) => {
                      try {
                        const id = savedTab && (savedTab.id || savedTab.tabId);
                        if (id !== tabId) return;
                        afterSaveInfo = info || { ok: true };
                      } catch (e) {
                        afterSaveErr = String(e);
                      }
                    };
                    evt.addListener(handler);
                    removeAfterSave = () => { try { evt.removeListener(handler); } catch {} };
                  }
                } catch {}

                let saveRes = null;
                let onAfterSaveAvailable = false;
                try {
                  const evt = composeApi && composeApi.onAfterSave;
                  onAfterSaveAvailable = !!(evt && typeof evt.addListener === "function");
                } catch {}

                try { saveRes = await _withTimeout(composeApi.saveMessage(tabId, { mode: "draft" }), 30000, "compose.saveMessage(draft)"); } catch (e) { afterSaveErr = String(e); }

                // Give onAfterSave a moment to arrive.
                for (let i = 0; i < 40 && !afterSaveInfo; i++) {
                  await _sleep(250);
                }
                try { if (removeAfterSave) removeAfterSave(); } catch {}

                await _sleep(1000);

                // Try to refresh the folder DB after saving (IMAP Drafts can lag).
                try {
                  const folderNative = MailServices.folderLookup.getFolderForURL(folderPath);
                  if (folderNative && typeof folderNative.updateFolder === "function") {
                    folderNative.updateFolder(null);
                  }
                } catch {}

                var __debugDetectDraft = null;

                // Detect the *actual* saved draft.
                let deletedOld = false;
                let newMessageId = messageId;

                const debugDetect = {
                  wantLines,
                  wantHasToken,
                  before,
                  onAfterSaveAvailable,
                  saveRes,
                  afterSaveInfo,
                  afterSaveErr,
                  attempts: 0,
                  checked: [],
                  picked: null,
                  errors: [],
                };

                // Best case: compose.saveMessage returned the saved draft message(s).
                try {
                  const m = saveRes && saveRes.messages && saveRes.messages[0];
                  const hid = m && (m.headerMessageId || m.headerMessageID || m.headerMessageId);
                  if (hid) {
                    newMessageId = hid;
                    debugDetect.picked = { id: newMessageId, via: "compose.saveMessage.return.messages[0].headerMessageId" };
                  }
                } catch {}
                __debugDetectDraft = debugDetect;

                try {
                  // If saveMessage already told us the id, we're done (avoid extra scanning).
                  if (debugDetect.picked && debugDetect.picked.id) {
                    // no-op
                  } else {
                    // Poll a few times; find the saved draft by walking recent items via messages API pagination.
                    // Rationale: msgDatabase ordering can be stale/weird for IMAP Drafts; and messages.list() order may not
                    // include the new draft in the first page. So we page a bit and filter by time window + subject.
                    const startedAtMs = Date.now();
                    let wantSubj = "";
                    try { wantSubj = (cw.gMsgCompose && cw.gMsgCompose.compFields && cw.gMsgCompose.compFields.subject) || ""; } catch {}
                    debugDetect.wantSubject = wantSubj;

                    for (let attempt = 0; attempt < 12; attempt++) {
                    debugDetect.attempts++;

                    const cutoff = startedAtMs - 10 * 60 * 1000; // 10 minutes
                    let candidates = [];
                    let listId = null;
                    let pages = 0;
                    try {
                      let chunk = await _withTimeout(messagesApi.list(folder.id), 10000, "messages.list(drafts after)");
                      listId = chunk && chunk.id;
                      while (chunk && pages++ < 8) {
                        const msgs = (chunk.messages || []);
                        for (const m of msgs) {
                          try {
                            if (!m || !m.headerMessageId) continue;
                            if (wantSubj && (m.subject || "") !== wantSubj) continue;
                            if (m.date && m.date < cutoff) continue;
                            candidates.push({ id: m.headerMessageId, date: m.date || null, subject: m.subject || null });
                          } catch {}
                        }
                        if (!chunk.id) break;
                        chunk = await _withTimeout(messagesApi.continueList(chunk.id), 10000, "messages.continueList(drafts after)");
                      }
                    } catch (e) {
                      debugDetect.errors.push({ stage: "messagesListPaged", error: String(e) });
                    } finally {
                      try { if (listId) await messagesApi.abortList(listId); } catch {}
                    }

                    debugDetect.pages = pages;
                    debugDetect.candidateCount = candidates.length;

                    // If subject filtering returned nothing, fall back to sampling newest regardless of subject.
                    if (!candidates.length) {
                      try {
                        const a = listLatestMessages(folderPath, 40);
                        candidates = (a && a.items) ? a.items.map(x => ({ id: x.id, date: x.date, subject: x.subject })) : [];
                      } catch (e) {
                        debugDetect.errors.push({ stage: "fallbackListLatest", error: String(e) });
                        candidates = [];
                      }
                    }

                    const beforeSet = new Set(before.ids || []);
                    // Prefer new ids first.
                    const newOnes = candidates.filter(x => x && x.id && !beforeSet.has(x.id));
                    const scanList = (newOnes.length ? newOnes : candidates).slice(0, 25);

                    let best = null;
                    let fetched = 0;
                    const maxFetch = 18;
                    for (const c of scanList) {
                      if (!c || !c.id) continue;
                      if (fetched++ >= maxFetch) break;
                      try {
                        const rawObj = getRawMessage(c.id, folderPath);
                        const src = rawObj && rawObj.source;
                        const body = (src && src.includes("\r\n\r\n")) ? src.split("\r\n\r\n", 2)[1] : (src || "");
                        let score = 0;
                        for (const ln of wantLines) {
                          if (body.includes(ln)) score++;
                        }
                        const isOld = (c.id === messageId);
                        debugDetect.checked.push({ id: c.id, isOld, score, subject: c.subject || null, date: c.date || null, isNew: !beforeSet.has(c.id) });

                        const threshold = wantHasToken ? 1 : 2;
                        if (score >= threshold) {
                          best = { id: c.id, score, isNew: !beforeSet.has(c.id) };
                          break;
                        }
                      } catch (e) {
                        debugDetect.checked.push({ id: c.id, error: true, err: String(e), isNew: !beforeSet.has(c.id) });
                      }
                    }

                    if (best && best.id) {
                      newMessageId = best.id;
                      debugDetect.picked = best;
                      break;
                    }

                    await _sleep(1200);
                  }
                  }
                } catch (e) {
                  debugDetect.errors.push({ stage: "detectLoop", error: String(e) });
                }

                // Best-effort delete old duplicate if we found a different id.
                if (newMessageId && newMessageId !== messageId) {
                  try {
                    const folderNative = MailServices.folderLookup.getFolderForURL(folderPath);
                    const db = folderNative && folderNative.msgDatabase;
                    let origHdr = null;
                    if (db) {
                      for (const hdr of db.enumerateMessages()) {
                        if (hdr && hdr.messageId === messageId) { origHdr = hdr; break; }
                      }
                    }
                    if (origHdr && origHdr.folder) {
                      origHdr.folder.deleteMessages([origHdr], null, true, false, null, false);
                      deletedOld = true;
                    }
                  } catch {}
                }

                if (closeAfterSave) {
                  try {
                    if (tabsApi && typeof tabsApi.remove === "function") {
                      await tabsApi.remove(tabId);
                    }
                  } catch {}
                }

                const debug = { buildStamp: "reviseDraftInPlaceNativeEditor-debug-2026-02-02T20:14Z" };
                try {
                  if (typeof __debugDetectDraft !== "undefined") debug.detectSavedDraft = __debugDetectDraft;
                } catch {}

                return { ok: true, replaced, oldMessageId: messageId, messageId: newMessageId, deletedOld, tabId, debug };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            async function replyToMessageDraftNativeEditor(messageId, folderPath, replyAll, plainTextBody, closeAfterSave = true) {
              try {
                // Open reply using native WebExtension compose API (so TB creates the quote/threading).
                if (!context || !context.apiCan || typeof context.apiCan.findAPIPath !== "function") {
                  return { error: "WebExtension API container (context.apiCan) is not available" };
                }
                const accountsApi = context.apiCan.findAPIPath("accounts");
                const messagesApi = context.apiCan.findAPIPath("messages");
                const composeApi = context.apiCan.findAPIPath("compose");
                const tabsApi = context.apiCan.findAPIPath("tabs");
                if (!accountsApi || !messagesApi || !composeApi) {
                  return { error: "Could not resolve accounts/messages/compose API via context.apiCan" };
                }

                // Resolve message id (numeric) from folderPath + RFC822 Message-ID header.
                function parseFolderUri(uri) {
                  const m = String(uri || "").match(/^imap:\/\/(.+?)@([^\/]+)\/(.+)$/i);
                  if (!m) return null;
                  const user = decodeURIComponent(m[1]);
                  const folderPathPart = "/" + m[3].replace(/^\/+/, "");
                  return { user, folderPathPart };
                }
                function findFolderByPath(folders, wantedPath) {
                  if (!Array.isArray(folders)) return null;
                  for (const f of folders) {
                    if (!f) continue;
                    if (f.path === wantedPath) return f;
                    const sub = findFolderByPath(f.subFolders || f.folders || f.subfolders, wantedPath);
                    if (sub) return sub;
                  }
                  return null;
                }
                const parsed = parseFolderUri(folderPath);
                if (!parsed) return { error: `Could not parse folderPath URI: ${folderPath}` };

                const accounts = await accountsApi.list();
                let folder = null;
                let desiredIdentityId = null;
                for (const acct of accounts || []) {
                  const match = (acct.identities || []).find(i => i && i.email && i.email.toLowerCase() === parsed.user.toLowerCase());
                  if (!match) continue;
                  desiredIdentityId = match.id || null;
                  folder = findFolderByPath(acct.folders, parsed.folderPathPart);
                  if (folder) break;
                }
                if (!folder || !folder.id) return { error: `Could not resolve folder for ${folderPath}` };

                let msgId = null;
                let chunk = await messagesApi.list(folder.id);
                const listId = chunk && chunk.id;
                let safety = 0;
                while (chunk && safety++ < 500) {
                  for (const m of (chunk.messages || [])) {
                    if (m && m.headerMessageId === messageId) { msgId = m.id; break; }
                  }
                  if (msgId) break;
                  if (!chunk.id) break;
                  try { chunk = await messagesApi.continueList(chunk.id); } catch { break; }
                }
                try { if (listId) await messagesApi.abortList(listId); } catch {}
                if (!msgId) return { error: `Message not found in folder by headerMessageId: ${messageId}` };

                const replyType = replyAll ? "replyToAll" : "replyToSender";
                const tab = await composeApi.beginReply(msgId, replyType);
                const tabId = tab && tab.id;

                // Ensure the correct identity (account) is selected for the compose tab.
                // Without this, TB can sometimes pick the wrong default identity and save the draft
                // into the wrong account's Drafts.
                if (desiredIdentityId) {
                  try {
                    await composeApi.setComposeDetails(tabId, { identityId: desiredIdentityId });
                  } catch {}
                }

                const sleep = (ms) => new Promise((resolve2) => {
                  try {
                    const t = Cc["@mozilla.org/timer;1"].createInstance(Ci.nsITimer);
                    t.initWithCallback({ notify: () => resolve2() }, ms, Ci.nsITimer.TYPE_ONE_SHOT);
                  } catch {
                    Services.tm.dispatchToMainThread({ run: () => resolve2() });
                  }
                });

                // Prefer resolving the compose tab via the extension's TabManager (most direct path to nativeTab).
                let composeWin = null;
                try {
                  const tm = context && context.extension && context.extension.tabManager;
                  if (tm && typeof tm.get === "function") {
                    const extTab = tm.get(tabId);
                    if (extTab && extTab.nativeTab) {
                      composeWin = extTab.nativeTab;
                    }
                  }
                } catch {}

                // Fallback: scan mail:3pane windows for a compose tab.
                function findComposeContentWindow() {
                  try {
                    const en = Services.wm.getEnumerator("mail:3pane");
                    while (en.hasMoreElements()) {
                      const w = en.getNext();
                      const tabmail = w.document && w.document.getElementById && w.document.getElementById("tabmail");
                      if (!tabmail || !tabmail.tabInfo) continue;
                      for (const ti of tabmail.tabInfo) {
                        const browser = ti.browser || ti.linkedBrowser;
                        const cw = browser && browser.contentWindow;
                        if (!cw) continue;
                        const href = cw.location && cw.location.href;
                        if (href && String(href).includes("messengercompose") && cw.gMsgCompose) {
                          return { w, tabmail, ti, cw };
                        }
                      }
                    }
                  } catch {}
                  return null;
                }

                let found = null;
                for (let i = 0; i < 60; i++) {
                  if (composeWin && composeWin.gMsgCompose) {
                    found = { cw: composeWin, tabmail: null, ti: null };
                    break;
                  }
                  found = findComposeContentWindow();
                  if (found) break;
                  await sleep(250);
                }
                if (!found || !found.cw) {
                  return { error: "Timeout locating compose editor" };
                }

                // Insert + save.
                const inserted = _insertTextAtTopOfCompose(found.cw, plainTextBody);
                try { if (typeof found.cw.goDoCommand === "function") found.cw.goDoCommand("cmd_saveAsDraft"); } catch {}

                // Let save finish.
                await sleep(2500);

                if (closeAfterSave) {
                  try {
                    if (found.tabmail && typeof found.tabmail.closeTab === "function" && found.ti) {
                      found.tabmail.closeTab(found.ti);
                    } else if (tabsApi && typeof tabsApi.remove === "function") {
                      await tabsApi.remove(tabId);
                    }
                  } catch {}
                }

                return { ok: true, tabId, inserted };
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

            async function debugMessagesList(folderPath, limit = 20) {
              try {
                if (!context || !context.apiCan || typeof context.apiCan.findAPIPath !== "function") {
                  return { error: "WebExtension API container (context.apiCan) is not available" };
                }
                const accountsApi = context.apiCan.findAPIPath("accounts");
                const messagesApi = context.apiCan.findAPIPath("messages");
                if (!accountsApi || !messagesApi) {
                  return { error: "Could not resolve accounts/messages API via context.apiCan" };
                }

                function parseFolderUri(uri) {
                  const m = String(uri || "").match(/^imap:\/\/(.+?)@([^\/]+)\/(.+)$/i);
                  if (!m) return null;
                  const user = decodeURIComponent(m[1]);
                  const folderPathPart = "/" + m[3].replace(/^\/+/, "");
                  return { user, folderPathPart };
                }
                function findFolderByPath(folders, wantedPath) {
                  if (!Array.isArray(folders)) return null;
                  for (const f of folders) {
                    if (!f) continue;
                    if (f.path === wantedPath) return f;
                    const sub = findFolderByPath(f.subFolders || f.folders || f.subfolders, wantedPath);
                    if (sub) return sub;
                  }
                  return null;
                }

                const parsed = parseFolderUri(folderPath);
                if (!parsed) return { error: `Could not parse folderPath URI: ${folderPath}` };

                const accounts = await accountsApi.list();
                let folder = null;
                for (const acct of accounts || []) {
                  const match = (acct.identities || []).find(i => i && i.email && i.email.toLowerCase() === parsed.user.toLowerCase());
                  if (!match) continue;
                  folder = findFolderByPath(acct.folders, parsed.folderPathPart);
                  if (folder) break;
                }
                if (!folder || !folder.id) {
                  return { error: `Could not resolve folder for ${folderPath} (wanted ${parsed.folderPathPart})` };
                }

                const res = await messagesApi.list(folder.id);
                const msgs = (res && res.messages) ? res.messages : [];
                const out = msgs.slice(0, Math.max(1, Math.min(100, limit || 20))).map(m => ({
                  id: m.id,
                  headerMessageId: m.headerMessageId,
                  subject: m.subject,
                  date: m.date,
                }));
                try { if (res && res.id) await messagesApi.abortList(res.id); } catch {}

                return { ok: true, folder, sample: out, totalSampled: out.length };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            async function replyToMessageDraftComposeApi(messageIdHeader, folderPath, replyAll, plainTextBody, htmlBody, includeQuotedOriginal = true, closeAfterSave = true) {
              try {
                // Prefer using context.apiCan to access WebExtension namespaces from the experiment context.
                if (!context || !context.apiCan || typeof context.apiCan.findAPIPath !== "function") {
                  return { error: "WebExtension API container (context.apiCan) is not available" };
                }

                const accountsApi = context.apiCan.findAPIPath("accounts");
                const messagesApi = context.apiCan.findAPIPath("messages");
                const composeApi = context.apiCan.findAPIPath("compose");
                const tabsApi = context.apiCan.findAPIPath("tabs");

                if (!accountsApi || !messagesApi || !composeApi) {
                  return { error: "Could not resolve accounts/messages/compose API via context.apiCan" };
                }

                // Resolve a WebExtension folder id from the IMAP folder URI.
                function parseFolderUri(uri) {
                  // Example: imap://jl4624%40cornell.edu@outlook.office365.com/INBOX
                  const m = String(uri || "").match(/^imap:\/\/(.+?)@([^\/]+)\/(.+)$/i);
                  if (!m) return null;
                  const userEnc = m[1];
                  const host = m[2];
                  const path = m[3];
                  const user = decodeURIComponent(userEnc);
                  const folderPathPart = "/" + path.replace(/^\/+/, "");
                  return { user, host, folderPathPart };
                }

                function findFolderByPath(folders, wantedPath) {
                  if (!Array.isArray(folders)) return null;
                  for (const f of folders) {
                    if (!f) continue;
                    if (f.path === wantedPath) return f;
                    const sub = findFolderByPath(f.subFolders || f.folders || f.subfolders, wantedPath);
                    if (sub) return sub;
                  }
                  return null;
                }

                const parsed = parseFolderUri(folderPath);
                if (!parsed) {
                  return { error: `Could not parse folderPath URI: ${folderPath}` };
                }

                const accounts = await accountsApi.list();
                let folderId = null;
                for (const acct of accounts || []) {
                  // Prefer matching identity email to URI user.
                  const ids = acct.identities || [];
                  const match = ids.find(i => i && typeof i.email === "string" && i.email.toLowerCase() === parsed.user.toLowerCase());
                  if (!match) continue;
                  const f = findFolderByPath(acct.folders, parsed.folderPathPart);
                  if (f && f.id) { folderId = f.id; break; }
                }
                if (!folderId) {
                  return { error: `Could not resolve folderId for ${folderPath} (wanted path ${parsed.folderPathPart})` };
                }

                // Now search within that folder for the message with matching headerMessageId.
                const listRes = await messagesApi.list(folderId);
                let msgId = null;
                let messageListId = listRes && listRes.id;
                let chunk = listRes;
                let safety = 0;
                while (chunk && safety++ < 500) {
                  const msgs = chunk.messages || [];
                  for (const m of msgs) {
                    if (m && m.headerMessageId === messageIdHeader) {
                      msgId = m.id;
                      break;
                    }
                  }
                  if (msgId) break;
                  if (!chunk.id || !chunk.messages || chunk.messages.length === 0) break;
                  // Continue if there may be more.
                  try {
                    chunk = await messagesApi.continueList(chunk.id);
                  } catch {
                    break;
                  }
                }
                try { if (messageListId) await messagesApi.abortList(messageListId); } catch {}

                if (!msgId) {
                  return { error: `Could not find message in folder via messages.list/continueList for headerMessageId=${messageIdHeader}` };
                }

                // Build prefix early so we can try to pass it directly to beginReply(details)
                const plainPrefix = (typeof plainTextBody === "string") ? plainTextBody : "";
                let htmlPrefix = "";
                if (typeof htmlBody === "string" && htmlBody.trim()) {
                  htmlPrefix = htmlBody;
                } else if (plainPrefix) {
                  const esc = (s) => String(s)
                    .replace(/&/g, "&amp;")
                    .replace(/</g, "&lt;")
                    .replace(/>/g, "&gt;");
                  htmlPrefix = `<p>${esc(plainPrefix).replace(/\r\n/g, "\n").replace(/\n/g, "<br>")}</p>`;
                }

                const replyType = replyAll ? "replyToAll" : "replyToSender";

                // Prefer passing the body directly to beginReply(details). This tends to be applied by TB
                // during compose initialization (and is less likely to be overwritten later).
                const beginDetails = (plainPrefix && !htmlBody) ? { plainTextBody: plainPrefix } : (htmlPrefix ? { body: htmlPrefix } : undefined);
                const tab = beginDetails ? await composeApi.beginReply(msgId, replyType, beginDetails) : await composeApi.beginReply(msgId, replyType);
                const tabId = tab && tab.id;

                // Wait for quote insertion and then verify whether our prefix stuck.
                let detailsBefore = null;
                let details = null;
                for (let i = 0; i < 60; i++) {
                  details = await composeApi.getComposeDetails(tabId);
                  if (!detailsBefore) detailsBefore = details;

                  const b = (details && typeof details.body === "string") ? details.body : "";
                  const pb = (details && typeof details.plainTextBody === "string") ? details.plainTextBody : "";

                  const quoteReady = (b && /<blockquote[^>]*type=\"cite\"/i.test(b)) || (pb && pb.includes("wrote:"));
                  const prefixPresent = (plainPrefix && pb && pb.includes(plainPrefix.trim().slice(0, Math.min(20, plainPrefix.trim().length)))) ||
                    (htmlPrefix && b && b.includes(htmlPrefix.replace(/\s+/g, " ").slice(0, 10)));

                  if (quoteReady) {
                    // If quote is ready but prefix isn't, we will patch it via setComposeDetails below.
                    break;
                  }
                  await new Promise(r => Services.tm.dispatchToMainThread(() => r()));
                }

                const isPlain = !!(details && details.isPlainText);
                const existingBody = (details && typeof details.body === "string") ? details.body : "";
                const existingPlain = (details && typeof details.plainTextBody === "string") ? details.plainTextBody : "";

                // If the prefix didn't make it in via beginReply(details), patch it now.
                if (plainPrefix) {
                  if (existingPlain && !existingPlain.includes(plainPrefix.trim().slice(0, Math.min(20, plainPrefix.trim().length)))) {
                    const newPlain = includeQuotedOriginal ? (plainPrefix + "\n\n" + existingPlain) : plainPrefix;
                    await composeApi.setComposeDetails(tabId, { plainTextBody: newPlain });
                  }
                } else if (htmlPrefix) {
                  if (existingBody && !existingBody.includes(htmlPrefix.slice(0, 10))) {
                    const newBody = includeQuotedOriginal ? (htmlPrefix + existingBody) : htmlPrefix;
                    await composeApi.setComposeDetails(tabId, { body: newBody });
                  }
                }

                // Verify the prefix is present in the compose window before saving.
                let detailsAfterSet = await composeApi.getComposeDetails(tabId);
                for (let i = 0; i < 10; i++) {
                  const b = (detailsAfterSet && typeof detailsAfterSet.body === "string") ? detailsAfterSet.body : "";
                  const pb = (detailsAfterSet && typeof detailsAfterSet.plainTextBody === "string") ? detailsAfterSet.plainTextBody : "";
                  const want = (plainPrefix || htmlPrefix || "").trim();
                  const ok = want ? (pb.includes(want.split("\n")[0]) || b.includes(want.split("\n")[0]) || b.includes(htmlPrefix.slice(0, 10))) : true;
                  if (ok) break;
                  // Re-apply once more if not present.
                  if (plainPrefix) {
                    const newPlain = includeQuotedOriginal ? (plainPrefix + "\n\n" + pb) : plainPrefix;
                    await composeApi.setComposeDetails(tabId, { plainTextBody: newPlain });
                  } else if (htmlPrefix) {
                    const newBody = includeQuotedOriginal ? (htmlPrefix + b) : htmlPrefix;
                    await composeApi.setComposeDetails(tabId, { body: newBody });
                  }
                  await new Promise(r => Services.tm.dispatchToMainThread(() => r()));
                  detailsAfterSet = await composeApi.getComposeDetails(tabId);
                }

                // Give the composer a beat to flush changes before saving.
                await new Promise(r => Services.tm.dispatchToMainThread(() => r()));

                await composeApi.saveMessage(tabId, { mode: "draft" });

                const debug = {
                  isPlain,
                  beginDetailsUsed: !!beginDetails,
                  hadBodyBefore: !!(detailsBefore && detailsBefore.body),
                  hadPlainBefore: !!(detailsBefore && detailsBefore.plainTextBody),
                  afterHasBody: !!(detailsAfterSet && detailsAfterSet.body),
                  afterHasPlain: !!(detailsAfterSet && detailsAfterSet.plainTextBody),
                };

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
                  debug,
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

                // Efficiently keep only the newest N headers (avoid sorting entire folder, which can be huge).
                const top = [];
                for (const hdr of db.enumerateMessages()) {
                  const d = hdr && hdr.date ? hdr.date : 0;
                  if (top.length < max) {
                    top.push(hdr);
                    top.sort((a, b) => (b.date || 0) - (a.date || 0));
                  } else {
                    const worst = top[top.length - 1];
                    const wd = worst && worst.date ? worst.date : 0;
                    if (d > wd) {
                      top[top.length - 1] = hdr;
                      top.sort((a, b) => (b.date || 0) - (a.date || 0));
                    }
                  }
                }

                for (const hdr of top) {
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
                  return searchMessages(
                    args.query || "",
                    args.startDate,
                    args.endDate,
                    args.maxResults,
                    args.sortOrder
                  );
                case "listAccounts":
                  return listAccounts();
                case "listFolders":
                  return listFolders(args.accountKey);
                case "getRecentMessages":
                  return getRecentMessages(args.folderPath, args.limit, args.daysBack, args.unreadOnly);
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
                  return await replyToMessageDraftComposeApi(args.messageId, args.folderPath, args.replyAll, args.plainTextBody, args.htmlBody, args.includeQuotedOriginal, args.closeAfterSave);
                case "replyToMessageDraftNativeEditor":
                  return await replyToMessageDraftNativeEditor(args.messageId, args.folderPath, args.replyAll, args.plainTextBody, args.closeAfterSave);
                case "reviseDraftInPlaceNativeEditor":
                  return await reviseDraftInPlaceNativeEditor(args.messageId, args.folderPath, args.plainTextBody, args.preserveQuotedOriginal, args.closeAfterSave);
                case "debugContext":
                  return debugContext();
                case "debugMessagesList":
                  return await debugMessagesList(args.folderPath, args.limit);
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
          })();

          try {
            const result = await globalThis.__tbMcpStartPromise;
            return result;
          } finally {
            // Keep the promise so future calls return immediately.
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
