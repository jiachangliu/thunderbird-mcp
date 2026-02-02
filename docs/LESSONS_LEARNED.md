# Lessons Learned (thunderbird-mcp)

This file is intentionally short and practical.

## 1) For Draft/reply-draft revision: don’t guess the saved Message-ID

Thunderbird may save a draft by **creating a new message** (new `Message-ID`) and leaving the original draft untouched.

**Correct approach:**
- When saving from a compose tab, use the return value of `compose.saveMessage(tabId, { mode: "draft" })`.
- It returns `messages[]`, and each item includes `headerMessageId`.
- That `headerMessageId` is the *actual* draft item that was saved and synced.

This is strictly more reliable than:
- scanning Drafts by date/subject,
- relying on `msgDatabase.enumerateMessages()` ordering,
- diffing lists from `messages.list()` (server ordering can be surprising),
- or parsing internal compose window state.

## 2) Keep IMAP operations bounded to avoid HTTP timeouts

Operations like:
- enumerating/sorting large folders,
- fetching raw MIME for many candidates,
- moving/deleting large batches,

can block the HTTP handler long enough to trigger client timeouts (`curl: (28) ...`). Prefer:
- page-based APIs (`messages.list/continueList/abortList`) where possible,
- small candidate sets and early exits,
- one-by-one cleanup when a provider is slow.

## 3) Restart/port conflicts can happen; always verify listener

If you see:
- `NS_ERROR_SOCKET_ADDRESS_IN_USE` or
- connection refused to `127.0.0.1:8765`,

then the MCP server is not listening. Verify with `ss -ltnp | grep 8765` and restart Thunderbird cleanly.

## 4) Cleanup strategy for test drafts

Bulk delete/move of Drafts can be slow. Prefer:
- selecting known test artifacts by subject/prefix/token,
- deleting/purging one-by-one with retries.

See: `scripts/cleanup-test-drafts.sh`.
