# Draft automation notes (Thunderbird + Outlook/OWA)

This document records a real-world investigation into **creating reply drafts in Thunderbird** that are:

- **Saved into IMAP Drafts**
- **Recognized by Outlook Web (OWA) as true Drafts** (sendable, proper Draft behavior)
- **Preserve Thunderbird-native reply quoting** (moz-cite-prefix + blockquote)
- **Do not change email state** unless explicitly instructed (read/unread/move/delete)

It also documents the tooling and the final reliable workflow that succeeded.

## Problem statement
We needed an automated way to generate a reply draft to a specific message (Candlewyck Park “Move Out Summary”), inserting a custom reply body at the top while keeping the original message quoted below, and ensuring the draft is **sendable in OWA**.

### Key constraints
- Do **not** change message state (mark read, move, delete) unless explicitly instructed.
- Draft must be **sendable in OWA**. A raw IMAP APPEND draft is often *not* sufficient.
- Reply draft must keep Thunderbird-style quoting:
  - `moz-cite-prefix`
  - `<blockquote type="cite" cite="mid:...">...` with the original message.
- Must avoid duplicate drafts on retries (idempotency).

## What we tried (and why it failed)

### 1) Server-side draft creation by appending RFC822 (MailServices.copy / copyFileMessage)
We implemented backend draft creation by writing an RFC822 message to a temp file and copying it into the Drafts folder.

**Pros:**
- Works without UI automation.
- Can be made idempotent via deterministic Message-ID.

**Cons / why insufficient here:**
- Drafts created via IMAP append can fail to get the correct server-side properties that OWA expects for a “real” draft.
- Some attempts produced empty bodies or inconsistent content.
- Even when content was correct, OWA semantics were unreliable.

### 2) Compose-window “Save as Draft” via internal listeners
We attempted to open compose windows and invoke save-draft flows, waiting on callbacks.

**Problems seen:**
- Compose save callbacks can hang or never fire.
- Tool calls sometimes returned success but no draft appeared.
- Extension logging was intermittent.

### 3) Opening a “reply” but accidentally using `nsIMsgCompType.New`
Early work opened a compose window that *looked* like a reply (subject/headers set), but because the compose type was New, Thunderbird did not insert the quoted original.

**Symptom:**
- Draft contains only the custom body, no quote.

## Final working approach

### Summary
The reliable solution was:

1. Open a **true native Reply/ReplyAll compose window** using `nsIMsgComposeService.OpenComposeWindowWithParams` with:
   - `type = Reply` or `ReplyAll`
   - `originalMsgURI = msgHdr.folder.getUriForMsg(msgHdr)`
2. Use OS-level UI automation on Linux to:
   - Focus the compose body
   - Insert the custom reply text at the top
   - Close the window (Ctrl+W)
   - Choose **Save** on the “Save this message?” prompt

This produced a draft that:
- Appears in the IMAP Drafts folder
- Is recognized by OWA as a sendable Draft
- Includes the quoted original message beneath the reply

### Why UI automation was needed
Thunderbird’s compose editor is not reliably scriptable from the WebExtension background context for inserting body text (clipboard/paste/editor APIs were inconsistent across attempts). OS-level tooling solved this.

### Linux tooling
- `xdotool` (window focus, keystrokes, mouse click)
- `xclip` (set clipboard text)

Install:
```bash
sudo apt-get install -y xdotool xclip
```

## Formatting gotcha: newlines/paragraphs
A subtle issue: pasting plain text could result in the reply rendering as a single paragraph (no blank lines) depending on how Thunderbird interprets pasted content.

**Fix:** use “paste as plain text” (often Ctrl+Shift+V) or type with explicit Return keystrokes so the draft contains `<br><br>` (or separate `<p>` blocks) in the HTML body.

We validated formatting by comparing to a user-created sample draft (e.g., “Re: CORNELLTECH Alert: …”).

## Tooling changes in `thunderbird-mcp`

### Added: `openNativeReplyCompose`
A new MCP tool was added to open a true Reply/ReplyAll compose window that includes Thunderbird-native quoting.

High-level behavior:
- Find message header by `messageId` in `folderPath`.
- Create `nsIMsgComposeParams` with:
  - `type = Reply` / `ReplyAll`
  - `format = HTML`
  - `identity = getIdentityForFolder(folder)`
  - `originalMsgURI = msgHdr.folder.getUriForMsg(msgHdr)`
- Call `OpenComposeWindowWithParams`.

### Notable lessons
- If you use `nsIMsgCompType.New`, Thunderbird will not generate the reply quote.
- If the compose is not “dirty”, closing may not trigger Save-as-draft.

## Verification techniques
- Poll Drafts via `listLatestMessages(folderPath, limit)`.
- Inspect raw MIME source via `getRawMessage`:
  - Ensure reply text is above the quote.
  - Ensure the quote exists (`moz-cite-prefix`, `<blockquote type="cite" ...>`).
  - Check for expected headers (`X-Mozilla-Draft-Info`, `X-Identity-Key`, etc.).

## Idempotency / avoiding duplicates
When performing operations that may be retried, add an `idempotencyKey` and use deterministic Message-IDs / a pending set to prevent duplicate draft creation.

## Operational notes
- Restarting Thunderbird is sometimes required after installing a new XPI.
- Port conflicts (`NS_ERROR_SOCKET_ADDRESS_IN_USE`) can occur on `8765` if the add-on didn’t fully shut down.

## Outcome
The final workflow produced a correct, sendable OWA draft with proper quoting and formatting. The user successfully sent the email.
