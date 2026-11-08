# Thunderbird MCP draft-only fork

This branch integrates [Thunderbird MCP v0.10.0](https://github.com/TKasperczyk/thunderbird-mcp/tree/v0.10.0) while keeping this fork's extension ID and local draft workflow. The Node stdio bridge and Thunderbird experiment use the same draft-only rules. Use the bridge and XPI from the same checkout.

## Mail workflow

- `getLatestUnread` and `getLatestUnreadBatch` read the newest unread messages without changing their read state. If no folder is specified, they use the Inbox of an accessible account. `setMessageRead` explicitly changes the flag through Thunderbird's folder API.
- `getMessage` defaults to plain text, uses upstream MIME decoding and fallback, and provides the legacy `bodyText`, `bodyHtml`, and `bodyType` fields when they fit the output budget. Request `bodyFormat: "html"` to get the stored HTML where available. Raw reads, including `getRawMessage`, use upstream account and encrypted-content checks.
- `saveDraft` and `saveDraftFromEmail` save drafts. `replyToMessage` and the legacy `replyToMessageDraft`, `replyToMessageDraftComposeApi`, and `replyToMessageDraftNativeEditor` names use the upstream threaded draft save. These reply aliases keep the quote and close the saved compose window. Legacy window and close-prompt controls return an error.
- `sendMail`, `sendMailNow`, `forwardMessage`, scheduled delivery, and unreviewed replies are blocked. Calendar writes and filter creation, edits, or manual runs that could deliver mail are also blocked. The upstream Options page may show its original controls. Those preferences cannot enable these operations in the fork. Send reviewed drafts manually in Thunderbird.
- A customized upstream tool allowlist may require enabling the compatibility aliases explicitly. An alias also requires its underlying read, update, or draft tool to be enabled.

The old native in-place editor, unsaved compose-window helper, and debug tools were not carried forward. To replace a draft through the current API, use `saveDraft` with `replaceMessageId` and `replaceFolderPath` and supply any fields to retain, including recipients, threading headers, and attachments. The old automation notes and shell scripts document the earlier transport and are historical.

## Build and offline tests

```sh
npm test
bash scripts/build.sh
```

The build writes `dist/thunderbird-mcp.xpi`. The bridge and extension use upstream's authenticated loopback transport. For transport details and the inherited read and organization tools, see the [upstream v0.10.0 documentation](https://github.com/TKasperczyk/thunderbird-mcp/blob/v0.10.0/README.md). The policy above applies to this fork.
