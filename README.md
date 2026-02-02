# Thunderbird MCP

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Thunderbird](https://img.shields.io/badge/Thunderbird-102%2B-0a84ff.svg)](https://www.thunderbird.net/)
[![MCP](https://img.shields.io/badge/MCP-compatible-green.svg)](https://modelcontextprotocol.io/)

> Inspired by [bb1/thunderbird-mcp](https://github.com/bb1/thunderbird-mcp). Rewritten from scratch with a bundled HTTP server, proper MIME decoding, and UTF-8 handling throughout.

An MCP server that lets you read email, search contacts, and draft replies in Thunderbird.

## How it works

```
MCP Client <--stdio--> mcp-bridge.cjs <--HTTP--> Thunderbird Extension
```

The Thunderbird extension runs a local HTTP server on port 8765. The Node.js bridge translates between MCP's stdio protocol and HTTP so MCP clients can talk to it.

## Setup

**1. Install the extension**

```bash
./scripts/build.sh
./scripts/install.sh
```

Restart Thunderbird.

**2. Configure your MCP client**

Example for `~/.claude.json`:

```json
{
  "mcpServers": {
    "thunderbird-mail": {
      "command": "node",
      "args": ["/path/to/thunderbird-mcp/mcp-bridge.cjs"]
    }
  }
}
```

## What you can do

| Tool | What it does |
|------|--------------|
| `searchMessages` | Find emails by subject, sender, or recipient |
| `getMessage` | Read full email content |
| `searchContacts` | Look up contacts |
| `listCalendars` | List your calendars |
| `sendMail` | Open a compose window with pre-filled content |
| `replyToMessage` | Open a reply with proper threading |

Compose tools open a window for you to review before sending. Nothing gets sent automatically.

## Use native Thunderbird WebExtension APIs (preferred)

When implementing new capabilities, **prefer Thunderbird's native WebExtension APIs** whenever possible. They are documented here:

- <https://webextension-api.thunderbird.net/en/mv3/index.html>

This documentation is the canonical reference for what extensions can do in a stable, supported way.

Examples of high-leverage native APIs:
- `browser.messages.*` – list/search/read messages, attachments, move/copy/delete (with the right permissions)
- `browser.compose.*` – beginReply/beginNew, setComposeDetails, saveMessage (draft), sendMessage
- `browser.accounts.*` – list accounts/identities/folders and resolve folder ids

We still use XPCOM/Thunderbird-internal APIs in the experiment context for a few things, but the long-term direction is to minimize that and build on the official WebExtension surface.

## Security

The extension only listens on localhost, but any local process can access it while Thunderbird is running. Keep this in mind on shared machines.

## Troubleshooting

**Extension not loading?**
Check Tools → Add-ons and Themes. For errors: Tools → Developer Tools → Error Console.

**Connection refused?**
Make sure Thunderbird is running and the extension is enabled.

**Can't find recent emails?**
IMAP folders can be stale. Click on the folder in Thunderbird to sync, or right-click → Properties → Repair Folder.

## Development

```bash
# Build the extension
./scripts/build.sh

# Test the HTTP API directly
curl -X POST http://localhost:8765 \
  -H "Content-Type: application/json" \
  -d '{"jsonrpc":"2.0","id":1,"method":"tools/list"}'

# Test the bridge
echo '{"jsonrpc":"2.0","id":1,"method":"tools/list"}' | node mcp-bridge.cjs
```

After changing extension code, you'll need to remove it from Thunderbird, restart, reinstall, and restart again. Thunderbird caches aggressively.

## Known issues

- Replies don't include the quoted original message (Thunderbird limitation workaround)
- IMAP folder databases can be stale until you click on them
- Email bodies with weird control characters get sanitized to avoid breaking JSON

## Project structure

```
thunderbird-mcp/
├── mcp-bridge.cjs              # stdio-to-HTTP bridge
├── extension/
│   ├── manifest.json
│   ├── background.js           # Extension entry point
│   ├── httpd.sys.mjs           # Mozilla's HTTP server lib
│   └── mcp_server/
│       ├── api.js              # The actual MCP implementation
│       └── schema.json
└── scripts/
    ├── build.sh
    └── install.sh
```

## License

MIT. The bundled `httpd.sys.mjs` is from Mozilla and licensed under MPL-2.0.
