/* global browser */

let _mcpStartPromise = null;

async function init() {
  try {
    if (_mcpStartPromise) {
      return;
    }
    _mcpStartPromise = (async () => {
      const result = await browser.mcpServer.start();
      if (result.success) {
        console.log("MCP server started on port", result.port);
      } else {
        console.error("Failed to start MCP server:", result.error);
      }
      return result;
    })();

    await _mcpStartPromise;
  } catch (e) {
    console.error("Error starting MCP server:", e);
  }
}

browser.runtime.onInstalled.addListener(init);
browser.runtime.onStartup.addListener(init);

// In case events are missed (or multiple fire quickly), attempt once at load.
init();
