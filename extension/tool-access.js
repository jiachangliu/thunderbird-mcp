"use strict";

// Searched isToolEnabled/getToolAccessConfig/onToolChange: they independently
// interpreted the deny-list. Share effective access between the Experiment and
// Options, while callers retain each tool's own configured preference.
// Classic script: loadSubScript binds `this` to its supplied scope; the
// Options page binds it to the window. No preference or DOM access lives here.
this.ToolAccess = (() => {
  // A fork policy, independent of the user's upstream preferences. Both the
  // authenticated experiment endpoint and stdio relay enforce it.
  const forkBlocked = new Set([
    "sendMail", "sendMailNow", "forwardMessage", "applyFilters", "createFilter", "updateFilter",
    "createEvent", "updateEvent", "deleteEvent", "createTask", "updateTask",
  ]);

  function forkPolicyError(name, args) {
    if (forkBlocked.has(name)) return `${name} is disabled for safety (draft-only workflow).`;
    if (name === "saveDraft" && args?.sendAt != null) {
      return "Scheduled sending is disabled for safety (draft-only workflow). Omit sendAt.";
    }
    if (name === "replyToMessage" && [true, "true"].includes(args?.skipReview)) {
      return "Direct replies are disabled for safety (draft-only workflow).";
    }
    return null;
  }

  const dependencies = Object.freeze({
    dryRunCompose: Object.freeze(["sendMail"]),
    getLatestUnread: Object.freeze(["getMessage"]),
    getLatestUnreadBatch: Object.freeze(["getMessage"]),
    getRawMessage: Object.freeze(["getMessage"]),
    setMessageRead: Object.freeze(["updateMessage"]),
    listLatestMessages: Object.freeze(["getRecentMessages"]),
    moveMessages: Object.freeze(["updateMessage"]),
    saveDraftFromEmail: Object.freeze(["saveDraft"]),
    replyToMessageDraft: Object.freeze(["replyToMessage", "saveDraft"]),
    replyToMessageDraftComposeApi: Object.freeze(["replyToMessage", "saveDraft"]),
    replyToMessageDraftNativeEditor: Object.freeze(["replyToMessage", "saveDraft"]),
  });
  // Permanent inventory, not a default for unregistered names. Every later
  // tool needs one introduction entry, even if another tool already blocks it.
  const baselineRelease = "0.9.3";
  const legacyTools = Object.freeze([
    "listAccounts", "listFolders", "getAccountAccess", "searchMessages", "getMessage",
    "getMessages", "getRecentMessages", "displayMessage", "sendMail", "replyToMessage",
    "forwardMessage", "updateMessage", "deleteMessages", "saveDraft", "createFolder",
    "renameFolder", "moveFolder", "deleteFolder", "emptyTrash", "emptyJunk", "searchContacts",
    "getContact", "createContact", "updateContact", "deleteContact", "listCalendars",
    "listEvents", "listCategories", "createEvent", "createTask", "listTasks", "updateTask",
    "updateEvent", "deleteEvent", "listFilters", "createFilter", "updateFilter",
    "reorderFilters", "applyFilters", "deleteFilter",
  ]);
  const introducedIn = Object.freeze({
    dryRunCompose: "0.10.0",
    refreshFolder: "0.10.0",
    refreshCalendar: "0.10.0",
    getMessageHeaders: "0.10.0",
    batchGetMessageHeaders: "0.10.0",
    searchAttachments: "0.10.0",
    createVirtualFolder: "0.10.0",
    listVirtualFolders: "0.10.0",
    deleteVirtualFolder: "0.10.0",
    getLatestUnread: "0.10.0",
    getLatestUnreadBatch: "0.10.0",
    getRawMessage: "0.10.0",
    setMessageRead: "0.10.0",
    listLatestMessages: "0.10.0",
    moveMessages: "0.10.0",
    saveDraftFromEmail: "0.10.0",
    replyToMessageDraft: "0.10.0",
    replyToMessageDraftComposeApi: "0.10.0",
    replyToMessageDraftNativeEditor: "0.10.0",
    sendMailNow: "0.10.0",
  });
  const stateVersion = 1;
  // Searched version comparisons and preference parsers: none handles our
  // release catalog or structured tracking record. Keep these pure for VM tests.
  function releaseParts(release) {
    if (typeof release !== "string" || !/^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$/.test(release)) {
      throw new Error("Invalid tool catalog release");
    }
    const parts = release.split(".").map(Number);
    if (!parts.every(Number.isSafeInteger)) throw new Error("Invalid tool catalog release");
    return parts;
  }

  function compareReleases(left, right) {
    const a = releaseParts(left);
    const b = releaseParts(right);
    for (let i = 0; i < a.length; i++) {
      if (a[i] !== b[i]) return a[i] < b[i] ? -1 : 1;
    }
    return 0;
  }

  function parseState(raw) {
    const state = JSON.parse(raw);
    if (!state || Array.isArray(state) || typeof state !== "object" ||
        Object.keys(state).sort().join(",") !== "consideredRelease,consideredTools,customized,version" ||
        state.version !== stateVersion || typeof state.customized !== "boolean" ||
        !Array.isArray(state.consideredTools) ||
        !state.consideredTools.every(name => typeof name === "string" && name.length > 0 && name !== "__all__") ||
        new Set(state.consideredTools).size !== state.consideredTools.length ||
        !legacyTools.every(name => state.consideredTools.includes(name)) ||
        compareReleases(state.consideredRelease, baselineRelease) < 0 ||
        state.consideredTools.some(name => Object.hasOwn(introducedIn, name) &&
          compareReleases(introducedIn[name], state.consideredRelease) > 0)) {
      throw new Error("Invalid tool access tracking record");
    }
    return state;
  }

  function considerTools(disabled, previous, names) {
    const considered = new Set(previous?.consideredTools || legacyTools);
    const customized = !!previous?.customized || disabled.length > 0;
    let consideredRelease = previous?.consideredRelease || baselineRelease;
    const denied = new Set(disabled);
    for (const name of names) {
      if (!legacyTools.includes(name)) {
        if (!Object.hasOwn(introducedIn, name) || compareReleases(introducedIn[name], baselineRelease) <= 0) {
          throw new Error(`Missing introduction release for tool: ${name}`);
        }
        const release = introducedIn[name];
        if (compareReleases(release, consideredRelease) > 0) consideredRelease = release;
        if (customized && !considered.has(name)) denied.add(name);
      }
      considered.add(name);
    }
    return {
      disabled: [...denied],
      state: { version: stateVersion, customized, consideredRelease, consideredTools: [...considered] },
    };
  }

  function isEnabled(name, disabled, undisableable, visiting = new Set()) {
    if (undisableable.has(name)) return true;
    if (disabled.includes("__all__") || disabled.includes(name) || visiting.has(name)) return false;
    const parents = Object.hasOwn(dependencies, name) ? dependencies[name] : [];
    const next = new Set(visiting).add(name);
    return parents.every(parent => isEnabled(parent, disabled, undisableable, next));
  }
  return Object.freeze({ forkPolicyError, dependencies, isEnabled, legacyTools, introducedIn, compareReleases, parseState, considerTools });
})();
