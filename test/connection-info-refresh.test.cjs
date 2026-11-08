"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function loadConnectionInfoRefreshHelpers(overrides = {}) {
  const apiPath = path.resolve(__dirname, "../extension/mcp_server/api.js");
  const source = fs.readFileSync(apiPath, "utf8");
  const snippets = [];
  for (const name of ["CONNECTION INFO REFRESH HELPERS", "CONNECTION INFO READER"]) {
    const start = source.indexOf(`// BEGIN ${name}`);
    const end = source.indexOf(`// END ${name}`, start);
    assert.ok(start >= 0 && end > start, `Missing production marker: ${name}`);
    snippets.push(source.slice(start, end));
  }
  const sandbox = { TextDecoder, ...overrides };
  vm.createContext(sandbox);
  vm.runInContext(
    `${snippets.join("\n")}
this.ensureFreshConnectionInfo = ensureFreshConnectionInfo;`,
    sandbox
  );
  return {
    ensureFreshConnectionInfo: sandbox.ensureFreshConnectionInfo,
    readConnectionInfo: sandbox.readConnectionInfo,
  };
}

const { ensureFreshConnectionInfo } = loadConnectionInfoRefreshHelpers();

function makeOptions(overrides) {
  return {
    port: 8765,
    token: "token",
    expectedPid: 1234,
    onCheckError() {},
    ...overrides,
  };
}

describe("connection info refresh decision", () => {
  for (const profileRoots of [undefined, ["/profiles/Łukasz", "/profiles/郵件"]]) {
    it(`reads UTF-8 connection data without unnecessary refresh: ${JSON.stringify(profileRoots)}`, () => {
      const data = { port: 21000, token: "a".repeat(64), pid: 1234, profileRoots };
      const bytes = Buffer.from(JSON.stringify(data), "utf8").toString("latin1");
      let closed = false;
      const file = { path: "/fixture/connection.json", append() {}, clone() { return this; }, exists: () => true };
      const helpers = loadConnectionInfoRefreshHelpers({
        Services: { dirsvc: { get: () => file } },
        Ci: {},
        Cc: {
          "@mozilla.org/network/file-input-stream;1": { createInstance: () => ({ init() {} }) },
          "@mozilla.org/scriptableinputstream;1": { createInstance: () => ({
            init() {}, available: () => bytes.length,
            read(length) { assert.equal(length, bytes.length); return bytes; },
            close() { closed = true; },
          }) },
        },
      });
      const current = helpers.readConnectionInfo();
      assert.deepEqual(JSON.parse(JSON.stringify(current.data)), JSON.parse(JSON.stringify(data)));
      assert.equal(closed, true);
      assert.equal(helpers.ensureFreshConnectionInfo({
        port: data.port, token: data.token, expectedPid: data.pid, expectedProfileRoots: profileRoots,
        readConnectionInfo: helpers.readConnectionInfo,
        writeConnectionInfo() { assert.fail("Matching UTF-8 metadata must not cause a rewrite"); },
      }), file.path);
    });
  }

  it("does not rewrite when connection.json already matches the running server", () => {
    const result = ensureFreshConnectionInfo(makeOptions({
      readConnectionInfo: () => ({
        path: "/tmp/thunderbird-mcp/connection.json",
        data: { port: 8765, token: "token", pid: 1234 },
      }),
      writeConnectionInfo() {
        throw new Error("writeConnectionInfo should not be called");
      },
    }));

    assert.equal(result, "/tmp/thunderbird-mcp/connection.json");
  });

  it("rewrites when connection.json is missing", () => {
    const writes = [];
    const result = ensureFreshConnectionInfo(makeOptions({
      readConnectionInfo: () => ({
        path: "/tmp/thunderbird-mcp/connection.json",
        data: null,
      }),
      writeConnectionInfo(port, token) {
        writes.push({ port, token });
        return "/tmp/thunderbird-mcp/connection.json";
      },
    }));

    assert.equal(result, "/tmp/thunderbird-mcp/connection.json");
    assert.deepEqual(writes, [{ port: 8765, token: "token" }]);
  });

  for (const staleCase of [
    {
      name: "port changed",
      data: { port: 8766, token: "token", pid: 1234 },
    },
    {
      name: "token changed",
      data: { port: 8765, token: "old-token", pid: 1234 },
    },
    {
      name: "pid changed",
      data: { port: 8765, token: "token", pid: 4321 },
    },
  ]) {
    it(`rewrites when connection.json is stale: ${staleCase.name}`, () => {
      let writeCount = 0;
      const result = ensureFreshConnectionInfo(makeOptions({
        readConnectionInfo: () => ({
          path: "/tmp/thunderbird-mcp/connection.json",
          data: staleCase.data,
        }),
        writeConnectionInfo(port, token) {
          writeCount++;
          assert.equal(port, 8765);
          assert.equal(token, "token");
          return "/tmp/thunderbird-mcp/connection.json";
        },
      }));

      assert.equal(result, "/tmp/thunderbird-mcp/connection.json");
      assert.equal(writeCount, 1);
    });
  }

  it("rewrites when reading or parsing connection.json fails", () => {
    const errors = [];
    let writeCount = 0;
    const parseError = new SyntaxError("Unexpected token");

    const result = ensureFreshConnectionInfo(makeOptions({
      readConnectionInfo() {
        throw parseError;
      },
      writeConnectionInfo() {
        writeCount++;
        return "/tmp/thunderbird-mcp/connection.json";
      },
      onCheckError(error) {
        errors.push(error);
      },
    }));

    assert.equal(result, "/tmp/thunderbird-mcp/connection.json");
    assert.deepEqual(errors, [parseError]);
    assert.equal(writeCount, 1);
  });

  const PROFILE_ROOTS = ["/profiles/current", "/storage/mail/current", "/config/thunderbird"];

  it("does not rewrite when the profile metadata matches", () => {
    const result = ensureFreshConnectionInfo(makeOptions({
      expectedProfileRoots: PROFILE_ROOTS,
      readConnectionInfo: () => ({
        path: "/tmp/thunderbird-mcp/connection.json",
        data: { port: 8765, token: "token", pid: 1234, profileRoots: [...PROFILE_ROOTS] },
      }),
      writeConnectionInfo() { assert.fail("Matching metadata must not cause a write"); },
    }));

    assert.equal(result, "/tmp/thunderbird-mcp/connection.json");
  });

  for (const [name, profileRoots] of [
    ["missing", undefined],
    ["stale", ["/profiles/old"]],
    ["incomplete", PROFILE_ROOTS.slice(0, 1)],
    ["malformed", PROFILE_ROOTS[0]],
    ["null", null],
  ]) {
    it(`rewrites when profile metadata is ${name}`, () => {
      const writes = [];
      const result = ensureFreshConnectionInfo(makeOptions({
        expectedProfileRoots: PROFILE_ROOTS,
        readConnectionInfo: () => ({
          path: "/tmp/thunderbird-mcp/connection.json",
          data: { port: 8765, token: "token", pid: 1234, ...(profileRoots === undefined ? {} : { profileRoots }) },
        }),
        writeConnectionInfo(port, token) {
          writes.push({ port, token });
          return "/tmp/thunderbird-mcp/connection.json";
        },
      }));

      assert.equal(result, "/tmp/thunderbird-mcp/connection.json");
      assert.deepEqual(writes, [{ port: 8765, token: "token" }]);
    });
  }

  it("propagates write failures", () => {
    const writeError = new Error("disk full");

    assert.throws(
      () => ensureFreshConnectionInfo(makeOptions({
        readConnectionInfo: () => ({
          path: "/tmp/thunderbird-mcp/connection.json",
          data: null,
        }),
        writeConnectionInfo() {
          throw writeError;
        },
      })),
      (error) => error === writeError
    );
  });
});
