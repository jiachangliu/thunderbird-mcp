"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const CONNECTION_READ_LIMIT = 4096;

function loadWriteConnectionInfo(sandbox) {
  const apiPath = path.resolve(__dirname, "../extension/mcp_server/api.js");
  const source = fs.readFileSync(apiPath, "utf8");
  const snippets = [];
  for (const name of ["SENSITIVE ATTACHMENT PATH HELPERS", "THUNDERBIRD PROFILE ROOTS", "CONNECTION INFO WRITER", "CONNECTION INFO REFRESH HELPERS"]) {
    const start = source.indexOf(`// BEGIN ${name}`);
    const end = source.indexOf(`// END ${name}`, start);
    assert.ok(start >= 0 && end > start, `Missing production marker: ${name}`);
    snippets.push(source.slice(start, end));
  }

  vm.createContext(sandbox);
  vm.runInContext(
    `${snippets.join("\n")}
this.writeConnectionInfo = writeConnectionInfo;`,
    sandbox
  );
  return sandbox.writeConnectionInfo;
}

/**
 * Minimal nsIFile stand-in for <TmpD>/thunderbird-mcp. `permissions` mimics
 * platform behaviour: Windows can report synthesised group/other bits that
 * chmod cannot clear; on POSIX the assignment takes effect when `chmodWorks`.
 */
function makeTmpDir({
  exists = true,
  permissions = 0o777,
  chmodWorks = true,
  symlink = false,
  connectionFileExists = false,
}) {
  let mode = permissions;
  const dir = {
    segments: [],
    chmodCalls: [],
    createCalls: [],
    removeCalls: [],
    events: [],
    permissionReads: 0,
    append(name) { this.segments.push(name); },
    exists() { return exists; },
    isSymlink() { return symlink; },
    create(type, permissions) {
      this.createCalls.push({ type, permissions });
      exists = true;
    },
    get permissions() {
      this.permissionReads++;
      return mode;
    },
    set permissions(value) {
      this.chmodCalls.push(value);
      if (chmodWorks) mode = value;
    },
    clone() {
      const file = {
        segments: [...this.segments],
        append(name) { this.segments.push(name); },
        exists() { return connectionFileExists; },
        remove(recursive) {
          dir.events.push("remove");
          dir.removeCalls.push(recursive);
          connectionFileExists = false;
        },
        get path() { return this.segments.join("/"); },
      };
      return file;
    },
  };
  return dir;
}

function makeSandbox({ os, tmpDir, profileDirs = {} }) {
  const written = [];
  const opened = [];
  const profilePath = os === "WINNT" ? path.win32 : path.posix;
  function makeProfileFile(filePath, dir) {
    let readableChecked = false;
    return {
      path: filePath,
      clone() { return makeProfileFile(this.path, dir); },
      normalize() {
        if (dir.normalizeError) throw dir.normalizeError;
        if (this.path === dir.path) this.path = dir.normalizedPath || this.path;
      },
      initWithPath(value) {
        if (typeof value !== "string" || !profilePath.isAbsolute(value)) throw new Error("Invalid profile target");
        this.path = value;
        readableChecked = false;
      },
      append(name) { this.path = profilePath.join(this.path, name); readableChecked = false; },
      get leafName() { return profilePath.basename(this.path); },
      get parent() {
        const parentPath = profilePath.dirname(this.path);
        return parentPath === this.path ? null : makeProfileFile(parentPath, dir);
      },
      isReadable() {
        readableChecked = true;
        return dir.readable !== false && !dir.unreadablePaths?.includes(this.path);
      },
      get target() {
        assert.equal(readableChecked, true, "Windows target requires a resolution check");
        if (Object.hasOwn(dir.targets || {}, this.path)) return dir.targets[this.path];
        return this.path === (dir.normalizedPath || dir.path) ? (dir.target || this.path) : this.path;
      },
    };
  }
  return {
    written,
    opened,
    sandbox: {
      TextEncoder,
      Services: {
        dirsvc: {
          get(key) {
            if (key === "TmpD") return tmpDir;
            const dir = profileDirs[key];
            if (!dir) throw new Error("Directory unavailable");
            return makeProfileFile(dir.path, dir);
          },
        },
        appinfo: { OS: os, processID: 4242 },
      },
      Ci: { nsIFile: { DIRECTORY_TYPE: 1 } },
      Cc: {
        "@mozilla.org/network/file-output-stream;1": {
          createInstance: () => ({
            init(file, flags, permissions, behavior) {
              tmpDir.events.push("open");
              opened.push({
                path: file.path,
                flags,
                permissions,
                behavior,
                fileExisted: file.exists(),
              });
            },
          }),
        },
        "@mozilla.org/intl/converter-output-stream;1": {
          createInstance: () => ({
            init() {},
            writeString(data) {
              tmpDir.events.push("write");
              written.push(data);
            },
            close() { tmpDir.events.push("close"); },
          }),
        },
      },
    },
  };
}

describe("writeConnectionInfo directory hardening", () => {
  for (const permissions of [0o666, 0o777]) {
    it(`skips the POSIX permission check on Windows with synthesised 0o${permissions.toString(8)}`, () => {
      const tmpDir = makeTmpDir({ permissions, chmodWorks: false });
      const { sandbox, written } = makeSandbox({ os: "WINNT", tmpDir });
      const writeConnectionInfo = loadWriteConnectionInfo(sandbox);

      const result = writeConnectionInfo(8765, "token");

      assert.equal(result, "thunderbird-mcp/connection.json");
      assert.equal(tmpDir.permissionReads, 0);
      assert.deepEqual(tmpDir.chmodCalls, []);
      assert.deepEqual(JSON.parse(written[0]), { port: 8765, token: "token", pid: 4242 });
    });
  }

  for (const os of ["Linux", "Darwin"]) {
    for (const permissions of [0o770, 0o707, 0o777]) {
      it(`refuses 0o${permissions.toString(8)} on ${os} when chmod does not clear group/other bits`, () => {
        const tmpDir = makeTmpDir({ permissions, chmodWorks: false });
        const { sandbox, opened, written } = makeSandbox({ os, tmpDir });
        const writeConnectionInfo = loadWriteConnectionInfo(sandbox);

        assert.throws(
          () => writeConnectionInfo(8765, "token"),
          /group\/world permissions/
        );
        assert.deepEqual(tmpDir.chmodCalls, [0o700]);
        assert.deepEqual(opened, []);
        assert.deepEqual(written, []);
      });
    }

    it(`repairs a permissive directory on ${os} when chmod succeeds`, () => {
      const tmpDir = makeTmpDir({ permissions: 0o755, chmodWorks: true });
      const { sandbox, written } = makeSandbox({ os, tmpDir });
      const writeConnectionInfo = loadWriteConnectionInfo(sandbox);

      const result = writeConnectionInfo(8770, "token");

      assert.equal(result, "thunderbird-mcp/connection.json");
      assert.deepEqual(tmpDir.chmodCalls, [0o700]);
      assert.equal(tmpDir.permissions, 0o700);
      assert.equal(written.length, 1);
    });

    for (const permissions of [0o700, 0]) {
      it(`preserves existing 0o${permissions.toString(8)} handling on ${os} without chmod`, () => {
        const tmpDir = makeTmpDir({ permissions });
        const { sandbox, written } = makeSandbox({ os, tmpDir });
        const writeConnectionInfo = loadWriteConnectionInfo(sandbox);

        assert.equal(writeConnectionInfo(8765, "token"), "thunderbird-mcp/connection.json");
        assert.equal(tmpDir.permissionReads, 1);
        assert.deepEqual(tmpDir.chmodCalls, []);
        assert.equal(written.length, 1);
      });
    }
  }

  for (const os of ["WINNT", "Linux", "Darwin"]) {
    it(`creates a missing directory with 0700 on ${os}`, () => {
      const tmpDir = makeTmpDir({ exists: false });
      const { sandbox, written } = makeSandbox({ os, tmpDir });
      const writeConnectionInfo = loadWriteConnectionInfo(sandbox);

      assert.equal(writeConnectionInfo(8765, "token"), "thunderbird-mcp/connection.json");
      assert.deepEqual(tmpDir.createCalls, [{ type: 1, permissions: 0o700 }]);
      assert.equal(tmpDir.exists(), true);
      assert.equal(tmpDir.permissionReads, 0);
      assert.deepEqual(tmpDir.chmodCalls, []);
      assert.equal(written.length, 1);
    });

    it(`rejects a symlink directory before writing on ${os}`, () => {
      const tmpDir = makeTmpDir({ symlink: true });
      const { sandbox, opened, written } = makeSandbox({ os, tmpDir });
      const writeConnectionInfo = loadWriteConnectionInfo(sandbox);

      assert.throws(() => writeConnectionInfo(8765, "token"), /directory is a symlink/);
      assert.deepEqual(tmpDir.createCalls, []);
      assert.equal(tmpDir.permissionReads, 0);
      assert.deepEqual(tmpDir.chmodCalls, []);
      assert.deepEqual(opened, []);
      assert.deepEqual(written, []);
    });

    for (const connectionFileExists of [false, true]) {
      it(`writes ${connectionFileExists ? "a replacement" : "a new"} file with exclusive creation and 0600 on ${os}`, () => {
        const tmpDir = makeTmpDir({ permissions: 0o700, connectionFileExists });
        const { sandbox, opened, written } = makeSandbox({ os, tmpDir });
        const writeConnectionInfo = loadWriteConnectionInfo(sandbox);

        const result = writeConnectionInfo(8770, "new-token");

        assert.equal(result, "thunderbird-mcp/connection.json");
        assert.deepEqual(tmpDir.removeCalls, connectionFileExists ? [false] : []);
        assert.deepEqual(tmpDir.events, [
          ...(connectionFileExists ? ["remove"] : []),
          "open", "write", "close",
        ]);
        assert.deepEqual(opened, [{
          path: "thunderbird-mcp/connection.json",
          flags: 0x02 | 0x08 | 0x80,
          permissions: 0o600,
          behavior: 0,
          fileExisted: false,
        }]);
        assert.deepEqual(JSON.parse(written[0]), { port: 8770, token: "new-token", pid: 4242 });
      });
    }
  }
});

describe("writeConnectionInfo profile metadata", () => {
  const POSIX_PROFILE_DIRS = {
    ProfD: { path: "/profiles/current", normalizedPath: "/storage/mail/current" },
    DefProfRt: { path: "/config/thunderbird", normalizedPath: "/storage/mail/profiles" },
  };
  const WINDOWS_PROFILE_DIRS = {
    ProfD: { path: "C:\\Profiles\\CURRENT~1", normalizedPath: "C:\\Profiles\\current", target: "D:\\Mail\\current" },
    DefProfRt: { path: "C:\\Profiles", target: "D:\\Mail\\profiles" },
  };
  const PROFILE_PREFIX = "/profiles/";

  for (const os of ["Linux", "Darwin"]) {
    it(`includes lexical and canonical locations for both profile directories on ${os}`, () => {
      const profileDirs = POSIX_PROFILE_DIRS;
      const { sandbox, written } = makeSandbox({ os, tmpDir: makeTmpDir({ permissions: 0o700 }), profileDirs });

      loadWriteConnectionInfo(sandbox)(8770, "token");

      assert.deepEqual(JSON.parse(written[0]), {
        port: 8770, token: "token", pid: 4242,
        profileRoots: [profileDirs.ProfD.path, profileDirs.ProfD.normalizedPath,
          profileDirs.DefProfRt.path, profileDirs.DefProfRt.normalizedPath],
      });
    });
  }

  it("includes the resolved Windows targets alongside lexical profile locations", () => {
    const profileDirs = WINDOWS_PROFILE_DIRS;
    const { sandbox, written } = makeSandbox({ os: "WINNT", tmpDir: makeTmpDir({}), profileDirs });

    loadWriteConnectionInfo(sandbox)(8770, "token");

    assert.deepEqual(JSON.parse(written[0]).profileRoots,
      [profileDirs.ProfD.path, profileDirs.ProfD.target, profileDirs.DefProfRt.path, profileDirs.DefProfRt.target]);
  });

  it("includes a normalized Windows location when no directory redirects it", () => {
    const profileDirs = { ProfD: { path: WINDOWS_PROFILE_DIRS.ProfD.path, normalizedPath: WINDOWS_PROFILE_DIRS.ProfD.normalizedPath } };
    const { sandbox, written } = makeSandbox({ os: "WINNT", tmpDir: makeTmpDir({}), profileDirs });

    loadWriteConnectionInfo(sandbox)(8770, "token");

    assert.deepEqual(JSON.parse(written[0]).profileRoots, [profileDirs.ProfD.path, profileDirs.ProfD.normalizedPath]);
  });

  const WINDOWS_PROFILE_PARENT = WINDOWS_PROFILE_DIRS.DefProfRt.path;
  const WINDOWS_PROFILE_CHILD = WINDOWS_PROFILE_PARENT + "\\custom";
  const WINDOWS_TARGET_PARENT = "D:\\Aliases";
  const WINDOWS_TARGET_ROOT = WINDOWS_TARGET_PARENT + "\\Profiles";
  const WINDOWS_FINAL_PARENT = "E:\\Archive";

  it("resolves a Windows profile through an ancestor junction", () => {
    const targets = { [WINDOWS_PROFILE_PARENT]: WINDOWS_TARGET_ROOT };
    const profileDirs = {
      ProfD: { path: WINDOWS_PROFILE_CHILD, targets },
      DefProfRt: { path: WINDOWS_PROFILE_PARENT, targets },
    };
    const { sandbox, written } = makeSandbox({ os: "WINNT", tmpDir: makeTmpDir({}), profileDirs });

    loadWriteConnectionInfo(sandbox)(8770, "token");

    assert.deepEqual(JSON.parse(written[0]).profileRoots,
      [WINDOWS_PROFILE_CHILD, WINDOWS_TARGET_ROOT + "\\custom", WINDOWS_PROFILE_PARENT, WINDOWS_TARGET_ROOT]);
  });

  it("resolves another ancestor junction inside the first Windows target", () => {
    const targets = { [WINDOWS_PROFILE_PARENT]: WINDOWS_TARGET_ROOT, [WINDOWS_TARGET_PARENT]: WINDOWS_FINAL_PARENT };
    const profileDirs = {
      ProfD: { path: WINDOWS_PROFILE_CHILD, targets },
      DefProfRt: { path: WINDOWS_PROFILE_PARENT, targets },
    };
    const { sandbox, written } = makeSandbox({ os: "WINNT", tmpDir: makeTmpDir({}), profileDirs });

    loadWriteConnectionInfo(sandbox)(8770, "token");

    assert.deepEqual(JSON.parse(written[0]).profileRoots,
      [WINDOWS_PROFILE_CHILD, WINDOWS_FINAL_PARENT + "\\Profiles\\custom", WINDOWS_PROFILE_PARENT, WINDOWS_FINAL_PARENT + "\\Profiles"]);
  });

  for (const [name, extra] of [
    ["unreadable ancestor", { unreadablePaths: [WINDOWS_PROFILE_PARENT] }],
    ["unreadable target ancestor", { targets: { [WINDOWS_PROFILE_PARENT]: WINDOWS_TARGET_ROOT }, unreadablePaths: [WINDOWS_TARGET_PARENT] }],
    ["invalid target", { targets: { [WINDOWS_PROFILE_PARENT]: "relative-folder" } }],
    ["cyclic target", { targets: { [WINDOWS_PROFILE_PARENT]: WINDOWS_TARGET_ROOT, [WINDOWS_TARGET_PARENT]: "c:/" } }],
  ]) {
    it(`keeps connection discovery working with only a lexical root for a Windows ${name}`, () => {
      const profileDirs = { ProfD: { path: WINDOWS_PROFILE_CHILD, ...extra } };
      const { sandbox, written } = makeSandbox({ os: "WINNT", tmpDir: makeTmpDir({}), profileDirs });

      loadWriteConnectionInfo(sandbox)(8770, "token");

      assert.deepEqual(JSON.parse(written[0]), {
        port: 8770, token: "token", pid: 4242, profileRoots: [WINDOWS_PROFILE_CHILD],
      });
    });
  }

  for (const os of ["Linux", "WINNT"]) {
    it(`retains the legacy connection fields when profile directories are unavailable on ${os}`, () => {
      const { sandbox, written } = makeSandbox({ os, tmpDir: makeTmpDir({ permissions: 0o700 }) });

      loadWriteConnectionInfo(sandbox)(8770, "token");

      assert.deepEqual(JSON.parse(written[0]), { port: 8770, token: "token", pid: 4242 });
    });
  }

  it("keeps a lexical profile location when normalization fails and still collects the other directory", () => {
    const profileDirs = {
      ProfD: { ...POSIX_PROFILE_DIRS.ProfD, normalizeError: new Error("Unavailable target") },
      DefProfRt: POSIX_PROFILE_DIRS.DefProfRt,
    };
    const { sandbox, written } = makeSandbox({ os: "Linux", tmpDir: makeTmpDir({ permissions: 0o700 }), profileDirs });

    loadWriteConnectionInfo(sandbox)(8770, "token");

    assert.deepEqual(JSON.parse(written[0]).profileRoots,
      [profileDirs.ProfD.path, profileDirs.DefProfRt.path, profileDirs.DefProfRt.normalizedPath]);
  });

  it("keeps the lexical Windows location when its target cannot be resolved", () => {
    const profileDirs = { ProfD: { ...WINDOWS_PROFILE_DIRS.ProfD, readable: false } };
    const { sandbox, written } = makeSandbox({ os: "WINNT", tmpDir: makeTmpDir({}), profileDirs });

    loadWriteConnectionInfo(sandbox)(8770, "token");

    assert.deepEqual(JSON.parse(written[0]).profileRoots, [profileDirs.ProfD.path]);
  });

  it("deduplicates locations shared by lexical and canonical directories", () => {
    const shared = POSIX_PROFILE_DIRS.DefProfRt.normalizedPath;
    const profileDirs = {
      ProfD: { ...POSIX_PROFILE_DIRS.ProfD, normalizedPath: shared },
      DefProfRt: { path: shared, normalizedPath: shared },
    };
    const { sandbox, written } = makeSandbox({ os: "Linux", tmpDir: makeTmpDir({ permissions: 0o700 }), profileDirs });

    loadWriteConnectionInfo(sandbox)(8770, "token");

    assert.deepEqual(JSON.parse(written[0]).profileRoots, [profileDirs.ProfD.path, shared]);
  });

  for (const extraBytes of [0, 1]) {
    it(`${extraBytes ? "omits" : "retains"} metadata at ${CONNECTION_READ_LIMIT + extraBytes} encoded bytes`, () => {
      const empty = { port: 8770, token: "token", pid: 4242, profileRoots: [PROFILE_PREFIX] };
      const root = PROFILE_PREFIX + "p".repeat(CONNECTION_READ_LIMIT + extraBytes - Buffer.byteLength(JSON.stringify(empty)));
      const { sandbox, written } = makeSandbox({ os: "Linux", tmpDir: makeTmpDir({ permissions: 0o700 }),
        profileDirs: { ProfD: { path: root } } });

      loadWriteConnectionInfo(sandbox)(8770, "token");

      assert.ok(Buffer.byteLength(written[0]) <= CONNECTION_READ_LIMIT);
      assert.deepEqual(JSON.parse(written[0]).profileRoots, extraBytes ? undefined : [root]);
    });
  }

  for (const oversizedFirst of [false, true]) {
    it(`keeps usable roots around an oversized entry (oversizedFirst=${oversizedFirst})`, () => {
      const oversized = PROFILE_PREFIX + "é".repeat(2100);
      const current = POSIX_PROFILE_DIRS.ProfD.path;
      const profiles = POSIX_PROFILE_DIRS.DefProfRt.path;
      const profileDirs = {
        ProfD: { path: oversizedFirst ? oversized : current, normalizedPath: oversizedFirst ? current : oversized },
        DefProfRt: { path: profiles },
      };
      const { sandbox, written } = makeSandbox({ os: "Linux", tmpDir: makeTmpDir({ permissions: 0o700 }), profileDirs });
      const connectionPath = loadWriteConnectionInfo(sandbox)(8770, "token");
      const data = JSON.parse(written[0]);
      assert.deepEqual(data.profileRoots, [current, profiles]);
      assert.ok(Buffer.byteLength(written[0]) <= CONNECTION_READ_LIMIT);
      assert.equal(sandbox.ensureFreshConnectionInfo({
        port: data.port, token: data.token, expectedPid: data.pid,
        expectedProfileRoots: sandbox.getConnectionInfoData(data.port, data.token).profileRoots,
        readConnectionInfo: () => ({ path: connectionPath, data }),
        writeConnectionInfo() { assert.fail("Retained metadata must not cause repeated writes"); },
      }), connectionPath);
    });
  }

  it("retains the roots that fit when the combined metadata exceeds the limit", () => {
    const empty = { port: 8770, token: "token", pid: 4242, profileRoots: [PROFILE_PREFIX] };
    const current = PROFILE_PREFIX + "p".repeat(CONNECTION_READ_LIMIT - Buffer.byteLength(JSON.stringify(empty)));
    const { sandbox, written } = makeSandbox({ os: "Linux", tmpDir: makeTmpDir({ permissions: 0o700 }),
      profileDirs: { ProfD: { path: current }, DefProfRt: { path: POSIX_PROFILE_DIRS.DefProfRt.path } } });
    loadWriteConnectionInfo(sandbox)(8770, "token");
    assert.equal(Buffer.byteLength(written[0]), CONNECTION_READ_LIMIT);
    assert.deepEqual(JSON.parse(written[0]).profileRoots, [current]);
  });

  for (const os of ["Linux", "WINNT"]) {
    it(`does not let unsupported roots consume the metadata budget on ${os}`, () => {
      const current = os === "WINNT" ? WINDOWS_PROFILE_DIRS.ProfD.path : POSIX_PROFILE_DIRS.ProfD.path;
      const empty = { port: 8770, token: "token", pid: 4242, profileRoots: ["//server/share/"] };
      const unsupported = empty.profileRoots[0] + "p".repeat(CONNECTION_READ_LIMIT - Buffer.byteLength(JSON.stringify(empty)));
      const { sandbox, written } = makeSandbox({ os, tmpDir: makeTmpDir({ permissions: 0o700 }),
        profileDirs: { ProfD: { path: unsupported }, DefProfRt: { path: current } } });
      loadWriteConnectionInfo(sandbox)(8770, "token");
      assert.deepEqual(JSON.parse(written[0]).profileRoots, [current]);
      assert.ok(Buffer.byteLength(written[0]) <= CONNECTION_READ_LIMIT);
    });
  }

  it("omits oversized UTF-8 metadata and does not repeatedly refresh the resulting legacy shape", () => {
    const root = PROFILE_PREFIX + "é".repeat(2100);
    const candidate = JSON.stringify({ port: 8770, token: "token", pid: 4242, profileRoots: [root] });
    assert.ok(candidate.length < CONNECTION_READ_LIMIT);
    assert.ok(Buffer.byteLength(candidate) > CONNECTION_READ_LIMIT);
    const { sandbox, written } = makeSandbox({ os: "Linux", tmpDir: makeTmpDir({ permissions: 0o700 }),
      profileDirs: { ProfD: { path: root } } });
    const connectionPath = loadWriteConnectionInfo(sandbox)(8770, "token");
    const data = JSON.parse(written[0]);
    assert.deepEqual(data, { port: 8770, token: "token", pid: 4242 });
    assert.ok(Buffer.byteLength(written[0]) <= CONNECTION_READ_LIMIT);

    const result = sandbox.ensureFreshConnectionInfo({
      port: data.port, token: data.token, expectedPid: data.pid,
      expectedProfileRoots: sandbox.getConnectionInfoData(data.port, data.token).profileRoots,
      readConnectionInfo: () => ({ path: connectionPath, data }),
      writeConnectionInfo() { assert.fail("Omitted metadata must not cause repeated writes"); },
    });
    assert.equal(result, connectionPath);
  });
});
