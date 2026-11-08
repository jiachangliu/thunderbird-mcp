"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const { makeDirectSendAttachments, assertNoAttachmentEffects } = require("./helpers/direct-send-attachments.cjs");

describe("Direct send attachment side-effect instrumentation", () => {
  it("runs production path and inline conversion with observable file access and creation", () => {
    const fixture = makeDirectSendAttachments();
    assertNoAttachmentEffects(fixture.calls);
    const { descs, failed } = fixture.convert(fixture.entries);
    assert.equal(failed.length, 0);
    assert.equal(descs.length, 2);
    assert.equal(descs[0].url, "file:///fixture/Documents/report.txt");
    assert.equal(descs[0].name, "report.txt");
    assert.equal(descs[0].size, 6);
    assert.match(descs[1].url, /^file:\/\/\/fixture\/tmp\/thunderbird-mcp\/attachments\/\d+_1_inline\.txt$/);
    assert.equal(descs[1].name, "inline.txt");
    assert.equal(descs[1].contentType, "text/plain");
    assert.equal(descs[1].size, 1);
    assert.ok(fixture.calls.fileAccesses.some(([operation, value]) =>
      operation === "createLocalFile" && value === fixture.entries[0]));
    assert.deepEqual(fixture.calls.directoryAccesses, ["ProfD", "DefProfRt", "TmpD"]);
    assert.equal(fixture.calls.directoryCreates.length, 1);
    assert.equal(fixture.calls.tempCreates.length, 1);
    assert.equal(fixture.calls.writes.length, 1);
    assert.equal(fixture.calls.writes[0].length, 1);
    assert.equal(fixture.calls.openStreams, 0);
    assert.equal(fixture.calls.nativeCreates, 2);
    assert.deepEqual(Array.from(fixture.calls.tempFiles.values(), bytes => Array.from(bytes)), [[65]]);
    const native = fixture.toNative(descs);
    assert.strictEqual(native[0], descs[0].msgAttachment);
    assert.strictEqual(native[1], descs[1].msgAttachment);
    assert.equal(fixture.calls.nativeCreates, 2, "Production conversion reuses prepared native attachments");
    assert.throws(() => assertNoAttachmentEffects(fixture.calls), /Encryption refusal/);
  });
});
