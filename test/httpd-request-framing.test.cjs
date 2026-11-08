"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const { HTTPD_PORT, createHttpdHarness, installMcpHandler } = require("./helpers/httpd.cjs");

const HEADER_LIMIT = 64 * 1024;
const BODY_LIMIT = 32 * 1024 * 1024;
const READ_SIZE = 8192;
const HOST = `localhost:${HTTPD_PORT}`;

// Existing Host tests assert helper results; these helpers assert actual wire
// responses from the whole module, including a subsequent request on the server.
function requestHead(headers = "", { method = "POST", target = "/", version = "1.1", host = HOST } = {}) {
  return `${method} ${target} HTTP/${version}\r\n${host === null ? "" : `Host: ${host}\r\n`}${headers}\r\n`;
}

function assertResponse(exchange, code) {
  const match = /^HTTP\/1\.[01] (\d{3}) [^\r\n]+\r\n/.exec(exchange.wire);
  assert.ok(match, `missing serialized status line: ${JSON.stringify(exchange.wire)}`);
  assert.equal(Number(match[1]), code, exchange.wire.slice(0, 100));
  const end = exchange.wire.indexOf("\r\n\r\n");
  assert.ok(end > 0, "response headers must be terminated");
  const headers = exchange.wire.slice(0, end);
  assert.match(headers, /\r\nconnection: close/i);
  const length = /\r\ncontent-length: (\d+)/i.exec(headers);
  if (length) assert.equal(exchange.wire.length - end - 4, Number(length[1]), "complete serialized body");
  if (code >= 400) assert.ok(exchange.wire.length > end + 4, "default error page was written");
  assert.equal(exchange.connection._closed, true, "httpd closed the request connection");
  assert.equal(exchange.output.closed, true);
}

function assertListening(harness) {
  assert.equal(harness.server._doQuit, false);
  assert.equal(harness.socket.closed, false);
  assert.equal(harness.server._socket, harness.socket);
  assertResponse(harness.request(requestHead()), 200);
  assert.equal(harness.server._socket, harness.socket);
}

function assertRejected(harness, raw, code, internalCode = code) {
  const exchange = harness.request(raw);
  assertResponse(exchange, code);
  assert.equal(harness.requests.length, 0, "invalid framing never reaches the handler");
  assert.equal(exchange.errors.length, 1);
  assert.ok(exchange.errors[0] instanceof harness.sandbox.HttpError);
  assert.equal(exchange.errors[0].code, internalCode);
  assert.deepEqual(harness.logs, [], "expected client errors are not unexpected failures");
  assertListening(harness);
  return exchange;
}

describe("httpd request framing over real Connection/Response/error pages", () => {
  for (const value of ["-1", "-0", "+1", "abc", "", " \t ", "1e3", "1.5", "0x10", "1junk", "1 2", "1,1", "1,2", "1;", "1\u00a0", "NaN", "Infinity"]) {
    it(`rejects non-decimal Content-Length ${JSON.stringify(value)}`, () => {
      assertRejected(createHttpdHarness(), requestHead(`Content-Length: ${value}\r\n`) + "ignored", 400);
    });
  }

  for (const second of ["2", "3"]) {
    it(`rejects repeated Content-Length fields (2 and ${second})`, () => {
      assertRejected(createHttpdHarness(), requestHead(`Content-Length: 2\r\ncontent-length: ${second}\r\n`) + "{}", 400);
    });
  }

  for (const value of ["0", "0000", "2", "0002", " \t2\t "]) {
    it(`accepts decimal Content-Length ${JSON.stringify(value)} and preserves the body`, () => {
      const harness = createHttpdHarness();
      const body = Number(value) ? "{}" : "";
      const exchange = harness.request(requestHead(`Content-Length: ${value}\r\n`) + body);
      assertResponse(exchange, 200);
      const stream = harness.requests[0].bodyInputStream;
      assert.equal(stream.read(stream.available()).toString("latin1"), body);
      assertListening(harness);
    });
  }

  it("treats absent Content-Length as an empty body", () => {
    const harness = createHttpdHarness();
    assertResponse(harness.request(requestHead()), 200);
    assert.equal(harness.requests[0].bodyInputStream.available(), 0);
    assertListening(harness);
  });

  it("accepts exactly 32 MiB, waiting for the final byte", () => {
    const harness = createHttpdHarness();
    const exchange = harness.request(requestHead(`Content-Length: ${BODY_LIMIT}\r\n`));
    const block = "x".repeat(1024 * 1024);
    for (let i = 0; i < 31; i++) exchange.feed(block);
    exchange.feed(block.slice(1));
    assert.equal(harness.requests.length, 0);
    assert.equal(exchange.wire, "");
    assert.equal(exchange.reader._contentLength, 1);
    exchange.feed("z");
    assertResponse(exchange, 200);
    const stream = harness.requests[0].bodyInputStream;
    assert.equal(stream.available(), BODY_LIMIT);
    const body = stream.read(stream.available());
    assert.equal(body[0], "x".charCodeAt(0));
    assert.equal(body.at(-1), "z".charCodeAt(0));
    assert.ok(exchange.peakBuffered <= READ_SIZE);
    assertListening(harness);
  });

  for (const value of [String(BODY_LIMIT + 1), "9".repeat(400)]) {
    it(`rejects oversized decimal Content-Length (${value.length} digits) before the body`, () => {
      // Vendored httpd has no 413 page: its existing x00 fallback is a wire 400.
      assertRejected(createHttpdHarness(), requestHead(`Content-Length: ${value}\r\n`), 400, 413);
    });
  }

  for (const value of ["", "chunked", "identity", "gzip, chunked", "unknown"]) {
    for (const contentLength of ["", "Content-Length: 2\r\n"]) {
      it(`refuses Transfer-Encoding ${JSON.stringify(value)} ${contentLength ? "with" : "without"} Content-Length`, () => {
        assertRejected(createHttpdHarness(), requestHead(`Transfer-Encoding: ${value}\r\n${contentLength}`) + "2\r\n{}\r\n0\r\n\r\n", 501);
      });
    }
  }

  it("refuses Transfer-Encoding even with an invalid Content-Length", () => {
    assertRejected(createHttpdHarness(), requestHead("Content-Length: -1\r\ntRaNsFeR-EnCoDiNg: chunked\r\n"), 501);
  });

  it("refuses duplicated Transfer-Encoding fields", () => {
    assertRejected(createHttpdHarness(), requestHead("Transfer-Encoding: gzip\r\nTransfer-Encoding: chunked\r\n"), 501);
  });

  for (const mode of ["loopback", "listenAll"]) {
    for (const host of ["localhost:65536", "localhost:99999", "[::1]:99999", "proxy.example:99999", `bad host:${HTTPD_PORT}`]) {
      it(`answers bad Host ${host} with 400 and keeps ${mode} listening`, () => {
        assertRejected(createHttpdHarness({ mode }), requestHead("", { host }), 400);
      });
    }

    it(`preserves mapped Host behavior in ${mode}`, () => {
      const harness = createHttpdHarness({ mode });
      const raw = requestHead("", { host: "reverse-proxy.example:443" });
      if (mode === "loopback") {
        assertRejected(harness, raw, 400);
      } else {
        assertResponse(harness.request(raw), 200);
        assert.equal(harness.requests[0].host, harness.server.identity.primaryHost);
        assert.equal(harness.requests[0].port, HTTPD_PORT);
        assertListening(harness);
      }
    });

    for (const version of ["1.0", "1.1", "2.0"]) {
      for (const target of [`localhost:${HTTPD_PORT}`, "!:1", "localhost:99999"]) {
        it(`refuses CONNECT ${target} HTTP/${version} in ${mode}`, () => {
          assertRejected(createHttpdHarness({ mode }), requestHead("", { method: "CONNECT", target, version }), 501);
        });
      }
    }
  }

  it("serializes 501 for an unsupported HTTP version", () => {
    assertRejected(createHttpdHarness(), requestHead("", { version: "0.9" }), 501);
  });

  it("serializes 400 for a malformed request line", () => {
    assertRejected(createHttpdHarness(), "not a request line\r\n", 400);
  });

  it("accepts HTTP/1.0 without Host", () => {
    const harness = createHttpdHarness();
    assertResponse(harness.request(requestHead("", { version: "1.0", host: null })), 200);
    assertListening(harness);
  });
});

describe("httpd cumulative request-line/header budget", () => {
  it("accepts byte-fragmented lines, CRLFs and body", () => {
    const harness = createHttpdHarness();
    const raw = requestHead("X-Example: one\r\n two\r\nContent-Length: 2\r\n") + "{}";
    assertResponse(harness.request([...raw]), 200);
    const req = harness.requests[0];
    assert.equal(req.getHeader("X-Example"), "one two");
    assert.equal(req.bodyInputStream.read(2).toString(), "{}");
    assertListening(harness);
  });

  for (const delta of [-1, 0, 1]) {
    it(`enforces the cumulative boundary at 64 KiB ${delta < 0 ? "- 1" : delta ? "+ 1" : "exactly"}`, () => {
      const harness = createHttpdHarness();
      const base = requestHead("X-Pad: \r\n");
      const raw = requestHead(`X-Pad: ${"a".repeat(HEADER_LIMIT - base.length + delta)}\r\n`);
      assert.equal(raw.length, HEADER_LIMIT + delta);
      if (delta > 0) {
        assertRejected(harness, raw, 400);
      } else {
        assertResponse(harness.request(raw), 200);
        assertListening(harness);
      }
    });
  }

  for (const prefix of ["POST /", requestHead().slice(0, -2) + "X-Pad: "]) {
    for (const fragmented of [false, true]) {
      it(`bounds an unterminated ${prefix.includes("Host:") ? "header" : "request"} line (${fragmented ? "fragmented" : "coalesced"})`, () => {
        const chunks = fragmented
          ? [prefix, ...Array(80).fill("a".repeat(1024))]
          : prefix + "a".repeat(2 * HEADER_LIMIT);
        const exchange = assertRejected(createHttpdHarness(), chunks, 400);
        assert.ok(exchange.peakBuffered <= HEADER_LIMIT + READ_SIZE);
        assert.equal(exchange.reader._data, null);
      });
    }
  }

  it("counts many small parsed headers even when the line buffer is empty", () => {
    const harness = createHttpdHarness();
    const prefix = requestHead().slice(0, -2);
    const exchange = harness.request(prefix);
    let bytes = prefix.length;
    let count = 0;
    while (!exchange.connection._closed && count < 10000) {
      const line = `X-${count++}: a\r\n`;
      bytes += line.length;
      exchange.feed(line);
      if (bytes <= HEADER_LIMIT) {
        assert.equal(exchange.reader._data._data.length, 0);
        assert.equal(exchange.wire, "");
      }
    }
    assert.ok(count > 1000);
    assert.ok(bytes > HEADER_LIMIT);
    assertResponse(exchange, 400);
    assert.equal(harness.requests.length, 0);
    assert.ok(exchange.errors[0] instanceof harness.sandbox.HttpError);
    assert.ok(exchange.peakBuffered < 128, "the limit is cumulative, not a line-buffer limit");
    assertListening(harness);
  });

  it("counts ignored leading empty lines", () => {
    assertRejected(createHttpdHarness(), "\r\n".repeat(HEADER_LIMIT / 2 + 1) + requestHead(), 400);
  });

  it("counts folded header continuation lines", () => {
    const prefix = requestHead().slice(0, -2) + "X-Pad: a\r\n";
    assertRejected(createHttpdHarness(), [prefix, ...Array(70).fill(" " + "a".repeat(1024) + "\r\n")], 400);
  });

  it("does not charge a coalesced body beyond 64 KiB to the header budget", () => {
    const harness = createHttpdHarness();
    const body = "x".repeat(2 * HEADER_LIMIT);
    const raw = requestHead(`Content-Length: ${body.length}\r\n`) + body;
    const exchange = harness.request(raw);
    assertResponse(exchange, 200);
    const stream = harness.requests[0].bodyInputStream;
    assert.equal(stream.read(stream.available()).toString(), body);
    assert.ok(exchange.peakBuffered <= READ_SIZE);
    assertListening(harness);
  });

  it("accepts a full header budget when the final LF arrives with a large body", () => {
    const harness = createHttpdHarness();
    const body = "y".repeat(2 * HEADER_LIMIT);
    const fields = `Content-Length: ${body.length}\r\nX-Pad: `;
    const base = requestHead(fields + "\r\n");
    const head = requestHead(fields + "a".repeat(HEADER_LIMIT - base.length) + "\r\n");
    assert.equal(head.length, HEADER_LIMIT);
    const exchange = harness.request([head.slice(0, -1), head.slice(-1) + body]);
    assertResponse(exchange, 200);
    const stream = harness.requests[0].bodyInputStream;
    assert.equal(stream.read(stream.available()).toString(), body);
    assertListening(harness);
  });
});

describe("httpd unexpected failures stay connection-local", () => {
  for (const stage of ["socket available", "socket read", "socket wait", "request line", "headers", "body", "identity"]) {
    it(`logs and serializes 500 for an unexpected ${stage} failure`, () => {
      const harness = createHttpdHarness();
      const exchange = harness.openRequest();
      const failure = new Error(`injected ${stage} failure`);
      const fail = () => { throw failure; };
      if (stage === "socket available") exchange.input.available = fail;
      if (stage === "socket read") exchange.input.read = fail;
      if (stage === "socket wait") exchange.input.asyncWait = fail;
      if (stage === "request line") exchange.reader._data.readLine = fail;
      if (stage === "headers") exchange.reader._parseHeaders = fail;
      if (stage === "body") exchange.reader._metadata._bodyOutputStream.write = fail;
      const getScheme = harness.server.identity.getScheme;
      if (stage === "identity") harness.server.identity.getScheme = fail;
      exchange.feed(stage === "socket wait" ? "POST /" : requestHead("Content-Length: 2\r\n") + "{}");
      harness.server.identity.getScheme = getScheme;
      assertResponse(exchange, 500);
      assert.equal(harness.requests.length, 0);
      assert.deepEqual(exchange.errors, [failure]);
      assert.ok(harness.logs.some(args => args.includes(failure)));
      assertListening(harness);
    });
  }

  for (const stage of ["socket", "parser"]) {
    it(`answers even a null thrown by the ${stage}`, () => {
      const harness = createHttpdHarness();
      const exchange = harness.openRequest();
      if (stage === "socket") exchange.input.available = () => { throw null; };
      else exchange.reader._data.readLine = () => { throw null; };
      exchange.feed(requestHead());
      assertResponse(exchange, 500);
      assert.deepEqual(exchange.errors, [null]);
      assert.ok(harness.logs.some(args => args.includes(null)));
      assertListening(harness);
    });
  }

  it("contains a real ServerIdentity validation exception from an absolute URI", () => {
    const harness = createHttpdHarness();
    // Inject a URI-service result; this does not assert Gecko accepts this host.
    harness.sandbox.Services.io.newURI = () => ({
      pathQueryRef: "/", scheme: "http", asciiHost: "a_b", port: HTTPD_PORT,
    });
    const exchange = harness.request(requestHead("", { target: `http://a_b:${HTTPD_PORT}/` }));
    assertResponse(exchange, 500);
    assert.equal(exchange.errors[0].result, harness.sandbox.Cr.NS_ERROR_ILLEGAL_VALUE);
    assert.ok(harness.logs.some(args => args.includes(exchange.errors[0])));
    assertListening(harness);
  });

  it("handles a genuine closed input without stopping the listener", () => {
    const harness = createHttpdHarness();
    const exchange = harness.request("POST /");
    exchange.input.close();
    harness.drain();
    assert.equal(exchange.connection._closed, true);
    assert.equal(exchange.wire, "");
    assertListening(harness);
  });

  it("logs an error-handler failure and aborts only that connection", () => {
    const harness = createHttpdHarness();
    const failure = new Error("injected error-page failure");
    let first = true;
    harness.server.registerPathHandler("/", (_req, res) => {
      if (first) {
        first = false;
        const error = new harness.sandbox.HttpError(400, "Bad Request");
        error.customErrorHandling = () => { throw failure; };
        throw error;
      }
      res.write("ok");
    });
    const exchange = harness.request(requestHead());
    assert.equal(exchange.connection._closed, true);
    assert.ok(harness.logs.some(args => args.includes(failure)));
    assertListening(harness);
  });
});

describe("httpd to the production MCP HTTP handler", () => {
  for (const mode of ["loopback", "listenAll"]) {
    for (const version of ["1.0", "1.1"]) {
      it(`validates supplied Host independently of authenticated HTTP/${version} authority (${mode})`, () => {
        const harness = createHttpdHarness({ mode });
        const token = "httpd-host-test-token";
        installMcpHandler(harness, token);
        const absoluteHost = "127.0.0.1";
        const target = version === "1.0" ? "/" : `http://${absoluteHost}:${HTTPD_PORT}/`;
        // Reuse the fixed URI-service stub pattern; authority checks stay real.
        harness.sandbox.Services.io.newURI = uri => {
          assert.equal(uri, target);
          return { pathQueryRef: "/", scheme: "http", asciiHost: absoluteHost, port: HTTPD_PORT };
        };
        const body = JSON.stringify({ jsonrpc: "2.0", id: "host-check", method: "resources/list" });
        const fields = `Authorization: Bearer ${token}\r\nContent-Length: ${Buffer.byteLength(body)}\r\n`;
        const reply = JSON.stringify({ jsonrpc: "2.0", id: "host-check", result: { resources: [] } });

        for (const host of ["localhost:65536", `bad host:${HTTPD_PORT}`, "localhost:not-a-port"]) {
          const rejected = harness.request(requestHead(fields, { version, target, host }) + body);
          assertResponse(rejected, 400);
          assert.equal(rejected.errors.length, 1);
          assert.ok(rejected.errors[0] instanceof harness.sandbox.HttpError);
          assert.equal(rejected.errors[0].code, 400);
          assert.equal(harness.server._doQuit, false);

          const followup = harness.request(requestHead(fields) + body);
          assertResponse(followup, 200);
          assert.ok(followup.wire.endsWith("\r\n\r\n" + reply), "authenticated follow-up reached MCP");
        }

        // A valid unknown Host must not override either authority. HTTP/1.0
        // also continues to accept no Host at all.
        const validHosts = ["unrecognized.example:65535"];
        if (version === "1.0") validHosts.push(null);
        for (const host of validHosts) {
          const accepted = harness.request(requestHead(fields, { version, target, host }) + body);
          assertResponse(accepted, 200);
          assert.ok(accepted.wire.endsWith("\r\n\r\n" + reply));
          assert.equal(accepted.connection.request.host,
            version === "1.0" ? harness.server.identity.primaryHost : absoluteHost);
          assert.equal(accepted.connection.request.port, HTTPD_PORT);
        }
        assert.equal(harness.server._doQuit, false);
        assert.equal(harness.server._socket, harness.socket);
        assert.equal(harness.socket.closed, false);
        assert.deepEqual(harness.logs, []);
      });
    }

    it(`authenticates and dispatches a real POST after a rejected request (${mode})`, () => {
      const harness = createHttpdHarness({ mode });
      assertRejected(harness, requestHead("Transfer-Encoding: chunked\r\n"), 501);
      const token = "httpd-test-token";
      installMcpHandler(harness, token);
      const body = Buffer.from(JSON.stringify({ jsonrpc: "2.0", id: "Zażółć 📬", method: "resources/list" }));
      const host = mode === "listenAll" ? "proxy.example:443" : HOST;
      const unauthenticated = harness.request(Buffer.concat([
        Buffer.from(requestHead(`Content-Length: ${body.length}\r\n`, { host })), body,
      ]));
      assertResponse(unauthenticated, 403);
      assert.match(unauthenticated.wire, /Invalid or missing auth token/);
      const authenticated = harness.request(Buffer.concat([
        Buffer.from(requestHead(`Authorization: Bearer ${token}\r\nContent-Length: ${body.length}\r\n`, { host })), body,
      ]));
      assertResponse(authenticated, 200);
      const responseBody = authenticated.wire.slice(authenticated.wire.indexOf("\r\n\r\n") + 4);
      assert.deepEqual(JSON.parse(Buffer.from(responseBody, "latin1").toString("utf8")), {
        jsonrpc: "2.0", id: "Zażółć 📬", result: { resources: [] },
      });
      assert.equal(harness.server._doQuit, false);
      assert.equal(harness.server._socket, harness.socket);
      assert.equal(harness.socket.closed, false);
    });
  }
});
