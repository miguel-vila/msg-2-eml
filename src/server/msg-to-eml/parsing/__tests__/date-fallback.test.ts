import assert from "node:assert";
import { describe, it } from "node:test";
import {
  type Msg,
  PidTagClientSubmitTime,
  PidTagMessageDeliveryTime,
  PidTagTransportMessageHeaders,
  type PropertyTag,
} from "msg-parser";
import { parseDateFromTransportHeaders, parseMsgFromMsg } from "../msg.js";

// Helper to create a mock Msg with specific property values
interface MockPropertyValue {
  tag: PropertyTag;
  value: unknown;
}

function createMockMsg(properties: MockPropertyValue[]): Msg {
  return {
    getProperty: <T>(propertyId: PropertyTag): T | undefined => {
      const found = properties.find((p) => p.tag === propertyId);
      return found?.value as T | undefined;
    },
    recipients: () => [],
    attachments: () => [],
    embeddedMessages: () => [],
  } as unknown as Msg;
}

describe("parseDateFromTransportHeaders", () => {
  it("should parse a standard RFC 5322 Date header", () => {
    const headers = "From: test@example.com\r\nDate: Fri, 15 Mar 2024 14:30:45 +0000\r\nSubject: Test\r\n";
    const result = parseDateFromTransportHeaders(headers);
    assert.ok(result);
    assert.strictEqual(result.toISOString(), "2024-03-15T14:30:45.000Z");
  });

  it("should parse a Date header with timezone offset", () => {
    const headers = "Date: Mon, 01 Jan 2024 09:00:00 -0500\r\nFrom: test@example.com\r\n";
    const result = parseDateFromTransportHeaders(headers);
    assert.ok(result);
    assert.strictEqual(result.toISOString(), "2024-01-01T14:00:00.000Z");
  });

  it("should handle folded Date headers", () => {
    const headers = "Date: Fri, 15 Mar 2024\r\n 14:30:45 +0000\r\nFrom: test@example.com\r\n";
    const result = parseDateFromTransportHeaders(headers);
    assert.ok(result);
    assert.strictEqual(result.toISOString(), "2024-03-15T14:30:45.000Z");
  });

  it("should handle tab-folded Date headers", () => {
    const headers = "Date: Fri, 15 Mar 2024\r\n\t14:30:45 +0000\r\nFrom: test@example.com\r\n";
    const result = parseDateFromTransportHeaders(headers);
    assert.ok(result);
    assert.strictEqual(result.toISOString(), "2024-03-15T14:30:45.000Z");
  });

  it("should return undefined when no Date header exists", () => {
    const headers = "From: test@example.com\r\nSubject: Test\r\n";
    const result = parseDateFromTransportHeaders(headers);
    assert.strictEqual(result, undefined);
  });

  it("should return undefined for an empty Date header value", () => {
    const headers = "Date: \r\nFrom: test@example.com\r\n";
    const result = parseDateFromTransportHeaders(headers);
    assert.strictEqual(result, undefined);
  });

  it("should return undefined for an invalid date string", () => {
    const headers = "Date: not-a-real-date\r\nFrom: test@example.com\r\n";
    const result = parseDateFromTransportHeaders(headers);
    assert.strictEqual(result, undefined);
  });

  it("should be case-insensitive for the Date header name", () => {
    const headers = "DATE: Fri, 15 Mar 2024 14:30:45 +0000\r\n";
    const result = parseDateFromTransportHeaders(headers);
    assert.ok(result);
    assert.strictEqual(result.toISOString(), "2024-03-15T14:30:45.000Z");
  });

  it("should handle LF-only line endings", () => {
    const headers = "From: test@example.com\nDate: Fri, 15 Mar 2024 14:30:45 +0000\nSubject: Test\n";
    const result = parseDateFromTransportHeaders(headers);
    assert.ok(result);
    assert.strictEqual(result.toISOString(), "2024-03-15T14:30:45.000Z");
  });
});

describe("parseMsgFromMsg date fallback", () => {
  const dummyMsgToEml = () => "";

  it("should use PidTagMessageDeliveryTime when available", () => {
    const deliveryDate = new Date("2024-03-15T14:30:45Z");
    const submitDate = new Date("2024-03-15T14:29:00Z");
    const msg = createMockMsg([
      { tag: PidTagMessageDeliveryTime, value: deliveryDate },
      { tag: PidTagClientSubmitTime, value: submitDate },
    ]);

    const result = parseMsgFromMsg(msg, dummyMsgToEml);
    assert.strictEqual(result.date, deliveryDate);
  });

  it("should fall back to PidTagClientSubmitTime when delivery time is missing", () => {
    const submitDate = new Date("2024-03-15T14:29:00Z");
    const msg = createMockMsg([{ tag: PidTagClientSubmitTime, value: submitDate }]);

    const result = parseMsgFromMsg(msg, dummyMsgToEml);
    assert.strictEqual(result.date, submitDate);
  });

  it("should fall back to Date from transport headers when both MAPI dates are missing", () => {
    const transportHeaders = "From: test@example.com\r\nDate: Fri, 15 Mar 2024 14:30:45 +0000\r\n";
    const msg = createMockMsg([{ tag: PidTagTransportMessageHeaders, value: transportHeaders }]);

    const result = parseMsgFromMsg(msg, dummyMsgToEml);
    assert.strictEqual(result.date.toISOString(), "2024-03-15T14:30:45.000Z");
  });

  it("should prefer PidTagClientSubmitTime over transport headers Date", () => {
    const submitDate = new Date("2024-03-15T14:29:00Z");
    const transportHeaders = "Date: Mon, 01 Jan 2024 09:00:00 +0000\r\n";
    const msg = createMockMsg([
      { tag: PidTagClientSubmitTime, value: submitDate },
      { tag: PidTagTransportMessageHeaders, value: transportHeaders },
    ]);

    const result = parseMsgFromMsg(msg, dummyMsgToEml);
    assert.strictEqual(result.date, submitDate);
  });

  it("should fall back to current date when no date source is available", () => {
    const before = new Date();
    const msg = createMockMsg([]);
    const result = parseMsgFromMsg(msg, dummyMsgToEml);
    const after = new Date();

    assert.ok(result.date.getTime() >= before.getTime());
    assert.ok(result.date.getTime() <= after.getTime());
  });

  it("should fall back to current date when transport headers have no Date", () => {
    const before = new Date();
    const transportHeaders = "From: test@example.com\r\nSubject: Test\r\n";
    const msg = createMockMsg([{ tag: PidTagTransportMessageHeaders, value: transportHeaders }]);
    const result = parseMsgFromMsg(msg, dummyMsgToEml);
    const after = new Date();

    assert.ok(result.date.getTime() >= before.getTime());
    assert.ok(result.date.getTime() <= after.getTime());
  });

  it("should fall back to current date when transport headers have an invalid Date", () => {
    const before = new Date();
    const transportHeaders = "Date: garbage-date-value\r\n";
    const msg = createMockMsg([{ tag: PidTagTransportMessageHeaders, value: transportHeaders }]);
    const result = parseMsgFromMsg(msg, dummyMsgToEml);
    const after = new Date();

    assert.ok(result.date.getTime() >= before.getTime());
    assert.ok(result.date.getTime() <= after.getTime());
  });
});
