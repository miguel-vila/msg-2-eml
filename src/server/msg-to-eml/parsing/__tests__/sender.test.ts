import assert from "node:assert";
import { describe, it } from "node:test";
import {
  type Msg,
  PidTagSenderEmailAddress,
  PidTagSenderName,
  PidTagSenderSmtpAddress,
  PidTagSentRepresentingEmailAddress,
  PidTagSentRepresentingName,
  PidTagSentRepresentingSmtpAddress,
  type PropertyTag,
} from "msg-parser";
import { extractSenderInfo, formatSender, parseFromTransportHeaders } from "../sender.js";

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
  } as unknown as Msg;
}

describe("formatSender", () => {
  it("should format with both name and email when different", () => {
    const result = formatSender("john@example.com", "John Doe");
    assert.strictEqual(result, '"John Doe" <john@example.com>');
  });

  it("should return just email when name equals email", () => {
    const result = formatSender("john@example.com", "john@example.com");
    assert.strictEqual(result, "john@example.com");
  });

  it("should return just email when name is undefined", () => {
    const result = formatSender("john@example.com", undefined);
    assert.strictEqual(result, "john@example.com");
  });

  it("should return just name when email is undefined", () => {
    const result = formatSender(undefined, "John Doe");
    assert.strictEqual(result, "John Doe");
  });

  it("should return fallback when both are undefined", () => {
    const result = formatSender(undefined, undefined);
    assert.strictEqual(result, "unknown@unknown.com");
  });

  it("should return just email when name is empty string", () => {
    const result = formatSender("john@example.com", "");
    assert.strictEqual(result, "john@example.com");
  });

  it("should return just name when email is empty string", () => {
    const result = formatSender("", "John Doe");
    assert.strictEqual(result, "John Doe");
  });

  it("should return fallback when both are empty strings", () => {
    const result = formatSender("", "");
    assert.strictEqual(result, "unknown@unknown.com");
  });

  it("should encode non-ASCII display name", () => {
    const result = formatSender("jose@example.com", "José García");
    assert.ok(result.includes("=?UTF-8?B?"), "Should encode non-ASCII name");
    assert.ok(result.includes("<jose@example.com>"), "Should include email in angle brackets");
  });

  it("should keep ASCII display names quoted", () => {
    const result = formatSender("john@example.com", "John Doe");
    assert.strictEqual(result, '"John Doe" <john@example.com>');
  });
});

describe("extractSenderInfo", () => {
  describe("normal scenarios (no delegation)", () => {
    it("should extract sender when only PidTagSender* properties are present", () => {
      const msg = createMockMsg([
        { tag: PidTagSenderSmtpAddress, value: "sender@example.com" },
        { tag: PidTagSenderName, value: "Sender Name" },
      ]);

      const result = extractSenderInfo(msg);

      assert.strictEqual(result.isOnBehalfOf, false);
      assert.strictEqual(result.from.email, "sender@example.com");
      assert.strictEqual(result.from.name, "Sender Name");
      assert.strictEqual(result.sender, undefined);
    });

    it("should extract sender when only PidTagSentRepresenting* properties are present", () => {
      const msg = createMockMsg([
        { tag: PidTagSentRepresentingSmtpAddress, value: "represented@example.com" },
        { tag: PidTagSentRepresentingName, value: "Represented Name" },
      ]);

      const result = extractSenderInfo(msg);

      assert.strictEqual(result.isOnBehalfOf, false);
      assert.strictEqual(result.from.email, "represented@example.com");
      assert.strictEqual(result.from.name, "Represented Name");
      assert.strictEqual(result.sender, undefined);
    });

    it("should prefer PidTagSender* when both are the same", () => {
      const msg = createMockMsg([
        { tag: PidTagSenderSmtpAddress, value: "same@example.com" },
        { tag: PidTagSenderName, value: "Same Person" },
        { tag: PidTagSentRepresentingSmtpAddress, value: "same@example.com" },
        { tag: PidTagSentRepresentingName, value: "Same Person" },
      ]);

      const result = extractSenderInfo(msg);

      assert.strictEqual(result.isOnBehalfOf, false);
      assert.strictEqual(result.from.email, "same@example.com");
      assert.strictEqual(result.from.name, "Same Person");
      assert.strictEqual(result.sender, undefined);
    });

    it("should handle case-insensitive email comparison", () => {
      const msg = createMockMsg([
        { tag: PidTagSenderSmtpAddress, value: "User@Example.COM" },
        { tag: PidTagSenderName, value: "User" },
        { tag: PidTagSentRepresentingSmtpAddress, value: "user@example.com" },
        { tag: PidTagSentRepresentingName, value: "User" },
      ]);

      const result = extractSenderInfo(msg);

      assert.strictEqual(result.isOnBehalfOf, false);
      assert.strictEqual(result.sender, undefined);
    });

    it("should prefer SMTP address over email address", () => {
      const msg = createMockMsg([
        { tag: PidTagSenderSmtpAddress, value: "smtp@example.com" },
        { tag: PidTagSenderEmailAddress, value: "/O=COMPANY/OU=EXCHANGE/CN=RECIPIENTS/CN=USER" },
        { tag: PidTagSenderName, value: "User" },
      ]);

      const result = extractSenderInfo(msg);

      assert.strictEqual(result.from.email, "smtp@example.com");
    });
  });

  describe("on behalf of scenarios", () => {
    it("should detect on behalf of when sender and represented emails differ", () => {
      const msg = createMockMsg([
        { tag: PidTagSenderSmtpAddress, value: "assistant@example.com" },
        { tag: PidTagSenderName, value: "Assistant" },
        { tag: PidTagSentRepresentingSmtpAddress, value: "boss@example.com" },
        { tag: PidTagSentRepresentingName, value: "Boss" },
      ]);

      const result = extractSenderInfo(msg);

      assert.strictEqual(result.isOnBehalfOf, true);
      assert.strictEqual(result.from.email, "boss@example.com");
      assert.strictEqual(result.from.name, "Boss");
      assert.notStrictEqual(result.sender, undefined);
      assert.strictEqual(result.sender?.email, "assistant@example.com");
      assert.strictEqual(result.sender?.name, "Assistant");
    });

    it("should detect on behalf of using EmailAddress properties when SMTP is not available", () => {
      const msg = createMockMsg([
        { tag: PidTagSenderEmailAddress, value: "assistant@example.com" },
        { tag: PidTagSenderName, value: "Assistant" },
        { tag: PidTagSentRepresentingEmailAddress, value: "boss@example.com" },
        { tag: PidTagSentRepresentingName, value: "Boss" },
      ]);

      const result = extractSenderInfo(msg);

      assert.strictEqual(result.isOnBehalfOf, true);
      assert.strictEqual(result.from.email, "boss@example.com");
      assert.strictEqual(result.sender?.email, "assistant@example.com");
    });

    it("should handle on behalf of with only names (no emails)", () => {
      const msg = createMockMsg([
        { tag: PidTagSenderName, value: "Assistant" },
        { tag: PidTagSentRepresentingName, value: "Boss" },
      ]);

      const result = extractSenderInfo(msg);

      assert.strictEqual(result.isOnBehalfOf, true);
      assert.strictEqual(result.from.name, "Boss");
      assert.strictEqual(result.from.email, undefined);
      assert.strictEqual(result.sender?.name, "Assistant");
      assert.strictEqual(result.sender?.email, undefined);
    });

    it("should detect on behalf of when sender has email but represented only has name", () => {
      const msg = createMockMsg([
        { tag: PidTagSenderSmtpAddress, value: "assistant@example.com" },
        { tag: PidTagSenderName, value: "Assistant" },
        { tag: PidTagSentRepresentingName, value: "Boss (no email)" },
      ]);

      const result = extractSenderInfo(msg);

      assert.strictEqual(result.isOnBehalfOf, true);
      assert.strictEqual(result.from.name, "Boss (no email)");
      assert.strictEqual(result.sender?.email, "assistant@example.com");
    });

    it("should handle shared mailbox scenario with different emails", () => {
      const msg = createMockMsg([
        { tag: PidTagSenderSmtpAddress, value: "john.doe@example.com" },
        { tag: PidTagSenderName, value: "John Doe" },
        { tag: PidTagSentRepresentingSmtpAddress, value: "support@example.com" },
        { tag: PidTagSentRepresentingName, value: "Support Team" },
      ]);

      const result = extractSenderInfo(msg);

      assert.strictEqual(result.isOnBehalfOf, true);
      assert.strictEqual(result.from.email, "support@example.com");
      assert.strictEqual(result.from.name, "Support Team");
      assert.strictEqual(result.sender?.email, "john.doe@example.com");
      assert.strictEqual(result.sender?.name, "John Doe");
    });

    it("should handle delegate scenario with different names but partial email", () => {
      const msg = createMockMsg([
        { tag: PidTagSenderSmtpAddress, value: "secretary@example.com" },
        { tag: PidTagSenderName, value: "Secretary Smith" },
        { tag: PidTagSentRepresentingSmtpAddress, value: "executive@example.com" },
        { tag: PidTagSentRepresentingName, value: "CEO Johnson" },
      ]);

      const result = extractSenderInfo(msg);

      assert.strictEqual(result.isOnBehalfOf, true);
      // From should be the represented person (CEO)
      assert.strictEqual(result.from.email, "executive@example.com");
      assert.strictEqual(result.from.name, "CEO Johnson");
      // Sender should be the actual sender (Secretary)
      assert.strictEqual(result.sender?.email, "secretary@example.com");
      assert.strictEqual(result.sender?.name, "Secretary Smith");
    });
  });

  describe("edge cases", () => {
    it("should handle empty MSG with no sender properties", () => {
      const msg = createMockMsg([]);

      const result = extractSenderInfo(msg);

      assert.strictEqual(result.isOnBehalfOf, false);
      assert.strictEqual(result.from.email, undefined);
      assert.strictEqual(result.from.name, undefined);
      assert.strictEqual(result.sender, undefined);
    });

    it("should handle whitespace in email addresses", () => {
      const msg = createMockMsg([
        { tag: PidTagSenderSmtpAddress, value: "  sender@example.com  " },
        { tag: PidTagSentRepresentingSmtpAddress, value: "sender@example.com" },
      ]);

      const result = extractSenderInfo(msg);

      // Should normalize and recognize as same email
      assert.strictEqual(result.isOnBehalfOf, false);
    });

    it("should not detect on behalf of when only actual sender is present", () => {
      const msg = createMockMsg([
        { tag: PidTagSenderSmtpAddress, value: "sender@example.com" },
        { tag: PidTagSenderName, value: "Sender" },
        // No PidTagSentRepresenting* properties
      ]);

      const result = extractSenderInfo(msg);

      assert.strictEqual(result.isOnBehalfOf, false);
      assert.strictEqual(result.sender, undefined);
    });

    it("should not detect on behalf of when only represented sender is present", () => {
      const msg = createMockMsg([
        // No PidTagSender* properties
        { tag: PidTagSentRepresentingSmtpAddress, value: "represented@example.com" },
        { tag: PidTagSentRepresentingName, value: "Represented" },
      ]);

      const result = extractSenderInfo(msg);

      assert.strictEqual(result.isOnBehalfOf, false);
      assert.strictEqual(result.sender, undefined);
    });
  });

  describe("transport header fallback", () => {
    it("should fall back to From: in transport headers when no MAPI sender properties exist", () => {
      const msg = createMockMsg([]);
      const transportHeaders = "From: fallback@example.com\r\nTo: recipient@example.com\r\n";

      const result = extractSenderInfo(msg, transportHeaders);

      assert.strictEqual(result.isOnBehalfOf, false);
      assert.strictEqual(result.from.email, "fallback@example.com");
      assert.strictEqual(result.from.name, undefined);
      assert.strictEqual(result.sender, undefined);
    });

    it("should extract both name and email from transport headers", () => {
      const msg = createMockMsg([]);
      const transportHeaders = 'From: "John Doe" <john@example.com>\r\nTo: recipient@example.com\r\n';

      const result = extractSenderInfo(msg, transportHeaders);

      assert.strictEqual(result.from.email, "john@example.com");
      assert.strictEqual(result.from.name, "John Doe");
    });

    it("should not use transport headers when MAPI sender properties exist", () => {
      const msg = createMockMsg([
        { tag: PidTagSenderSmtpAddress, value: "mapi@example.com" },
        { tag: PidTagSenderName, value: "MAPI Sender" },
      ]);
      const transportHeaders = 'From: "Transport Sender" <transport@example.com>\r\n';

      const result = extractSenderInfo(msg, transportHeaders);

      assert.strictEqual(result.from.email, "mapi@example.com");
      assert.strictEqual(result.from.name, "MAPI Sender");
    });

    it("should not use transport headers when represented sender properties exist", () => {
      const msg = createMockMsg([{ tag: PidTagSentRepresentingSmtpAddress, value: "represented@example.com" }]);
      const transportHeaders = 'From: "Transport Sender" <transport@example.com>\r\n';

      const result = extractSenderInfo(msg, transportHeaders);

      assert.strictEqual(result.from.email, "represented@example.com");
    });

    it("should handle missing transport headers gracefully", () => {
      const msg = createMockMsg([]);

      const result = extractSenderInfo(msg, undefined);

      assert.strictEqual(result.from.email, undefined);
      assert.strictEqual(result.from.name, undefined);
    });

    it("should handle transport headers without From: header", () => {
      const msg = createMockMsg([]);
      const transportHeaders = "To: recipient@example.com\r\nSubject: Test\r\n";

      const result = extractSenderInfo(msg, transportHeaders);

      assert.strictEqual(result.from.email, undefined);
      assert.strictEqual(result.from.name, undefined);
    });
  });
});

describe("parseFromTransportHeaders", () => {
  it("should parse bare email address", () => {
    const result = parseFromTransportHeaders("From: user@example.com\r\n");

    assert.notStrictEqual(result, undefined);
    assert.strictEqual(result?.email, "user@example.com");
    assert.strictEqual(result?.name, undefined);
  });

  it("should parse quoted display name with angle-bracket email", () => {
    const result = parseFromTransportHeaders('From: "John Doe" <john@example.com>\r\n');

    assert.notStrictEqual(result, undefined);
    assert.strictEqual(result?.email, "john@example.com");
    assert.strictEqual(result?.name, "John Doe");
  });

  it("should parse unquoted display name with angle-bracket email", () => {
    const result = parseFromTransportHeaders("From: John Doe <john@example.com>\r\n");

    assert.notStrictEqual(result, undefined);
    assert.strictEqual(result?.email, "john@example.com");
    assert.strictEqual(result?.name, "John Doe");
  });

  it("should parse angle-bracket email without display name", () => {
    const result = parseFromTransportHeaders("From: <john@example.com>\r\n");

    assert.notStrictEqual(result, undefined);
    assert.strictEqual(result?.email, "john@example.com");
    assert.strictEqual(result?.name, undefined);
  });

  it("should handle folded From: header", () => {
    const headers = "From: Very Long Display Name\r\n <user@example.com>\r\nTo: other@example.com\r\n";

    const result = parseFromTransportHeaders(headers);

    assert.notStrictEqual(result, undefined);
    assert.strictEqual(result?.email, "user@example.com");
    assert.strictEqual(result?.name, "Very Long Display Name");
  });

  it("should return undefined when From: header is missing", () => {
    const result = parseFromTransportHeaders("To: user@example.com\r\nSubject: Test\r\n");

    assert.strictEqual(result, undefined);
  });

  it("should return undefined for empty string", () => {
    const result = parseFromTransportHeaders("");

    assert.strictEqual(result, undefined);
  });

  it("should handle case-insensitive From: header", () => {
    const result = parseFromTransportHeaders("FROM: user@example.com\r\n");

    assert.notStrictEqual(result, undefined);
    assert.strictEqual(result?.email, "user@example.com");
  });

  it("should handle From: header among many headers", () => {
    const headers = [
      "Received: from mail.example.com",
      "Date: Mon, 1 Jan 2024 12:00:00 +0000",
      'From: "Sender Name" <sender@example.com>',
      "To: recipient@example.com",
      "Subject: Test",
      "",
    ].join("\r\n");

    const result = parseFromTransportHeaders(headers);

    assert.notStrictEqual(result, undefined);
    assert.strictEqual(result?.email, "sender@example.com");
    assert.strictEqual(result?.name, "Sender Name");
  });

  it("should handle LF line endings", () => {
    const headers = "From: user@example.com\nTo: other@example.com\n";

    const result = parseFromTransportHeaders(headers);

    assert.notStrictEqual(result, undefined);
    assert.strictEqual(result?.email, "user@example.com");
  });

  it("should handle From: header with empty value", () => {
    const result = parseFromTransportHeaders("From: \r\nTo: other@example.com\r\n");

    assert.strictEqual(result, undefined);
  });
});
