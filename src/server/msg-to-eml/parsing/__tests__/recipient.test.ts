import assert from "node:assert";
import { describe, it } from "node:test";
import {
  type Recipient as MsgRecipient,
  PidTagDisplayName,
  PidTagEmailAddress,
  PidTagRecipientType,
  PidTagSmtpAddress,
  type PropertyTag,
} from "msg-parser";
import { parseRecipient, parseRecipientsFromTransportHeaders, type TransportRecipients } from "../recipient.js";

// Helper to create a mock Recipient with specific property values
interface MockPropertyValue {
  tag: PropertyTag;
  value: unknown;
}

function createMockRecipient(properties: MockPropertyValue[]): MsgRecipient {
  return {
    getProperty: <T>(propertyId: PropertyTag): T | undefined => {
      const found = properties.find((p) => p.tag === propertyId);
      return found?.value as T | undefined;
    },
  } as unknown as MsgRecipient;
}

describe("parseRecipientsFromTransportHeaders", () => {
  it("should parse a simple To: header with one address", () => {
    const headers = "To: user@example.com\r\n";
    const result = parseRecipientsFromTransportHeaders(headers);

    assert.strictEqual(result.to.length, 1);
    assert.strictEqual(result.to[0].email, "user@example.com");
    assert.strictEqual(result.to[0].name, undefined);
    assert.strictEqual(result.cc.length, 0);
    assert.strictEqual(result.bcc.length, 0);
  });

  it("should parse To: header with display name and email", () => {
    const headers = 'To: "John Doe" <john@example.com>\r\n';
    const result = parseRecipientsFromTransportHeaders(headers);

    assert.strictEqual(result.to.length, 1);
    assert.strictEqual(result.to[0].email, "john@example.com");
    assert.strictEqual(result.to[0].name, "John Doe");
  });

  it("should parse unquoted display name with email", () => {
    const headers = "To: John Doe <john@example.com>\r\n";
    const result = parseRecipientsFromTransportHeaders(headers);

    assert.strictEqual(result.to.length, 1);
    assert.strictEqual(result.to[0].email, "john@example.com");
    assert.strictEqual(result.to[0].name, "John Doe");
  });

  it("should parse multiple To: addresses", () => {
    const headers = 'To: "Alice" <alice@example.com>, "Bob" <bob@example.com>\r\n';
    const result = parseRecipientsFromTransportHeaders(headers);

    assert.strictEqual(result.to.length, 2);
    assert.strictEqual(result.to[0].email, "alice@example.com");
    assert.strictEqual(result.to[0].name, "Alice");
    assert.strictEqual(result.to[1].email, "bob@example.com");
    assert.strictEqual(result.to[1].name, "Bob");
  });

  it("should parse To: and Cc: headers", () => {
    const headers = ['To: "To User" <to@example.com>', 'Cc: "Cc User" <cc@example.com>', ""].join("\r\n");
    const result = parseRecipientsFromTransportHeaders(headers);

    assert.strictEqual(result.to.length, 1);
    assert.strictEqual(result.to[0].email, "to@example.com");
    assert.strictEqual(result.cc.length, 1);
    assert.strictEqual(result.cc[0].email, "cc@example.com");
  });

  it("should parse Bcc: header", () => {
    const headers = 'Bcc: "Secret" <bcc@example.com>\r\n';
    const result = parseRecipientsFromTransportHeaders(headers);

    assert.strictEqual(result.bcc.length, 1);
    assert.strictEqual(result.bcc[0].email, "bcc@example.com");
    assert.strictEqual(result.bcc[0].name, "Secret");
  });

  it("should handle folded headers", () => {
    const headers = ["To: Very Long Name", " <user@example.com>", "Subject: Test", ""].join("\r\n");
    const result = parseRecipientsFromTransportHeaders(headers);

    assert.strictEqual(result.to.length, 1);
    assert.strictEqual(result.to[0].email, "user@example.com");
    assert.strictEqual(result.to[0].name, "Very Long Name");
  });

  it("should handle case-insensitive header names", () => {
    const headers = "TO: user@example.com\r\nCC: other@example.com\r\n";
    const result = parseRecipientsFromTransportHeaders(headers);

    assert.strictEqual(result.to.length, 1);
    assert.strictEqual(result.to[0].email, "user@example.com");
    assert.strictEqual(result.cc.length, 1);
    assert.strictEqual(result.cc[0].email, "other@example.com");
  });

  it("should handle angle brackets without display name", () => {
    const headers = "To: <user@example.com>\r\n";
    const result = parseRecipientsFromTransportHeaders(headers);

    assert.strictEqual(result.to.length, 1);
    assert.strictEqual(result.to[0].email, "user@example.com");
    assert.strictEqual(result.to[0].name, undefined);
  });

  it("should handle empty transport headers", () => {
    const result = parseRecipientsFromTransportHeaders("");

    assert.strictEqual(result.to.length, 0);
    assert.strictEqual(result.cc.length, 0);
    assert.strictEqual(result.bcc.length, 0);
  });

  it("should handle transport headers with no To/Cc/Bcc", () => {
    const headers = "From: sender@example.com\r\nSubject: Test\r\n";
    const result = parseRecipientsFromTransportHeaders(headers);

    assert.strictEqual(result.to.length, 0);
    assert.strictEqual(result.cc.length, 0);
    assert.strictEqual(result.bcc.length, 0);
  });

  it("should handle multiple addresses with mixed formats", () => {
    const headers = 'To: "Alice Smith" <alice@example.com>, bob@example.com, Charlie <charlie@example.com>\r\n';
    const result = parseRecipientsFromTransportHeaders(headers);

    assert.strictEqual(result.to.length, 3);
    assert.strictEqual(result.to[0].email, "alice@example.com");
    assert.strictEqual(result.to[0].name, "Alice Smith");
    assert.strictEqual(result.to[1].email, "bob@example.com");
    assert.strictEqual(result.to[1].name, undefined);
    assert.strictEqual(result.to[2].email, "charlie@example.com");
    assert.strictEqual(result.to[2].name, "Charlie");
  });

  it("should handle LF line endings", () => {
    const headers = "To: user@example.com\nCc: other@example.com\n";
    const result = parseRecipientsFromTransportHeaders(headers);

    assert.strictEqual(result.to.length, 1);
    assert.strictEqual(result.cc.length, 1);
  });

  it("should handle headers among many other headers", () => {
    const headers = [
      "Received: from mail.example.com",
      "Date: Mon, 1 Jan 2024 12:00:00 +0000",
      'From: "Sender" <sender@example.com>',
      'To: "Recipient One" <recip1@example.com>,',
      ' "Recipient Two" <recip2@example.com>',
      "Cc: cc@example.com",
      "Subject: Test",
      "",
    ].join("\r\n");
    const result = parseRecipientsFromTransportHeaders(headers);

    assert.strictEqual(result.to.length, 2);
    assert.strictEqual(result.to[0].email, "recip1@example.com");
    assert.strictEqual(result.to[0].name, "Recipient One");
    assert.strictEqual(result.to[1].email, "recip2@example.com");
    assert.strictEqual(result.to[1].name, "Recipient Two");
    assert.strictEqual(result.cc.length, 1);
    assert.strictEqual(result.cc[0].email, "cc@example.com");
  });
});

describe("parseRecipient with transport header fallback", () => {
  it("should use SMTP address when available (ignores transport headers)", () => {
    const recipient = createMockRecipient([
      { tag: PidTagDisplayName, value: "John Doe" },
      { tag: PidTagSmtpAddress, value: "john@example.com" },
      { tag: PidTagRecipientType, value: 1 },
    ]);
    const transportRecipients: TransportRecipients = {
      to: [{ name: "John Doe", email: "different@example.com" }],
      cc: [],
      bcc: [],
    };

    const result = parseRecipient(recipient, transportRecipients);

    assert.strictEqual(result.email, "john@example.com");
    assert.strictEqual(result.name, "John Doe");
    assert.strictEqual(result.type, "to");
  });

  it("should use EmailAddress when SMTP is missing (ignores transport headers)", () => {
    const recipient = createMockRecipient([
      { tag: PidTagDisplayName, value: "John Doe" },
      { tag: PidTagEmailAddress, value: "john@exchange.local" },
      { tag: PidTagRecipientType, value: 1 },
    ]);
    const transportRecipients: TransportRecipients = {
      to: [{ name: "John Doe", email: "john@example.com" }],
      cc: [],
      bcc: [],
    };

    const result = parseRecipient(recipient, transportRecipients);

    assert.strictEqual(result.email, "john@exchange.local");
  });

  it("should resolve email from transport headers when MAPI properties are missing", () => {
    const recipient = createMockRecipient([
      { tag: PidTagDisplayName, value: "John Doe" },
      { tag: PidTagRecipientType, value: 1 },
    ]);
    const transportRecipients: TransportRecipients = {
      to: [{ name: "John Doe", email: "john@example.com" }],
      cc: [],
      bcc: [],
    };

    const result = parseRecipient(recipient, transportRecipients);

    assert.strictEqual(result.email, "john@example.com");
    assert.strictEqual(result.name, "John Doe");
    assert.strictEqual(result.type, "to");
  });

  it("should resolve Cc recipient email from transport headers", () => {
    const recipient = createMockRecipient([
      { tag: PidTagDisplayName, value: "Jane Smith" },
      { tag: PidTagRecipientType, value: 2 },
    ]);
    const transportRecipients: TransportRecipients = {
      to: [{ name: "Other Person", email: "other@example.com" }],
      cc: [{ name: "Jane Smith", email: "jane@example.com" }],
      bcc: [],
    };

    const result = parseRecipient(recipient, transportRecipients);

    assert.strictEqual(result.email, "jane@example.com");
    assert.strictEqual(result.type, "cc");
  });

  it("should resolve Bcc recipient email from transport headers", () => {
    const recipient = createMockRecipient([
      { tag: PidTagDisplayName, value: "Secret Recipient" },
      { tag: PidTagRecipientType, value: 3 },
    ]);
    const transportRecipients: TransportRecipients = {
      to: [],
      cc: [],
      bcc: [{ name: "Secret Recipient", email: "secret@example.com" }],
    };

    const result = parseRecipient(recipient, transportRecipients);

    assert.strictEqual(result.email, "secret@example.com");
    assert.strictEqual(result.type, "bcc");
  });

  it("should match display name case-insensitively", () => {
    const recipient = createMockRecipient([
      { tag: PidTagDisplayName, value: "JOHN DOE" },
      { tag: PidTagRecipientType, value: 1 },
    ]);
    const transportRecipients: TransportRecipients = {
      to: [{ name: "John Doe", email: "john@example.com" }],
      cc: [],
      bcc: [],
    };

    const result = parseRecipient(recipient, transportRecipients);

    assert.strictEqual(result.email, "john@example.com");
  });

  it("should fall back to display name when no match in transport headers", () => {
    const recipient = createMockRecipient([
      { tag: PidTagDisplayName, value: "Unknown User" },
      { tag: PidTagRecipientType, value: 1 },
    ]);
    const transportRecipients: TransportRecipients = {
      to: [{ name: "John Doe", email: "john@example.com" }],
      cc: [],
      bcc: [],
    };

    const result = parseRecipient(recipient, transportRecipients);

    assert.strictEqual(result.email, "Unknown User");
  });

  it("should fall back to display name when no transport recipients provided", () => {
    const recipient = createMockRecipient([
      { tag: PidTagDisplayName, value: "John Doe" },
      { tag: PidTagRecipientType, value: 1 },
    ]);

    const result = parseRecipient(recipient);

    assert.strictEqual(result.email, "John Doe");
    assert.strictEqual(result.name, "John Doe");
  });

  it("should fall back to name when display name is empty", () => {
    const recipient = createMockRecipient([{ tag: PidTagRecipientType, value: 1 }]);

    const result = parseRecipient(recipient);

    assert.strictEqual(result.email, "");
    assert.strictEqual(result.name, "");
  });

  it("should try other recipient types when no match in same type", () => {
    // Recipient is marked as To: but matching address is in Cc: header
    const recipient = createMockRecipient([
      { tag: PidTagDisplayName, value: "Cross Type" },
      { tag: PidTagRecipientType, value: 1 },
    ]);
    const transportRecipients: TransportRecipients = {
      to: [],
      cc: [{ name: "Cross Type", email: "cross@example.com" }],
      bcc: [],
    };

    const result = parseRecipient(recipient, transportRecipients);

    assert.strictEqual(result.email, "cross@example.com");
  });

  it("should handle transport address with name but no email", () => {
    const recipient = createMockRecipient([
      { tag: PidTagDisplayName, value: "No Email Person" },
      { tag: PidTagRecipientType, value: 1 },
    ]);
    const transportRecipients: TransportRecipients = {
      to: [{ name: "No Email Person", email: undefined }],
      cc: [],
      bcc: [],
    };

    const result = parseRecipient(recipient, transportRecipients);

    // Should fall back to display name since transport has no email either
    assert.strictEqual(result.email, "No Email Person");
  });
});
