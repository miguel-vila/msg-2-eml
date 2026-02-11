import { describe, it } from "node:test";
import assert from "node:assert";
import { convertToEml, formatSender } from "./msg-to-eml.js";

describe("convertToEml", () => {
  it("should generate basic EML headers", () => {
    const parsed = {
      subject: "Test Subject",
      from: "sender@example.com",
      recipients: [{ name: "Recipient", email: "recipient@example.com", type: "to" as const }],
      date: new Date("2024-01-15T10:30:00Z"),
      body: "Hello World",
      attachments: [],
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("From: sender@example.com"));
    assert.ok(eml.includes('To: "Recipient" <recipient@example.com>'));
    assert.ok(eml.includes("Subject: Test Subject"));
    assert.ok(eml.includes("MIME-Version: 1.0"));
    assert.ok(eml.includes("Hello World"));
  });

  it("should handle CC and BCC recipients", () => {
    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [
        { name: "", email: "to@example.com", type: "to" as const },
        { name: "CC Person", email: "cc@example.com", type: "cc" as const },
        { name: "", email: "bcc@example.com", type: "bcc" as const },
      ],
      date: new Date(),
      body: "Test body",
      attachments: [],
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("To: to@example.com"));
    assert.ok(eml.includes('Cc: "CC Person" <cc@example.com>'));
    assert.ok(eml.includes("Bcc: bcc@example.com"));
  });

  it("should generate multipart/alternative for HTML content", () => {
    const parsed = {
      subject: "HTML Email",
      from: "sender@example.com",
      recipients: [{ name: "", email: "recipient@example.com", type: "to" as const }],
      date: new Date(),
      body: "Plain text version",
      bodyHtml: "<html><body><p>HTML version</p></body></html>",
      attachments: [],
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("multipart/alternative"));
    assert.ok(eml.includes("text/plain"));
    assert.ok(eml.includes("text/html"));
    assert.ok(eml.includes("Plain text version"));
  });

  it("should encode attachments in base64", () => {
    const parsed = {
      subject: "With Attachment",
      from: "sender@example.com",
      recipients: [{ name: "", email: "recipient@example.com", type: "to" as const }],
      date: new Date(),
      body: "See attached",
      attachments: [
        {
          fileName: "test.txt",
          content: new Uint8Array([72, 101, 108, 108, 111]), // "Hello"
          contentType: "text/plain",
        },
      ],
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("multipart/mixed"));
    assert.ok(eml.includes('filename="test.txt"'));
    assert.ok(eml.includes("Content-Transfer-Encoding: base64"));
    assert.ok(eml.includes("SGVsbG8=")); // "Hello" in base64
  });

  it("should format dates correctly", () => {
    const parsed = {
      subject: "Date Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date("2024-03-15T14:30:45Z"),
      body: "Test",
      attachments: [],
    };

    const eml = convertToEml(parsed);

    // Should be in RFC 822 format
    assert.ok(eml.includes("Date: Fri, 15 Mar 2024 14:30:45 +0000"));
  });

  it("should handle special characters with quoted-printable encoding", () => {
    const parsed = {
      subject: "Special chars",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Café résumé naïve",
      attachments: [],
    };

    const eml = convertToEml(parsed);

    // Quoted-printable encodes non-ASCII characters
    assert.ok(eml.includes("Content-Transfer-Encoding: quoted-printable"));
    assert.ok(!eml.includes("Café")); // Should be encoded
  });
});

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
});
