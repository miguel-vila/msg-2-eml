import { describe, it } from "node:test";
import assert from "node:assert";
import { convertToEml, formatSender, mapToXPriority, foldHeader, extractBodyFromRtf } from "./msg-to-eml.js";

/**
 * Creates an uncompressed RTF format for PidTagRtfCompressed.
 * The format uses MELA (uncompressed) magic number.
 * Structure: fileSize (4) + rawSize (4) + compType (4) + crc (4) + raw data
 */
function createUncompressedRtf(rtfContent: string): number[] {
  const rawBytes = Buffer.from(rtfContent, "latin1");
  const rawSize = rawBytes.length;
  const fileSize = rawSize + 12; // rawSize + compType + crc (fileSize excludes itself)
  const UNCOMPRESSED = 0x414C454D; // "MELA" in little-endian
  const crc = 0; // CRC is not checked for uncompressed

  const result: number[] = [];

  // Write fileSize (little-endian 32-bit)
  result.push(fileSize & 0xff);
  result.push((fileSize >> 8) & 0xff);
  result.push((fileSize >> 16) & 0xff);
  result.push((fileSize >> 24) & 0xff);

  // Write rawSize (little-endian 32-bit)
  result.push(rawSize & 0xff);
  result.push((rawSize >> 8) & 0xff);
  result.push((rawSize >> 16) & 0xff);
  result.push((rawSize >> 24) & 0xff);

  // Write compType (little-endian 32-bit)
  result.push(UNCOMPRESSED & 0xff);
  result.push((UNCOMPRESSED >> 8) & 0xff);
  result.push((UNCOMPRESSED >> 16) & 0xff);
  result.push((UNCOMPRESSED >> 24) & 0xff);

  // Write crc (little-endian 32-bit)
  result.push(crc & 0xff);
  result.push((crc >> 8) & 0xff);
  result.push((crc >> 16) & 0xff);
  result.push((crc >> 24) & 0xff);

  // Write raw RTF content
  for (let i = 0; i < rawBytes.length; i++) {
    result.push(rawBytes[i]);
  }

  return result;
}

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

  it("should fold long To headers with many recipients", () => {
    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [
        { name: "Alice Smith", email: "alice@example.com", type: "to" as const },
        { name: "Bob Johnson", email: "bob@example.com", type: "to" as const },
        { name: "Charlie Brown", email: "charlie@example.com", type: "to" as const },
        { name: "Diana Prince", email: "diana@example.com", type: "to" as const },
      ],
      date: new Date(),
      body: "Test body",
      attachments: [],
    };

    const eml = convertToEml(parsed);

    // Should contain folded To header (CRLF followed by tab)
    const toHeaderMatch = eml.match(/^To:.*$/m);
    assert.ok(toHeaderMatch, "Should have To header");

    // Find the full folded header
    const headerSection = eml.split("\r\n\r\n")[0];
    assert.ok(headerSection.includes("\r\n\t"), "Should contain folded continuation lines");

    // All recipients should be present
    assert.ok(eml.includes("<alice@example.com>"));
    assert.ok(eml.includes("<bob@example.com>"));
    assert.ok(eml.includes("<charlie@example.com>"));
    assert.ok(eml.includes("<diana@example.com>"));
  });

  it("should fold long Subject headers", () => {
    const longSubject = "This is a very very very long subject line that will definitely exceed the 78 character limit according to RFC 5322";
    const parsed = {
      subject: longSubject,
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
    };

    const eml = convertToEml(parsed);

    // The header section should contain folding
    const headerSection = eml.split("\r\n\r\n")[0];
    assert.ok(headerSection.includes("\r\n\t"), "Long subject should be folded");

    // Original content should still be present (just split across lines)
    const normalizedEml = eml.replace(/\r\n[\t ]/g, " ");
    assert.ok(normalizedEml.includes(longSubject), "Full subject content should be preserved");
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

describe("mapToXPriority", () => {
  it("should map PidTagPriority urgent (1) to X-Priority 1", () => {
    assert.strictEqual(mapToXPriority(1, undefined), 1);
  });

  it("should map PidTagPriority normal (0) to X-Priority 3", () => {
    assert.strictEqual(mapToXPriority(0, undefined), 3);
  });

  it("should map PidTagPriority non-urgent (-1) to X-Priority 5", () => {
    assert.strictEqual(mapToXPriority(-1, undefined), 5);
  });

  it("should map PidTagImportance high (2) to X-Priority 1", () => {
    assert.strictEqual(mapToXPriority(undefined, 2), 1);
  });

  it("should map PidTagImportance normal (1) to X-Priority 3", () => {
    assert.strictEqual(mapToXPriority(undefined, 1), 3);
  });

  it("should map PidTagImportance low (0) to X-Priority 5", () => {
    assert.strictEqual(mapToXPriority(undefined, 0), 5);
  });

  it("should prefer PidTagPriority over PidTagImportance", () => {
    // Priority urgent with importance low should still return 1
    assert.strictEqual(mapToXPriority(1, 0), 1);
  });

  it("should return undefined when both are undefined", () => {
    assert.strictEqual(mapToXPriority(undefined, undefined), undefined);
  });
});

describe("foldHeader", () => {
  it("should not fold headers shorter than 78 characters", () => {
    const result = foldHeader("Subject", "Short subject line");
    assert.strictEqual(result, "Subject: Short subject line");
    assert.ok(!result.includes("\r\n"));
  });

  it("should fold headers longer than 78 characters", () => {
    const longSubject = "This is a very long subject line that exceeds the 78 character limit and needs to be folded properly";
    const result = foldHeader("Subject", longSubject);

    // Should contain CRLF followed by tab (continuation)
    assert.ok(result.includes("\r\n\t"), "Should contain CRLF+TAB for continuation");

    // Each line should be <= 78 characters
    const lines = result.split("\r\n");
    for (const line of lines) {
      assert.ok(line.length <= 78, `Line too long: ${line.length} chars - "${line}"`);
    }
  });

  it("should not break inside encoded words", () => {
    const encodedWord = "=?UTF-8?B?VGhpcyBpcyBhIHZlcnkgbG9uZyBlbmNvZGVkIHdvcmQgdGhhdCBzaG91bGQgbm90IGJlIGJyb2tlbg==?=";
    const subject = `Prefix ${encodedWord} suffix text`;
    const result = foldHeader("Subject", subject);

    // The encoded word should remain intact
    assert.ok(result.includes(encodedWord), "Encoded word should not be broken");
  });

  it("should not break inside email addresses in angle brackets", () => {
    const emailAddress = "<verylongemailaddress@verylongdomainname.example.com>";
    const result = foldHeader("From", `"Very Long Display Name That Makes The Header Long" ${emailAddress}`);

    // The email address in angle brackets should remain intact
    assert.ok(result.includes(emailAddress), "Email address should not be broken");
  });

  it("should not break inside quoted strings", () => {
    const quotedString = '"This is a very long quoted display name"';
    const result = foldHeader("To", `${quotedString} <user@example.com>`);

    // The quoted string should remain intact
    assert.ok(result.includes(quotedString), "Quoted string should not be broken");
  });

  it("should fold recipient lists at appropriate points", () => {
    const recipients = '"Alice Smith" <alice@example.com>, "Bob Johnson" <bob@example.com>, "Charlie Brown" <charlie@example.com>';
    const result = foldHeader("To", recipients);

    // Should be folded
    assert.ok(result.includes("\r\n"), "Long recipient list should be folded");

    // Each email should remain intact
    assert.ok(result.includes("<alice@example.com>"));
    assert.ok(result.includes("<bob@example.com>"));
    assert.ok(result.includes("<charlie@example.com>"));
  });

  it("should respect custom max line length", () => {
    // "Subject: " is 9 chars, so we need content that makes total > 40
    const subject = "A moderately long subject line that exceeds forty characters";
    const result = foldHeader("Subject", subject, 40);

    // Should be folded with a shorter line length
    assert.ok(result.includes("\r\n"), "Should be folded with custom max length");

    const lines = result.split("\r\n");
    for (const line of lines) {
      assert.ok(line.length <= 40, `Line too long for custom limit: ${line.length} chars`);
    }
  });

  it("should fold References header with multiple message IDs", () => {
    const references = "<msg1@example.com> <msg2@example.com> <msg3@example.com> <msg4@example.com>";
    const result = foldHeader("References", references);

    // Each message ID should remain intact
    assert.ok(result.includes("<msg1@example.com>"));
    assert.ok(result.includes("<msg2@example.com>"));
    assert.ok(result.includes("<msg3@example.com>"));
    assert.ok(result.includes("<msg4@example.com>"));
  });

  it("should use tab for continuation lines", () => {
    const longSubject = "This is a very long subject line that exceeds the 78 character limit and needs proper folding";
    const result = foldHeader("Subject", longSubject);

    // Continuation lines should start with tab
    const lines = result.split("\r\n");
    for (let i = 1; i < lines.length; i++) {
      assert.ok(lines[i].startsWith("\t"), `Continuation line ${i} should start with tab`);
    }
  });
});

describe("convertToEml with message headers", () => {
  it("should include Message-ID header when present", () => {
    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        messageId: "<abc123@example.com>",
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("Message-ID: <abc123@example.com>"));
  });

  it("should include In-Reply-To header when present", () => {
    const parsed = {
      subject: "Re: Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        inReplyTo: "<original123@example.com>",
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("In-Reply-To: <original123@example.com>"));
  });

  it("should include References header when present", () => {
    const parsed = {
      subject: "Re: Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        references: "<msg1@example.com> <msg2@example.com>",
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("References: <msg1@example.com> <msg2@example.com>"));
  });

  it("should include Reply-To header when present", () => {
    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        replyTo: "reply@example.com",
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("Reply-To: reply@example.com"));
  });

  it("should include X-Priority header when present", () => {
    const parsed = {
      subject: "Urgent Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        priority: 1,
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("X-Priority: 1"));
  });

  it("should include all headers when present", () => {
    const parsed = {
      subject: "Full Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        messageId: "<full123@example.com>",
        inReplyTo: "<original@example.com>",
        references: "<ref1@example.com>",
        replyTo: "reply@example.com",
        priority: 2,
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("Message-ID: <full123@example.com>"));
    assert.ok(eml.includes("In-Reply-To: <original@example.com>"));
    assert.ok(eml.includes("References: <ref1@example.com>"));
    assert.ok(eml.includes("Reply-To: reply@example.com"));
    assert.ok(eml.includes("X-Priority: 2"));
  });

  it("should not include headers section when headers is undefined", () => {
    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
    };

    const eml = convertToEml(parsed);

    assert.ok(!eml.includes("Message-ID:"));
    assert.ok(!eml.includes("In-Reply-To:"));
    assert.ok(!eml.includes("References:"));
    assert.ok(!eml.includes("Reply-To:"));
    assert.ok(!eml.includes("X-Priority:"));
  });

  it("should place message headers after MIME-Version header", () => {
    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        messageId: "<test@example.com>",
      },
    };

    const eml = convertToEml(parsed);
    const mimeVersionIndex = eml.indexOf("MIME-Version:");
    const messageIdIndex = eml.indexOf("Message-ID:");
    const contentTypeIndex = eml.indexOf("Content-Type:");

    assert.ok(mimeVersionIndex < messageIdIndex, "MIME-Version should come before Message-ID");
    assert.ok(messageIdIndex < contentTypeIndex, "Message-ID should come before Content-Type");
  });
});

describe("extractBodyFromRtf", () => {
  it("should extract plain text from simple RTF", () => {
    // Simple RTF with plain text
    const rtf = "{\\rtf1\\ansi Hello World}";
    const uncompressed = createUncompressedRtf(rtf);
    const result = extractBodyFromRtf(uncompressed);

    assert.ok(result !== null, "Should return a result");
    assert.ok(result!.text.includes("Hello World"), "Should extract text content");
    assert.strictEqual(result!.html, undefined, "Should not have HTML for plain text RTF");
  });

  it("should extract HTML from RTF-encapsulated HTML (fromhtml)", () => {
    // RTF with encapsulated HTML using \fromhtml1
    const rtf = "{\\rtf1\\ansi\\fromhtml1 {\\*\\htmltag64 <html>}{\\*\\htmltag64 <body>}Hello HTML{\\*\\htmltag64 </body>}{\\*\\htmltag64 </html>}}";
    const uncompressed = createUncompressedRtf(rtf);
    const result = extractBodyFromRtf(uncompressed);

    assert.ok(result !== null, "Should return a result");
    assert.ok(result!.html !== undefined, "Should extract HTML");
    assert.ok(result!.html!.includes("Hello HTML"), "HTML should contain content");
    assert.ok(result!.text.includes("Hello HTML"), "Should also provide plain text fallback");
  });

  it("should return null for invalid compressed RTF", () => {
    const invalidData = [0, 1, 2, 3, 4, 5];
    const result = extractBodyFromRtf(invalidData);

    assert.strictEqual(result, null, "Should return null for invalid data");
  });

  it("should return null for empty input", () => {
    const result = extractBodyFromRtf([]);

    assert.strictEqual(result, null, "Should return null for empty input");
  });

  it("should handle RTF with special characters", () => {
    const rtf = "{\\rtf1\\ansi Caf\\'e9 r\\'e9sum\\'e9}";
    const uncompressed = createUncompressedRtf(rtf);
    const result = extractBodyFromRtf(uncompressed);

    assert.ok(result !== null, "Should return a result");
    // The text should contain the decoded special characters
    assert.ok(result!.text.length > 0, "Should have extracted text");
  });

  it("should strip HTML tags when providing plain text from HTML content", () => {
    const rtf = "{\\rtf1\\ansi\\fromhtml1 {\\*\\htmltag64 <p>}Paragraph text{\\*\\htmltag64 </p>}}";
    const uncompressed = createUncompressedRtf(rtf);
    const result = extractBodyFromRtf(uncompressed);

    assert.ok(result !== null, "Should return a result");
    if (result!.html) {
      // Plain text version should not contain HTML tags
      assert.ok(!result!.text.includes("<p>"), "Plain text should not contain <p> tag");
      assert.ok(!result!.text.includes("</p>"), "Plain text should not contain </p> tag");
    }
  });
});
