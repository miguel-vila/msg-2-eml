import { describe, it } from "node:test";
import assert from "node:assert";
import { convertToEml, formatSender, mapToXPriority, foldHeader, extractBodyFromRtf, parseMsg, encodeRfc2231, formatFilenameParams, encodeRfc2047, encodeDisplayName } from "./msg-to-eml.js";

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

describe("convertToEml with inline images", () => {
  it("should use Content-Disposition: inline for attachments with contentId", () => {
    const parsed = {
      subject: "Email with inline image",
      from: "sender@example.com",
      recipients: [{ name: "", email: "recipient@example.com", type: "to" as const }],
      date: new Date(),
      body: "See the image",
      bodyHtml: '<html><body><img src="cid:image001@domain.com"></body></html>',
      attachments: [
        {
          fileName: "image.png",
          content: new Uint8Array([137, 80, 78, 71]), // PNG magic bytes
          contentType: "image/png",
          contentId: "image001@domain.com",
        },
      ],
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("Content-Disposition: inline"), "Should have inline disposition");
    assert.ok(eml.includes("Content-ID: <image001@domain.com>"), "Should have Content-ID header with angle brackets");
    assert.ok(eml.includes("multipart/related"), "Should use multipart/related for inline images");
  });

  it("should use multipart/related for HTML with inline images", () => {
    const parsed = {
      subject: "HTML with inline image",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Plain text",
      bodyHtml: '<img src="cid:img1">',
      attachments: [
        {
          fileName: "test.jpg",
          content: new Uint8Array([255, 216, 255]), // JPEG magic bytes
          contentType: "image/jpeg",
          contentId: "img1",
        },
      ],
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("multipart/related"), "Should use multipart/related");
    assert.ok(eml.includes("multipart/alternative"), "Should contain multipart/alternative for text+html");
  });

  it("should handle mixed inline and regular attachments", () => {
    const parsed = {
      subject: "Mixed attachments",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "See attached",
      bodyHtml: '<img src="cid:inline1"><p>Check the attachment</p>',
      attachments: [
        {
          fileName: "inline-image.png",
          content: new Uint8Array([137, 80, 78, 71]),
          contentType: "image/png",
          contentId: "inline1",
        },
        {
          fileName: "document.pdf",
          content: new Uint8Array([37, 80, 68, 70]), // PDF magic bytes
          contentType: "application/pdf",
        },
      ],
    };

    const eml = convertToEml(parsed);

    // Should have multipart/mixed at the top level
    assert.ok(eml.includes("multipart/mixed"), "Should have multipart/mixed for regular attachments");
    assert.ok(eml.includes("multipart/related"), "Should have multipart/related for inline images");

    // Inline attachment should have Content-ID and inline disposition
    assert.ok(eml.includes("Content-ID: <inline1>"), "Inline attachment should have Content-ID");
    assert.ok(eml.includes("Content-Disposition: inline"), "Inline attachment should have inline disposition");

    // Regular attachment should have attachment disposition
    assert.ok(eml.includes('Content-Disposition: attachment; filename="document.pdf"'), "Regular attachment should have attachment disposition");
  });

  it("should not add Content-ID header for regular attachments", () => {
    const parsed = {
      subject: "Regular attachment only",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "See attached",
      attachments: [
        {
          fileName: "document.pdf",
          content: new Uint8Array([37, 80, 68, 70]),
          contentType: "application/pdf",
        },
      ],
    };

    const eml = convertToEml(parsed);

    assert.ok(!eml.includes("Content-ID:"), "Regular attachment should not have Content-ID");
    assert.ok(eml.includes("Content-Disposition: attachment"), "Should have attachment disposition");
    assert.ok(!eml.includes("multipart/related"), "Should not use multipart/related without inline images");
  });

  it("should handle multiple inline images", () => {
    const parsed = {
      subject: "Multiple inline images",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Images",
      bodyHtml: '<img src="cid:img1"><img src="cid:img2">',
      attachments: [
        {
          fileName: "image1.png",
          content: new Uint8Array([137, 80, 78, 71]),
          contentType: "image/png",
          contentId: "img1",
        },
        {
          fileName: "image2.jpg",
          content: new Uint8Array([255, 216, 255]),
          contentType: "image/jpeg",
          contentId: "img2",
        },
      ],
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("Content-ID: <img1>"), "Should include first Content-ID");
    assert.ok(eml.includes("Content-ID: <img2>"), "Should include second Content-ID");

    // Count inline dispositions
    const inlineCount = (eml.match(/Content-Disposition: inline/g) || []).length;
    assert.strictEqual(inlineCount, 2, "Should have two inline dispositions");
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

describe("convertToEml with embedded messages (message/rfc822)", () => {
  it("should use message/rfc822 content type for embedded messages", () => {
    const embeddedEmlContent = `From: nested@example.com\r
To: recipient@example.com\r
Subject: Forwarded Email\r
Date: Mon, 1 Jan 2024 12:00:00 +0000\r
MIME-Version: 1.0\r
Content-Type: text/plain; charset="utf-8"\r
\r
This is the forwarded message body.`;

    const parsed = {
      subject: "FW: Original Subject",
      from: "sender@example.com",
      recipients: [{ name: "", email: "recipient@example.com", type: "to" as const }],
      date: new Date("2024-01-15T10:30:00Z"),
      body: "Please see the forwarded email below.",
      attachments: [
        {
          fileName: "forwarded.eml",
          content: new TextEncoder().encode(embeddedEmlContent),
          contentType: "message/rfc822",
          isEmbeddedMessage: true,
        },
      ],
    };

    const eml = convertToEml(parsed);

    // Should have message/rfc822 content type
    assert.ok(eml.includes("Content-Type: message/rfc822"), "Should use message/rfc822 content type");
    // Should use 7bit encoding (not base64) for message/rfc822
    assert.ok(eml.includes("Content-Transfer-Encoding: 7bit"), "Should use 7bit encoding for embedded messages");
    // The embedded message content should be present
    assert.ok(eml.includes("From: nested@example.com"), "Embedded message headers should be present");
    assert.ok(eml.includes("Subject: Forwarded Email"), "Embedded message subject should be present");
    assert.ok(eml.includes("This is the forwarded message body"), "Embedded message body should be present");
  });

  it("should preserve filename for embedded messages", () => {
    const embeddedContent = "From: test@example.com\r\nSubject: Test\r\n\r\nBody";

    const parsed = {
      subject: "With embedded email",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "See attached",
      attachments: [
        {
          fileName: "original-email.eml",
          content: new TextEncoder().encode(embeddedContent),
          contentType: "message/rfc822",
        },
      ],
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes('name="original-email.eml"'), "Should preserve filename in Content-Type");
    assert.ok(eml.includes('filename="original-email.eml"'), "Should preserve filename in Content-Disposition");
  });

  it("should handle mixed regular attachments and embedded messages", () => {
    const embeddedContent = "From: test@example.com\r\nSubject: Test\r\n\r\nBody";

    const parsed = {
      subject: "Mixed attachments",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Multiple attachments",
      attachments: [
        {
          fileName: "document.pdf",
          content: new Uint8Array([37, 80, 68, 70]), // PDF magic bytes
          contentType: "application/pdf",
        },
        {
          fileName: "forwarded.eml",
          content: new TextEncoder().encode(embeddedContent),
          contentType: "message/rfc822",
        },
      ],
    };

    const eml = convertToEml(parsed);

    // Regular attachment should use base64
    assert.ok(eml.includes("Content-Type: application/pdf"), "Should have PDF attachment");
    // Embedded message should use 7bit
    assert.ok(eml.includes("Content-Type: message/rfc822"), "Should have embedded message");
    // Count encoding types
    const base64Count = (eml.match(/Content-Transfer-Encoding: base64/g) || []).length;
    const sevenBitCount = (eml.match(/Content-Transfer-Encoding: 7bit/g) || []).length;
    assert.strictEqual(base64Count, 1, "Should have one base64 encoded attachment");
    assert.strictEqual(sevenBitCount, 1, "Should have one 7bit encoded attachment");
  });

  it("should handle deeply nested embedded messages", () => {
    // Inner embedded message
    const innerEml = `From: inner@example.com\r
Subject: Inner Message\r
\r
Inner body`;

    // Outer embedded message containing the inner one
    const outerEml = `From: outer@example.com\r
Subject: Outer Message\r
MIME-Version: 1.0\r
Content-Type: multipart/mixed; boundary="nested"\r
\r
--nested\r
Content-Type: text/plain\r
\r
Outer body\r
--nested\r
Content-Type: message/rfc822\r
\r
${innerEml}\r
--nested--`;

    const parsed = {
      subject: "Forwarded chain",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "See forwarded chain",
      attachments: [
        {
          fileName: "chain.eml",
          content: new TextEncoder().encode(outerEml),
          contentType: "message/rfc822",
        },
      ],
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("From: outer@example.com"), "Should contain outer embedded message");
    assert.ok(eml.includes("From: inner@example.com"), "Should contain inner embedded message");
  });
});

describe("encodeRfc2231", () => {
  it("should encode ASCII filename without special characters", () => {
    const result = encodeRfc2231("document.pdf");
    assert.strictEqual(result, "UTF-8''document.pdf");
  });

  it("should encode non-ASCII characters using percent encoding", () => {
    const result = encodeRfc2231("naïve.pdf");
    // ï is UTF-8 bytes: 0xC3 0xAF
    assert.strictEqual(result, "UTF-8''na%C3%AFve.pdf");
  });

  it("should encode multiple non-ASCII characters", () => {
    const result = encodeRfc2231("Café résumé.pdf");
    // é is UTF-8 bytes: 0xC3 0xA9
    // Space is 0x20, needs encoding
    assert.strictEqual(result, "UTF-8''Caf%C3%A9%20r%C3%A9sum%C3%A9.pdf");
  });

  it("should encode special characters like spaces", () => {
    const result = encodeRfc2231("my document.pdf");
    assert.strictEqual(result, "UTF-8''my%20document.pdf");
  });

  it("should preserve hyphens, dots, and underscores", () => {
    const result = encodeRfc2231("my-doc_2024.final.pdf");
    assert.strictEqual(result, "UTF-8''my-doc_2024.final.pdf");
  });

  it("should encode parentheses and brackets", () => {
    const result = encodeRfc2231("report (final).pdf");
    assert.strictEqual(result, "UTF-8''report%20%28final%29.pdf");
  });

  it("should encode Japanese characters", () => {
    const result = encodeRfc2231("文書.pdf");
    // 文 = E6 96 87, 書 = E6 9B B8
    assert.strictEqual(result, "UTF-8''%E6%96%87%E6%9B%B8.pdf");
  });

  it("should encode emoji", () => {
    const result = encodeRfc2231("📄document.pdf");
    // 📄 = F0 9F 93 84
    assert.strictEqual(result, "UTF-8''%F0%9F%93%84document.pdf");
  });
});

describe("formatFilenameParams", () => {
  it("should return simple quoted format for ASCII filenames", () => {
    const result = formatFilenameParams("document.pdf");
    assert.strictEqual(result.name, 'name="document.pdf"');
    assert.strictEqual(result.disposition, 'filename="document.pdf"');
  });

  it("should return RFC 2231 encoded format for non-ASCII filenames", () => {
    const result = formatFilenameParams("naïve.pdf");
    assert.strictEqual(result.name, "name*=UTF-8''na%C3%AFve.pdf");
    assert.strictEqual(result.disposition, "filename*=UTF-8''na%C3%AFve.pdf");
  });

  it("should return RFC 2231 encoded format for filenames with spaces", () => {
    const result = formatFilenameParams("my document.pdf");
    // ASCII with spaces should still use simple format
    assert.strictEqual(result.name, 'name="my document.pdf"');
    assert.strictEqual(result.disposition, 'filename="my document.pdf"');
  });
});

describe("convertToEml with non-ASCII filenames", () => {
  it("should use RFC 2231 encoding for non-ASCII attachment filenames", () => {
    const parsed = {
      subject: "Test with special filename",
      from: "sender@example.com",
      recipients: [{ name: "", email: "recipient@example.com", type: "to" as const }],
      date: new Date(),
      body: "See attached",
      attachments: [
        {
          fileName: "naïve.pdf",
          content: new Uint8Array([37, 80, 68, 70]),
          contentType: "application/pdf",
        },
      ],
    };

    const eml = convertToEml(parsed);

    // Should use RFC 2231 encoded filename
    assert.ok(eml.includes("name*=UTF-8''na%C3%AFve.pdf"), "Content-Type should use RFC 2231 encoding");
    assert.ok(eml.includes("filename*=UTF-8''na%C3%AFve.pdf"), "Content-Disposition should use RFC 2231 encoding");
    // Should NOT contain the unencoded filename in quotes
    assert.ok(!eml.includes('name="naïve.pdf"'), "Should not use simple quoted format for non-ASCII");
    assert.ok(!eml.includes('filename="naïve.pdf"'), "Should not use simple quoted format for non-ASCII");
  });

  it("should keep ASCII filenames in simple quoted format", () => {
    const parsed = {
      subject: "Test with ASCII filename",
      from: "sender@example.com",
      recipients: [{ name: "", email: "recipient@example.com", type: "to" as const }],
      date: new Date(),
      body: "See attached",
      attachments: [
        {
          fileName: "document.pdf",
          content: new Uint8Array([37, 80, 68, 70]),
          contentType: "application/pdf",
        },
      ],
    };

    const eml = convertToEml(parsed);

    // Should use simple quoted format
    assert.ok(eml.includes('name="document.pdf"'), "Content-Type should use simple quoted format");
    assert.ok(eml.includes('filename="document.pdf"'), "Content-Disposition should use simple quoted format");
    // Should NOT contain RFC 2231 encoding
    assert.ok(!eml.includes("name*="), "Should not use RFC 2231 for ASCII filenames");
    assert.ok(!eml.includes("filename*="), "Should not use RFC 2231 for ASCII filenames");
  });

  it("should handle inline attachments with non-ASCII filenames", () => {
    const parsed = {
      subject: "Test with inline non-ASCII",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Text",
      bodyHtml: '<img src="cid:img1">',
      attachments: [
        {
          fileName: "übung.png",
          content: new Uint8Array([137, 80, 78, 71]),
          contentType: "image/png",
          contentId: "img1",
        },
      ],
    };

    const eml = convertToEml(parsed);

    // ü is UTF-8 bytes: 0xC3 0xBC
    assert.ok(eml.includes("name*=UTF-8''%C3%BCbung.png"), "Inline Content-Type should use RFC 2231");
    assert.ok(eml.includes("filename*=UTF-8''%C3%BCbung.png"), "Inline Content-Disposition should use RFC 2231");
    assert.ok(eml.includes("Content-Disposition: inline"), "Should have inline disposition");
  });

  it("should handle mixed ASCII and non-ASCII filenames", () => {
    const parsed = {
      subject: "Mixed filenames",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Multiple attachments",
      attachments: [
        {
          fileName: "ascii-doc.pdf",
          content: new Uint8Array([37, 80, 68, 70]),
          contentType: "application/pdf",
        },
        {
          fileName: "日本語.txt",
          content: new Uint8Array([72, 101, 108, 108, 111]),
          contentType: "text/plain",
        },
      ],
    };

    const eml = convertToEml(parsed);

    // ASCII filename should use simple format
    assert.ok(eml.includes('name="ascii-doc.pdf"'));
    assert.ok(eml.includes('filename="ascii-doc.pdf"'));

    // Japanese filename should use RFC 2231
    assert.ok(eml.includes("name*=UTF-8''"), "Japanese filename should use RFC 2231");
    assert.ok(!eml.includes('name="日本語.txt"'), "Should not use simple quoted format for Japanese");
  });
});

describe("encodeRfc2047", () => {
  it("should return ASCII text unchanged", () => {
    const result = encodeRfc2047("Hello World");
    assert.strictEqual(result, "Hello World");
  });

  it("should return empty string unchanged", () => {
    const result = encodeRfc2047("");
    assert.strictEqual(result, "");
  });

  it("should encode simple non-ASCII text", () => {
    const result = encodeRfc2047("Café");
    // "Café" in UTF-8 is [67, 97, 102, C3, A9] -> base64: Q2Fmw6k=
    assert.strictEqual(result, "=?UTF-8?B?Q2Fmw6k=?=");
  });

  it("should encode text with multiple non-ASCII characters", () => {
    const result = encodeRfc2047("Ñoño");
    // Should be a valid RFC 2047 encoded-word
    assert.ok(result.startsWith("=?UTF-8?B?"), "Should start with UTF-8 B encoding prefix");
    assert.ok(result.endsWith("?="), "Should end with encoded-word suffix");
  });

  it("should encode Japanese text", () => {
    const result = encodeRfc2047("こんにちは");
    assert.ok(result.startsWith("=?UTF-8?B?"), "Should use UTF-8 Base64 encoding");
    assert.ok(result.endsWith("?="), "Should end with encoded-word suffix");
    // Decode to verify
    const base64Part = result.slice(10, -2); // Remove =?UTF-8?B? and ?=
    const decoded = Buffer.from(base64Part, "base64").toString("utf-8");
    assert.strictEqual(decoded, "こんにちは", "Should decode back to original");
  });

  it("should split long encoded text into multiple encoded-words", () => {
    // Create a long non-ASCII string that will exceed 75 chars when encoded
    const longText = "日本語のテキストがとても長い場合はどうなりますか";
    const result = encodeRfc2047(longText);

    // Should contain multiple encoded-words separated by space
    const encodedWords = result.split(" ");
    assert.ok(encodedWords.length >= 1, "Should have at least one encoded-word");

    // Each encoded word should be <= 75 characters
    for (const word of encodedWords) {
      assert.ok(word.length <= 75, `Encoded word too long: ${word.length} chars`);
      assert.ok(word.startsWith("=?UTF-8?B?"), "Each word should start with encoding prefix");
      assert.ok(word.endsWith("?="), "Each word should end with suffix");
    }

    // Verify we can decode the full text back
    const decodedParts = encodedWords.map((word) => {
      const base64Part = word.slice(10, -2);
      return Buffer.from(base64Part, "base64").toString("utf-8");
    });
    assert.strictEqual(decodedParts.join(""), longText, "Should decode back to original");
  });

  it("should encode emoji correctly", () => {
    const result = encodeRfc2047("Hello 👋 World");
    assert.ok(result.startsWith("=?UTF-8?B?"), "Should use UTF-8 Base64 encoding");
    // Verify decoding
    const base64Part = result.slice(10, -2);
    const decoded = Buffer.from(base64Part, "base64").toString("utf-8");
    assert.strictEqual(decoded, "Hello 👋 World");
  });

  it("should handle mixed ASCII and non-ASCII", () => {
    const result = encodeRfc2047("Hello Wörld");
    assert.ok(result.startsWith("=?UTF-8?B?"), "Should encode the entire string");
    const base64Part = result.slice(10, -2);
    const decoded = Buffer.from(base64Part, "base64").toString("utf-8");
    assert.strictEqual(decoded, "Hello Wörld");
  });
});

describe("encodeDisplayName", () => {
  it("should return empty string for empty input", () => {
    const result = encodeDisplayName("");
    assert.strictEqual(result, "");
  });

  it("should wrap ASCII names in quotes", () => {
    const result = encodeDisplayName("John Doe");
    assert.strictEqual(result, '"John Doe"');
  });

  it("should encode non-ASCII names with RFC 2047", () => {
    const result = encodeDisplayName("José García");
    assert.ok(result.startsWith("=?UTF-8?B?"), "Should use RFC 2047 encoding");
    assert.ok(!result.includes('"'), "Should not be quoted");
  });

  it("should encode Japanese names", () => {
    const result = encodeDisplayName("山田太郎");
    assert.ok(result.startsWith("=?UTF-8?B?"), "Should use RFC 2047 encoding");
    // Verify decoding
    const base64Part = result.slice(10, -2);
    const decoded = Buffer.from(base64Part, "base64").toString("utf-8");
    assert.strictEqual(decoded, "山田太郎");
  });
});

describe("formatSender with non-ASCII names", () => {
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

describe("convertToEml with non-ASCII subjects", () => {
  it("should encode non-ASCII subject with RFC 2047", () => {
    const parsed = {
      subject: "Café résumé",
      from: "sender@example.com",
      recipients: [],
      date: new Date("2024-01-15T10:30:00Z"),
      body: "Hello",
      attachments: [],
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("Subject: =?UTF-8?B?"), "Should encode subject with RFC 2047");
    assert.ok(!eml.includes("Subject: Café"), "Should not contain unencoded subject");
  });

  it("should not encode ASCII-only subject", () => {
    const parsed = {
      subject: "Hello World",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test",
      attachments: [],
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("Subject: Hello World"), "ASCII subject should not be encoded");
    assert.ok(!eml.includes("=?UTF-8?B?"), "Should not contain encoding");
  });

  it("should encode Japanese subject", () => {
    const parsed = {
      subject: "日本語のメール",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test",
      attachments: [],
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("Subject: =?UTF-8?B?"), "Should encode Japanese subject");

    // Extract and decode the subject to verify
    const subjectLine = eml.split("\r\n").find((line) => line.startsWith("Subject:"));
    assert.ok(subjectLine, "Should have Subject header");

    // Handle potential folding by normalizing
    const normalizedSubject = subjectLine!.replace(/\r\n[\t ]/g, " ");
    const encodedParts = normalizedSubject.match(/=\?UTF-8\?B\?[^?]+\?=/g) || [];
    const decoded = encodedParts
      .map((part) => {
        const base64 = part.slice(10, -2);
        return Buffer.from(base64, "base64").toString("utf-8");
      })
      .join("");

    assert.strictEqual(decoded, "日本語のメール", "Decoded subject should match original");
  });
});

describe("convertToEml with non-ASCII display names", () => {
  it("should encode non-ASCII To recipient name", () => {
    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [{ name: "François Müller", email: "francois@example.com", type: "to" as const }],
      date: new Date(),
      body: "Test",
      attachments: [],
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("To: =?UTF-8?B?"), "Should encode non-ASCII recipient name");
    assert.ok(eml.includes("<francois@example.com>"), "Should include email address");
  });

  it("should encode non-ASCII Cc recipient name", () => {
    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [{ name: "北京用户", email: "beijing@example.com", type: "cc" as const }],
      date: new Date(),
      body: "Test",
      attachments: [],
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("Cc: =?UTF-8?B?"), "Should encode Chinese recipient name");
    assert.ok(eml.includes("<beijing@example.com>"), "Should include email address");
  });

  it("should encode non-ASCII From display name", () => {
    const parsed = {
      subject: "Test",
      from: "=?UTF-8?B?5bGx55Sw5aSq6YOO?= <yamada@example.com>", // Pre-encoded by formatSender
      recipients: [],
      date: new Date(),
      body: "Test",
      attachments: [],
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("From: =?UTF-8?B?"), "Should have encoded From name");
  });

  it("should handle mixed ASCII and non-ASCII recipients", () => {
    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [
        { name: "John Doe", email: "john@example.com", type: "to" as const },
        { name: "José García", email: "jose@example.com", type: "to" as const },
      ],
      date: new Date(),
      body: "Test",
      attachments: [],
    };

    const eml = convertToEml(parsed);

    // ASCII name should be quoted
    assert.ok(eml.includes('"John Doe" <john@example.com>'), "ASCII name should be quoted");
    // Non-ASCII name should be encoded
    assert.ok(eml.includes("=?UTF-8?B?"), "Non-ASCII name should be encoded");
    assert.ok(eml.includes("<jose@example.com>"), "Should include Jose's email");
  });
});
