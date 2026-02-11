import assert from "node:assert";
import { describe, it } from "node:test";
import { convertToEml } from "../eml.js";

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
    const longSubject =
      "This is a very very very long subject line that will definitely exceed the 78 character limit according to RFC 5322";
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
    assert.ok(
      eml.includes('Content-Disposition: attachment; filename="document.pdf"'),
      "Regular attachment should have attachment disposition",
    );
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
    const normalizedSubject = subjectLine?.replace(/\r\n[\t ]/g, " ");
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

describe("convertToEml with calendar events", () => {
  it("should include text/calendar part for calendar events", () => {
    const parsed = {
      subject: "Team Meeting",
      from: "organizer@example.com",
      recipients: [{ name: "Attendee", email: "attendee@example.com", type: "to" as const }],
      date: new Date("2024-03-15T10:00:00Z"),
      body: "Please join the meeting",
      attachments: [],
      calendarEvent: {
        startTime: new Date("2024-03-15T14:00:00Z"),
        endTime: new Date("2024-03-15T15:00:00Z"),
        location: "Conference Room",
        organizer: "organizer@example.com",
        attendees: ["attendee@example.com"],
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("multipart/alternative"), "Should use multipart/alternative");
    assert.ok(eml.includes("text/plain"), "Should have text/plain part");
    assert.ok(eml.includes("text/calendar"), "Should have text/calendar part");
    assert.ok(eml.includes('charset="utf-8"; method=REQUEST'), "Should have calendar charset and method");
    assert.ok(eml.includes("BEGIN:VCALENDAR"), "Should contain VCALENDAR");
    assert.ok(eml.includes("BEGIN:VEVENT"), "Should contain VEVENT");
    assert.ok(eml.includes("DTSTART:20240315T140000Z"), "Should have correct start time");
    assert.ok(eml.includes("DTEND:20240315T150000Z"), "Should have correct end time");
    assert.ok(eml.includes("LOCATION:Conference Room"), "Should have location");
  });

  it("should include both HTML and calendar parts when both are present", () => {
    const parsed = {
      subject: "Team Meeting",
      from: "organizer@example.com",
      recipients: [],
      date: new Date(),
      body: "Please join the meeting",
      bodyHtml: "<p>Please join the meeting</p>",
      attachments: [],
      calendarEvent: {
        startTime: new Date("2024-03-15T14:00:00Z"),
        endTime: new Date("2024-03-15T15:00:00Z"),
        attendees: [],
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("multipart/alternative"), "Should use multipart/alternative");
    assert.ok(eml.includes("text/plain"), "Should have text/plain part");
    assert.ok(eml.includes("text/html"), "Should have text/html part");
    assert.ok(eml.includes("text/calendar"), "Should have text/calendar part");
  });

  it("should handle calendar event with attachments", () => {
    const parsed = {
      subject: "Meeting with Agenda",
      from: "organizer@example.com",
      recipients: [],
      date: new Date(),
      body: "Agenda attached",
      attachments: [
        {
          fileName: "agenda.pdf",
          content: new Uint8Array([37, 80, 68, 70]),
          contentType: "application/pdf",
        },
      ],
      calendarEvent: {
        startTime: new Date("2024-03-15T14:00:00Z"),
        endTime: new Date("2024-03-15T15:00:00Z"),
        attendees: [],
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("multipart/mixed"), "Should use multipart/mixed for attachments");
    assert.ok(eml.includes("multipart/alternative"), "Should have multipart/alternative for content");
    assert.ok(eml.includes("text/calendar"), "Should have calendar part");
    assert.ok(eml.includes("application/pdf"), "Should have PDF attachment");
  });

  it("should not include calendar part for regular emails", () => {
    const parsed = {
      subject: "Regular Email",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Just a regular email",
      attachments: [],
    };

    const eml = convertToEml(parsed);

    assert.ok(!eml.includes("text/calendar"), "Should not have text/calendar");
    assert.ok(!eml.includes("BEGIN:VCALENDAR"), "Should not have VCALENDAR");
  });

  it("should include all attendees from the event", () => {
    const parsed = {
      subject: "Group Meeting",
      from: "organizer@example.com",
      recipients: [],
      date: new Date(),
      body: "Meeting notes",
      attachments: [],
      calendarEvent: {
        startTime: new Date("2024-03-15T14:00:00Z"),
        endTime: new Date("2024-03-15T15:00:00Z"),
        organizer: "organizer@example.com",
        attendees: ["alice@example.com", "bob@example.com", "charlie@example.com"],
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("ATTENDEE:mailto:alice@example.com"), "Should have Alice as attendee");
    assert.ok(eml.includes("ATTENDEE:mailto:bob@example.com"), "Should have Bob as attendee");
    assert.ok(eml.includes("ATTENDEE:mailto:charlie@example.com"), "Should have Charlie as attendee");
    assert.ok(eml.includes("ORGANIZER:mailto:organizer@example.com"), "Should have organizer");
  });
});

describe("convertToEml with read and delivery receipt headers", () => {
  it("should include Disposition-Notification-To header when read receipt is requested", () => {
    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [{ name: "", email: "recipient@example.com", type: "to" as const }],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        dispositionNotificationTo: "sender@example.com",
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(
      eml.includes("Disposition-Notification-To: sender@example.com"),
      "Should have Disposition-Notification-To header",
    );
  });

  it("should include Return-Receipt-To header when delivery receipt is requested", () => {
    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [{ name: "", email: "recipient@example.com", type: "to" as const }],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        returnReceiptTo: "sender@example.com",
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("Return-Receipt-To: sender@example.com"), "Should have Return-Receipt-To header");
  });

  it("should include both receipt headers when both are requested", () => {
    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        dispositionNotificationTo: "sender@example.com",
        returnReceiptTo: "sender@example.com",
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(
      eml.includes("Disposition-Notification-To: sender@example.com"),
      "Should have Disposition-Notification-To header",
    );
    assert.ok(eml.includes("Return-Receipt-To: sender@example.com"), "Should have Return-Receipt-To header");
  });

  it("should not include receipt headers when not requested", () => {
    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
    };

    const eml = convertToEml(parsed);

    assert.ok(!eml.includes("Disposition-Notification-To:"), "Should not have Disposition-Notification-To header");
    assert.ok(!eml.includes("Return-Receipt-To:"), "Should not have Return-Receipt-To header");
  });

  it("should place receipt headers after other message headers", () => {
    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        messageId: "<test@example.com>",
        dispositionNotificationTo: "sender@example.com",
        returnReceiptTo: "sender@example.com",
      },
    };

    const eml = convertToEml(parsed);
    const mimeVersionIndex = eml.indexOf("MIME-Version:");
    const messageIdIndex = eml.indexOf("Message-ID:");
    const dntIndex = eml.indexOf("Disposition-Notification-To:");
    const rrtIndex = eml.indexOf("Return-Receipt-To:");
    const contentTypeIndex = eml.indexOf("Content-Type:");

    assert.ok(mimeVersionIndex < messageIdIndex, "MIME-Version should come before Message-ID");
    assert.ok(messageIdIndex < dntIndex, "Message-ID should come before Disposition-Notification-To");
    assert.ok(dntIndex < rrtIndex, "Disposition-Notification-To should come before Return-Receipt-To");
    assert.ok(rrtIndex < contentTypeIndex, "Return-Receipt-To should come before Content-Type");
  });

  it("should work with other headers like X-Priority", () => {
    const parsed = {
      subject: "Urgent with Receipt",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Urgent message",
      attachments: [],
      headers: {
        priority: 1,
        dispositionNotificationTo: "sender@example.com",
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("X-Priority: 1"), "Should have X-Priority header");
    assert.ok(
      eml.includes("Disposition-Notification-To: sender@example.com"),
      "Should have Disposition-Notification-To header",
    );
  });
});

describe("convertToEml with transport message headers", () => {
  it("should include Received headers from transport headers", () => {
    const transportHeaders = `Received: from mail.example.com by mx.example.org with SMTP id abc123; Mon, 15 Jan 2024 10:30:00 +0000\r
Received: from localhost by mail.example.com with ESMTP id def456; Mon, 15 Jan 2024 10:29:59 +0000\r
From: sender@example.com\r
To: recipient@example.com\r
Subject: Test`;

    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [{ name: "", email: "recipient@example.com", type: "to" as const }],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        transportMessageHeaders: transportHeaders,
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("Received: from mail.example.com"), "Should include first Received header");
    assert.ok(eml.includes("Received: from localhost"), "Should include second Received header");
  });

  it("should include DKIM-Signature from transport headers", () => {
    const transportHeaders = `DKIM-Signature: v=1; a=rsa-sha256; d=example.com; s=selector;\r
\tc=relaxed/relaxed; q=dns/txt; h=from:to:subject:date;\r
\tbh=abc123=; b=xyz789=\r
From: sender@example.com`;

    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        transportMessageHeaders: transportHeaders,
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("DKIM-Signature: v=1; a=rsa-sha256"), "Should include DKIM-Signature header");
    assert.ok(eml.includes("c=relaxed/relaxed"), "Should preserve folded DKIM header content");
  });

  it("should include Authentication-Results from transport headers", () => {
    const transportHeaders = `Authentication-Results: mx.example.com;\r
\tdkim=pass header.d=example.com;\r
\tspf=pass smtp.mailfrom=sender@example.com\r
From: sender@example.com`;

    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        transportMessageHeaders: transportHeaders,
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("Authentication-Results: mx.example.com"), "Should include Authentication-Results header");
    assert.ok(eml.includes("dkim=pass"), "Should preserve authentication result details");
  });

  it("should include X-* headers from transport headers", () => {
    const transportHeaders = `X-Spam-Status: No, score=-2.0\r
X-Spam-Score: -2.0\r
X-Originating-IP: [192.168.1.1]\r
From: sender@example.com`;

    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        transportMessageHeaders: transportHeaders,
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("X-Spam-Status: No, score=-2.0"), "Should include X-Spam-Status header");
    assert.ok(eml.includes("X-Spam-Score: -2.0"), "Should include X-Spam-Score header");
    assert.ok(eml.includes("X-Originating-IP: [192.168.1.1]"), "Should include X-Originating-IP header");
  });

  it("should exclude headers that are already generated (From, To, Subject, etc.)", () => {
    const transportHeaders = `Received: from mail.example.com by mx.example.org\r
From: original-sender@example.com\r
To: original-recipient@example.com\r
Subject: Original Subject\r
Date: Mon, 1 Jan 2024 00:00:00 +0000\r
MIME-Version: 1.0\r
Message-ID: <original@example.com>\r
Content-Type: text/plain`;

    const parsed = {
      subject: "New Subject",
      from: "new-sender@example.com",
      recipients: [{ name: "", email: "new-recipient@example.com", type: "to" as const }],
      date: new Date("2024-01-15T10:30:00Z"),
      body: "Test body",
      attachments: [],
      headers: {
        messageId: "<new@example.com>",
        transportMessageHeaders: transportHeaders,
      },
    };

    const eml = convertToEml(parsed);

    // Should include Received (transport header we want to preserve)
    assert.ok(eml.includes("Received: from mail.example.com"), "Should include Received header");

    // Should use our generated headers, not the ones from transport
    assert.ok(eml.includes("From: new-sender@example.com"), "Should use generated From header");
    assert.ok(eml.includes("To: new-recipient@example.com"), "Should use generated To header");
    assert.ok(eml.includes("Subject: New Subject"), "Should use generated Subject header");
    assert.ok(eml.includes("Message-ID: <new@example.com>"), "Should use generated Message-ID header");

    // Should NOT include duplicates from transport headers
    assert.ok(!eml.includes("From: original-sender@example.com"), "Should not include original From");
    assert.ok(!eml.includes("To: original-recipient@example.com"), "Should not include original To");
    assert.ok(!eml.includes("Subject: Original Subject"), "Should not include original Subject");
    assert.ok(!eml.includes("Message-ID: <original@example.com>"), "Should not include original Message-ID");
  });

  it("should place transport headers before Content-Type", () => {
    const transportHeaders = `Received: from mail.example.com by mx.example.org\r
Authentication-Results: mx.example.com; spf=pass`;

    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        transportMessageHeaders: transportHeaders,
      },
    };

    const eml = convertToEml(parsed);

    const receivedIndex = eml.indexOf("Received:");
    const authResultsIndex = eml.indexOf("Authentication-Results:");
    const contentTypeIndex = eml.indexOf("Content-Type:");

    assert.ok(receivedIndex > 0, "Should have Received header");
    assert.ok(authResultsIndex > 0, "Should have Authentication-Results header");
    assert.ok(receivedIndex < contentTypeIndex, "Received should come before Content-Type");
    assert.ok(authResultsIndex < contentTypeIndex, "Authentication-Results should come before Content-Type");
  });

  it("should handle transport headers with LF line endings", () => {
    const transportHeaders =
      "Received: from mail.example.com by mx.example.org\nX-Custom: value\nFrom: sender@example.com";

    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        transportMessageHeaders: transportHeaders,
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("Received: from mail.example.com"), "Should include Received header");
    assert.ok(eml.includes("X-Custom: value"), "Should include X-Custom header");
  });

  it("should not include transport headers when not present", () => {
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

    assert.ok(!eml.includes("Received:"), "Should not have Received header when transport headers not present");
    assert.ok(!eml.includes("Authentication-Results:"), "Should not have Authentication-Results when not present");
  });

  it("should preserve complete email provenance chain", () => {
    const transportHeaders = `Received: from outgoing.example.com (192.168.1.100) by mx.destination.org with ESMTPS id abc123; Mon, 15 Jan 2024 10:30:00 +0000\r
Received: from internal.example.com (10.0.0.50) by outgoing.example.com with ESMTP id def456; Mon, 15 Jan 2024 10:29:58 +0000\r
Received: from sender-workstation.local (10.0.0.10) by internal.example.com with ESMTP id ghi789; Mon, 15 Jan 2024 10:29:55 +0000\r
Authentication-Results: mx.destination.org;\r
\tdkim=pass header.d=example.com header.s=selector header.b=abc123;\r
\tspf=pass (mx.destination.org: domain of sender@example.com designates 192.168.1.100 as permitted sender) smtp.mailfrom=sender@example.com;\r
\tdmarc=pass (p=REJECT sp=REJECT dis=NONE) header.from=example.com\r
DKIM-Signature: v=1; a=rsa-sha256; c=relaxed/relaxed; d=example.com;\r
\ts=selector; t=1705315795;\r
\th=from:to:subject:date:message-id:mime-version:content-type;\r
\tbh=47DEQpj8HBSa+/TImW+5JCeuQeRkm5NMpJWZG3hSuFU=;\r
\tb=XYZ789...=`;

    const parsed = {
      subject: "Important Email",
      from: "sender@example.com",
      recipients: [{ name: "Recipient", email: "recipient@destination.org", type: "to" as const }],
      date: new Date("2024-01-15T10:30:00Z"),
      body: "This email has complete provenance.",
      attachments: [],
      headers: {
        transportMessageHeaders: transportHeaders,
      },
    };

    const eml = convertToEml(parsed);

    // Verify all three Received headers are present (complete routing chain)
    const receivedMatches = eml.match(/Received:/g);
    assert.strictEqual(receivedMatches?.length, 3, "Should have all three Received headers");

    // Verify authentication chain
    assert.ok(eml.includes("dkim=pass"), "Should include DKIM result");
    assert.ok(eml.includes("spf=pass"), "Should include SPF result");
    assert.ok(eml.includes("dmarc=pass"), "Should include DMARC result");

    // Verify DKIM signature
    assert.ok(eml.includes("DKIM-Signature: v=1"), "Should include DKIM-Signature");
    assert.ok(eml.includes("a=rsa-sha256"), "Should preserve DKIM algorithm");
    assert.ok(eml.includes("d=example.com"), "Should preserve DKIM domain");
  });
});

describe("convertToEml with sender on behalf of", () => {
  it("should include Sender header when sender differs from From", () => {
    const parsed = {
      subject: "Test",
      from: '"Boss" <boss@example.com>',
      sender: '"Assistant" <assistant@example.com>',
      recipients: [{ name: "", email: "recipient@example.com", type: "to" as const }],
      date: new Date(),
      body: "Test body",
      attachments: [],
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes('From: "Boss" <boss@example.com>'), "Should have From header with boss");
    assert.ok(eml.includes('Sender: "Assistant" <assistant@example.com>'), "Should have Sender header with assistant");
  });

  it("should not include Sender header when not provided", () => {
    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("From: sender@example.com"), "Should have From header");
    assert.ok(!eml.includes("Sender:"), "Should not have Sender header");
  });

  it("should place Sender header immediately after From header", () => {
    const parsed = {
      subject: "Test",
      from: '"Boss" <boss@example.com>',
      sender: '"Assistant" <assistant@example.com>',
      recipients: [{ name: "", email: "recipient@example.com", type: "to" as const }],
      date: new Date(),
      body: "Test body",
      attachments: [],
    };

    const eml = convertToEml(parsed);

    const fromIndex = eml.indexOf("From:");
    const senderIndex = eml.indexOf("Sender:");
    const toIndex = eml.indexOf("To:");

    assert.ok(fromIndex > -1, "Should have From header");
    assert.ok(senderIndex > -1, "Should have Sender header");
    assert.ok(fromIndex < senderIndex, "From should come before Sender");
    assert.ok(senderIndex < toIndex, "Sender should come before To");
  });

  it("should fold long Sender headers", () => {
    const longName = "Very Long Assistant Name That Definitely Exceeds The Recommended Line Length For Email Headers";
    const parsed = {
      subject: "Test",
      from: '"Boss" <boss@example.com>',
      sender: `"${longName}" <assistant@example.com>`,
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("Sender:"), "Should have Sender header");
    // The email part should be present in the output (possibly after folding)
    assert.ok(eml.includes("<assistant@example.com>"), "Should include email address");
  });

  it("should handle encoded Sender display names", () => {
    const parsed = {
      subject: "Test",
      from: "=?UTF-8?B?5LiK5Y+4?= <boss@example.com>", // Japanese "上司" (boss)
      sender: "=?UTF-8?B?56eY5pu4?= <secretary@example.com>", // Japanese "秘書" (secretary)
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("From: =?UTF-8?B?"), "Should have encoded From header");
    assert.ok(eml.includes("Sender: =?UTF-8?B?"), "Should have encoded Sender header");
    assert.ok(eml.includes("<secretary@example.com>"), "Should include secretary email");
  });

  it("should exclude Sender from transport headers when we generate it", () => {
    const transportHeaders = `Received: from mail.example.com by mx.example.org\r
Sender: old-assistant@example.com\r
From: boss@example.com`;

    const parsed = {
      subject: "Test",
      from: '"Boss" <boss@example.com>',
      sender: '"New Assistant" <new-assistant@example.com>',
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        transportMessageHeaders: transportHeaders,
      },
    };

    const eml = convertToEml(parsed);

    // Should use our generated Sender, not the one from transport
    assert.ok(eml.includes('Sender: "New Assistant" <new-assistant@example.com>'), "Should use generated Sender");

    // Count Sender occurrences - should have exactly one
    const senderMatches = eml.match(/Sender:/g);
    assert.strictEqual(senderMatches?.length, 1, "Should have exactly one Sender header");
  });

  it("should work with all other headers", () => {
    const parsed = {
      subject: "Important Message",
      from: '"CEO" <ceo@example.com>',
      sender: '"Executive Assistant" <ea@example.com>',
      recipients: [
        { name: "Recipient", email: "recipient@example.com", type: "to" as const },
        { name: "CC Person", email: "cc@example.com", type: "cc" as const },
      ],
      date: new Date("2024-01-15T10:30:00Z"),
      body: "Test body",
      attachments: [],
      headers: {
        messageId: "<test@example.com>",
        priority: 1,
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes('From: "CEO" <ceo@example.com>'), "Should have From header");
    assert.ok(eml.includes('Sender: "Executive Assistant" <ea@example.com>'), "Should have Sender header");
    assert.ok(eml.includes("To:"), "Should have To header");
    assert.ok(eml.includes("Cc:"), "Should have Cc header");
    assert.ok(eml.includes("Subject: Important Message"), "Should have Subject header");
    assert.ok(eml.includes("Message-ID: <test@example.com>"), "Should have Message-ID header");
    assert.ok(eml.includes("X-Priority: 1"), "Should have X-Priority header");
  });
});

describe("convertToEml with conversation threading headers", () => {
  it("should include Thread-Index header when present", () => {
    // Example Thread-Index: a base64-encoded binary blob
    // The first 22 bytes represent the header GUID and timestamp
    const threadIndex = "AQHZDGHRmhV1R06lnVmPq8IAAAAnAA==";

    const parsed = {
      subject: "Re: Test Thread",
      from: "sender@example.com",
      recipients: [{ name: "", email: "recipient@example.com", type: "to" as const }],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        threadIndex,
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes(`Thread-Index: ${threadIndex}`), "Should have Thread-Index header");
  });

  it("should include Thread-Topic header when present", () => {
    const parsed = {
      subject: "Re: Test Thread",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        threadTopic: "Test Thread",
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("Thread-Topic: Test Thread"), "Should have Thread-Topic header");
  });

  it("should include both Thread-Index and Thread-Topic when present", () => {
    const threadIndex = "AQHZDGHRmhV1R06lnVmPq8IAAAAnAA==";

    const parsed = {
      subject: "Re: Important Discussion",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        threadIndex,
        threadTopic: "Important Discussion",
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes(`Thread-Index: ${threadIndex}`), "Should have Thread-Index header");
    assert.ok(eml.includes("Thread-Topic: Important Discussion"), "Should have Thread-Topic header");
  });

  it("should not include threading headers when not present", () => {
    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
    };

    const eml = convertToEml(parsed);

    assert.ok(!eml.includes("Thread-Index:"), "Should not have Thread-Index header when not present");
    assert.ok(!eml.includes("Thread-Topic:"), "Should not have Thread-Topic header when not present");
  });

  it("should encode non-ASCII Thread-Topic with RFC 2047", () => {
    const parsed = {
      subject: "Re: 日本語のスレッド",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        threadTopic: "日本語のスレッド",
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("Thread-Topic: =?UTF-8?B?"), "Should encode non-ASCII Thread-Topic with RFC 2047");
    assert.ok(!eml.includes("Thread-Topic: 日本語のスレッド"), "Should not contain unencoded non-ASCII topic");
  });

  it("should place threading headers before Content-Type", () => {
    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        messageId: "<test@example.com>",
        threadIndex: "AQHZDGHRmhV1R06lnVmPq8IAAAAnAA==",
        threadTopic: "Test Thread",
      },
    };

    const eml = convertToEml(parsed);

    const threadIndexPos = eml.indexOf("Thread-Index:");
    const threadTopicPos = eml.indexOf("Thread-Topic:");
    const contentTypePos = eml.indexOf("Content-Type:");

    assert.ok(threadIndexPos > 0, "Should have Thread-Index header");
    assert.ok(threadTopicPos > 0, "Should have Thread-Topic header");
    assert.ok(threadIndexPos < contentTypePos, "Thread-Index should come before Content-Type");
    assert.ok(threadTopicPos < contentTypePos, "Thread-Topic should come before Content-Type");
  });

  it("should fold long Thread-Index headers", () => {
    // A longer thread index with multiple response entries (each response adds 5 bytes)
    // This simulates a long email thread with many replies
    const longThreadIndex =
      "AQHZDGHRmhV1R06lnVmPq8IAAAAnAAEBxg5u0QAAABwAAQHGDm7SAAAAHAABAdkMYdGaFXVHTqWdWY+rwgAAACcAAQHZDGHRmhV1R06lnVmPq8IAAAAnAA==";

    const parsed = {
      subject: "Re: Re: Re: Long Thread",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        threadIndex: longThreadIndex,
      },
    };

    const eml = convertToEml(parsed);

    // Should contain the Thread-Index header
    assert.ok(eml.includes("Thread-Index:"), "Should have Thread-Index header");

    // The full value should be present (possibly folded)
    const normalizedEml = eml.replace(/\r\n[\t ]/g, "");
    assert.ok(normalizedEml.includes(longThreadIndex), "Full Thread-Index value should be present");
  });

  it("should work alongside other message headers", () => {
    const parsed = {
      subject: "Re: Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        messageId: "<reply@example.com>",
        inReplyTo: "<original@example.com>",
        references: "<original@example.com>",
        priority: 3,
        threadIndex: "AQHZDGHRmhV1R06lnVmPq8IAAAAnAA==",
        threadTopic: "Test",
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("Message-ID: <reply@example.com>"), "Should have Message-ID");
    assert.ok(eml.includes("In-Reply-To: <original@example.com>"), "Should have In-Reply-To");
    assert.ok(eml.includes("References: <original@example.com>"), "Should have References");
    assert.ok(eml.includes("X-Priority: 3"), "Should have X-Priority");
    assert.ok(eml.includes("Thread-Index:"), "Should have Thread-Index");
    assert.ok(eml.includes("Thread-Topic: Test"), "Should have Thread-Topic");
  });

  it("should exclude Thread-Index and Thread-Topic from transport headers", () => {
    const transportHeaders = `Received: from mail.example.com by mx.example.org\r
Thread-Index: AQHZDGHRmhV1R06lnVmPq8IAAAAnAA==\r
Thread-Topic: Original Topic\r
From: sender@example.com`;

    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        threadIndex: "NewThreadIndex123==",
        threadTopic: "New Topic",
        transportMessageHeaders: transportHeaders,
      },
    };

    const eml = convertToEml(parsed);

    // Should use our generated headers, not the ones from transport
    assert.ok(eml.includes("Thread-Index: NewThreadIndex123=="), "Should use generated Thread-Index");
    assert.ok(eml.includes("Thread-Topic: New Topic"), "Should use generated Thread-Topic");

    // Count occurrences - should have exactly one of each
    const threadIndexMatches = eml.match(/Thread-Index:/g);
    const threadTopicMatches = eml.match(/Thread-Topic:/g);
    assert.strictEqual(threadIndexMatches?.length, 1, "Should have exactly one Thread-Index header");
    assert.strictEqual(threadTopicMatches?.length, 1, "Should have exactly one Thread-Topic header");
  });
});

describe("convertToEml with sensitivity header", () => {
  it("should include Sensitivity: Personal header", () => {
    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [{ name: "", email: "recipient@example.com", type: "to" as const }],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        sensitivity: "Personal" as const,
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("Sensitivity: Personal"), "Should have Sensitivity: Personal header");
  });

  it("should include Sensitivity: Private header", () => {
    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        sensitivity: "Private" as const,
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("Sensitivity: Private"), "Should have Sensitivity: Private header");
  });

  it("should include Sensitivity: Company-Confidential header", () => {
    const parsed = {
      subject: "Confidential Report",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Confidential information",
      attachments: [],
      headers: {
        sensitivity: "Company-Confidential" as const,
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(
      eml.includes("Sensitivity: Company-Confidential"),
      "Should have Sensitivity: Company-Confidential header",
    );
  });

  it("should not include Sensitivity header when not present", () => {
    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
    };

    const eml = convertToEml(parsed);

    assert.ok(!eml.includes("Sensitivity:"), "Should not have Sensitivity header when not present");
  });

  it("should place Sensitivity header after X-Priority", () => {
    const parsed = {
      subject: "Urgent Private Message",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        priority: 1,
        sensitivity: "Private" as const,
      },
    };

    const eml = convertToEml(parsed);

    const priorityIndex = eml.indexOf("X-Priority:");
    const sensitivityIndex = eml.indexOf("Sensitivity:");
    const contentTypeIndex = eml.indexOf("Content-Type:");

    assert.ok(priorityIndex > 0, "Should have X-Priority header");
    assert.ok(sensitivityIndex > 0, "Should have Sensitivity header");
    assert.ok(priorityIndex < sensitivityIndex, "X-Priority should come before Sensitivity");
    assert.ok(sensitivityIndex < contentTypeIndex, "Sensitivity should come before Content-Type");
  });

  it("should work with other message headers", () => {
    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        messageId: "<test@example.com>",
        priority: 2,
        sensitivity: "Personal" as const,
        dispositionNotificationTo: "sender@example.com",
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("Message-ID: <test@example.com>"), "Should have Message-ID");
    assert.ok(eml.includes("X-Priority: 2"), "Should have X-Priority");
    assert.ok(eml.includes("Sensitivity: Personal"), "Should have Sensitivity");
    assert.ok(
      eml.includes("Disposition-Notification-To: sender@example.com"),
      "Should have Disposition-Notification-To",
    );
  });

  it("should exclude Sensitivity from transport headers when we generate it", () => {
    const transportHeaders = `Received: from mail.example.com by mx.example.org\r
Sensitivity: Private\r
From: sender@example.com`;

    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        sensitivity: "Company-Confidential" as const,
        transportMessageHeaders: transportHeaders,
      },
    };

    const eml = convertToEml(parsed);

    // Should use our generated Sensitivity, not the one from transport
    assert.ok(eml.includes("Sensitivity: Company-Confidential"), "Should use generated Sensitivity");

    // Count Sensitivity occurrences - should have exactly one
    const sensitivityMatches = eml.match(/Sensitivity:/g);
    assert.strictEqual(sensitivityMatches?.length, 1, "Should have exactly one Sensitivity header");
  });
});

describe("convertToEml with received-by headers", () => {
  it("should include Delivered-To header when transport headers are not available", () => {
    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [{ name: "", email: "recipient@example.com", type: "to" as const }],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        receivedByEmail: "final-recipient@example.com",
        receivedByName: "Final Recipient",
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("Delivered-To: final-recipient@example.com"), "Should have Delivered-To header");
  });

  it("should not include Delivered-To header when transport headers are present", () => {
    const transportHeaders = `Received: from mail.example.com by mx.example.org with SMTP id abc123\r
From: sender@example.com`;

    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [{ name: "", email: "recipient@example.com", type: "to" as const }],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        receivedByEmail: "final-recipient@example.com",
        receivedByName: "Final Recipient",
        transportMessageHeaders: transportHeaders,
      },
    };

    const eml = convertToEml(parsed);

    // Should have Received header from transport
    assert.ok(eml.includes("Received: from mail.example.com"), "Should include Received header from transport");
    // Should NOT have Delivered-To header since transport headers are present
    assert.ok(!eml.includes("Delivered-To:"), "Should not have Delivered-To when transport headers are present");
  });

  it("should not include Delivered-To header when receivedByEmail is not present", () => {
    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
    };

    const eml = convertToEml(parsed);

    assert.ok(!eml.includes("Delivered-To:"), "Should not have Delivered-To when receivedByEmail is not present");
  });

  it("should place Delivered-To header before Content-Type", () => {
    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        receivedByEmail: "final-recipient@example.com",
      },
    };

    const eml = convertToEml(parsed);

    const deliveredToIndex = eml.indexOf("Delivered-To:");
    const contentTypeIndex = eml.indexOf("Content-Type:");

    assert.ok(deliveredToIndex > 0, "Should have Delivered-To header");
    assert.ok(deliveredToIndex < contentTypeIndex, "Delivered-To should come before Content-Type");
  });

  it("should fold long Delivered-To headers", () => {
    const longEmail =
      "very-long-email-address-that-exceeds-normal-length@very-long-domain-name-that-makes-the-header-exceed-78-characters.example.com";
    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        receivedByEmail: longEmail,
      },
    };

    const eml = convertToEml(parsed);

    // Should have Delivered-To header
    assert.ok(eml.includes("Delivered-To:"), "Should have Delivered-To header");

    // The full email should be present (possibly folded)
    const normalizedEml = eml.replace(/\r\n[\t ]/g, "");
    assert.ok(normalizedEml.includes(longEmail), "Full email address should be present");
  });

  it("should exclude Delivered-To from transport headers when we generate it", () => {
    const transportHeaders = `Received: from mail.example.com by mx.example.org\r
Delivered-To: old-recipient@example.com\r
From: sender@example.com`;

    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        transportMessageHeaders: transportHeaders,
      },
    };

    const eml = convertToEml(parsed);

    // Should include Received header from transport
    assert.ok(eml.includes("Received: from mail.example.com"), "Should include Received header");
    // Should NOT include Delivered-To from transport headers (it's excluded)
    assert.ok(!eml.includes("Delivered-To: old-recipient@example.com"), "Should exclude Delivered-To from transport");
  });

  it("should work alongside other message headers", () => {
    const parsed = {
      subject: "Test",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        messageId: "<test@example.com>",
        priority: 3,
        threadTopic: "Test Thread",
        receivedByEmail: "final-recipient@example.com",
        receivedByName: "Final Recipient",
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("Message-ID: <test@example.com>"), "Should have Message-ID");
    assert.ok(eml.includes("X-Priority: 3"), "Should have X-Priority");
    assert.ok(eml.includes("Thread-Topic: Test Thread"), "Should have Thread-Topic");
    assert.ok(eml.includes("Delivered-To: final-recipient@example.com"), "Should have Delivered-To");
  });
});

describe("convertToEml with mailing list headers", () => {
  it("should include List-Help header when present", () => {
    const parsed = {
      subject: "Mailing List Message",
      from: "list@example.com",
      recipients: [{ name: "", email: "recipient@example.com", type: "to" as const }],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        listHelp: "<mailto:list-help@example.com>",
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("List-Help: <mailto:list-help@example.com>"), "Should have List-Help header");
  });

  it("should include List-Subscribe header when present", () => {
    const parsed = {
      subject: "Mailing List Message",
      from: "list@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        listSubscribe: "<https://example.com/subscribe>",
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("List-Subscribe: <https://example.com/subscribe>"), "Should have List-Subscribe header");
  });

  it("should include List-Unsubscribe header when present", () => {
    const parsed = {
      subject: "Mailing List Message",
      from: "list@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        listUnsubscribe: "<mailto:unsubscribe@example.com>, <https://example.com/unsubscribe>",
      },
    };

    const eml = convertToEml(parsed);

    // Check for List-Unsubscribe header (may be folded)
    assert.ok(eml.includes("List-Unsubscribe:"), "Should have List-Unsubscribe header");
    // Normalize folded headers to check content
    const normalizedEml = eml.replace(/\r\n[\t ]/g, " ");
    assert.ok(
      normalizedEml.includes("<mailto:unsubscribe@example.com>, <https://example.com/unsubscribe>"),
      "Should include unsubscribe URLs",
    );
  });

  it("should include all mailing list headers when present", () => {
    const parsed = {
      subject: "Newsletter",
      from: "newsletter@example.com",
      recipients: [{ name: "", email: "subscriber@example.com", type: "to" as const }],
      date: new Date(),
      body: "Newsletter content",
      attachments: [],
      headers: {
        listHelp: "<mailto:help@example.com>",
        listSubscribe: "<https://example.com/subscribe>",
        listUnsubscribe: "<https://example.com/unsubscribe>",
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("List-Help: <mailto:help@example.com>"), "Should have List-Help header");
    assert.ok(eml.includes("List-Subscribe: <https://example.com/subscribe>"), "Should have List-Subscribe header");
    assert.ok(
      eml.includes("List-Unsubscribe: <https://example.com/unsubscribe>"),
      "Should have List-Unsubscribe header",
    );
  });

  it("should not include mailing list headers when not present", () => {
    const parsed = {
      subject: "Regular Email",
      from: "sender@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
    };

    const eml = convertToEml(parsed);

    assert.ok(!eml.includes("List-Help:"), "Should not have List-Help header when not present");
    assert.ok(!eml.includes("List-Subscribe:"), "Should not have List-Subscribe header when not present");
    assert.ok(!eml.includes("List-Unsubscribe:"), "Should not have List-Unsubscribe header when not present");
  });

  it("should place mailing list headers before Content-Type", () => {
    const parsed = {
      subject: "Test",
      from: "list@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        listHelp: "<mailto:help@example.com>",
        listSubscribe: "<https://example.com/subscribe>",
        listUnsubscribe: "<https://example.com/unsubscribe>",
      },
    };

    const eml = convertToEml(parsed);

    const listHelpIndex = eml.indexOf("List-Help:");
    const listSubscribeIndex = eml.indexOf("List-Subscribe:");
    const listUnsubscribeIndex = eml.indexOf("List-Unsubscribe:");
    const contentTypeIndex = eml.indexOf("Content-Type:");

    assert.ok(listHelpIndex > 0, "Should have List-Help header");
    assert.ok(listSubscribeIndex > 0, "Should have List-Subscribe header");
    assert.ok(listUnsubscribeIndex > 0, "Should have List-Unsubscribe header");
    assert.ok(listHelpIndex < contentTypeIndex, "List-Help should come before Content-Type");
    assert.ok(listSubscribeIndex < contentTypeIndex, "List-Subscribe should come before Content-Type");
    assert.ok(listUnsubscribeIndex < contentTypeIndex, "List-Unsubscribe should come before Content-Type");
  });

  it("should place mailing list headers after threading headers", () => {
    const parsed = {
      subject: "Re: Mailing List Thread",
      from: "list@example.com",
      recipients: [],
      date: new Date(),
      body: "Test body",
      attachments: [],
      headers: {
        threadTopic: "Mailing List Thread",
        listUnsubscribe: "<https://example.com/unsubscribe>",
      },
    };

    const eml = convertToEml(parsed);

    const threadTopicIndex = eml.indexOf("Thread-Topic:");
    const listUnsubscribeIndex = eml.indexOf("List-Unsubscribe:");

    assert.ok(threadTopicIndex > 0, "Should have Thread-Topic header");
    assert.ok(listUnsubscribeIndex > 0, "Should have List-Unsubscribe header");
    assert.ok(threadTopicIndex < listUnsubscribeIndex, "Thread-Topic should come before List-Unsubscribe");
  });

  it("should fold long List-Unsubscribe headers", () => {
    const longUnsubscribe =
      "<mailto:very-long-unsubscribe-address@example.com?subject=unsubscribe>, <https://example.com/unsubscribe?token=very-long-token-that-exceeds-line-length>";

    const parsed = {
      subject: "Newsletter",
      from: "newsletter@example.com",
      recipients: [],
      date: new Date(),
      body: "Newsletter content",
      attachments: [],
      headers: {
        listUnsubscribe: longUnsubscribe,
      },
    };

    const eml = convertToEml(parsed);

    // Should have List-Unsubscribe header
    assert.ok(eml.includes("List-Unsubscribe:"), "Should have List-Unsubscribe header");

    // The full value should be present (possibly folded)
    // Normalize by replacing CRLF+whitespace with single space (how folded headers are joined)
    const normalizedEml = eml.replace(/\r\n[\t ]/g, " ");
    assert.ok(normalizedEml.includes(longUnsubscribe), "Full List-Unsubscribe value should be present");
  });

  it("should exclude mailing list headers from transport headers when we generate them", () => {
    const transportHeaders = `Received: from mail.example.com by mx.example.org\r
List-Help: <mailto:old-help@example.com>\r
List-Subscribe: <https://old.example.com/subscribe>\r
List-Unsubscribe: <https://old.example.com/unsubscribe>\r
From: list@example.com`;

    const parsed = {
      subject: "Newsletter",
      from: "list@example.com",
      recipients: [],
      date: new Date(),
      body: "Newsletter content",
      attachments: [],
      headers: {
        listHelp: "<mailto:new-help@example.com>",
        listSubscribe: "<https://new.example.com/subscribe>",
        listUnsubscribe: "<https://new.example.com/unsubscribe>",
        transportMessageHeaders: transportHeaders,
      },
    };

    const eml = convertToEml(parsed);

    // Should use our generated headers, not the ones from transport
    assert.ok(eml.includes("List-Help: <mailto:new-help@example.com>"), "Should use generated List-Help");
    assert.ok(
      eml.includes("List-Subscribe: <https://new.example.com/subscribe>"),
      "Should use generated List-Subscribe",
    );
    assert.ok(
      eml.includes("List-Unsubscribe: <https://new.example.com/unsubscribe>"),
      "Should use generated List-Unsubscribe",
    );

    // Count occurrences - should have exactly one of each
    const listHelpMatches = eml.match(/List-Help:/g);
    const listSubscribeMatches = eml.match(/List-Subscribe:/g);
    const listUnsubscribeMatches = eml.match(/List-Unsubscribe:/g);
    assert.strictEqual(listHelpMatches?.length, 1, "Should have exactly one List-Help header");
    assert.strictEqual(listSubscribeMatches?.length, 1, "Should have exactly one List-Subscribe header");
    assert.strictEqual(listUnsubscribeMatches?.length, 1, "Should have exactly one List-Unsubscribe header");
  });

  it("should work with other message headers", () => {
    const parsed = {
      subject: "Newsletter",
      from: "newsletter@example.com",
      recipients: [],
      date: new Date(),
      body: "Newsletter content",
      attachments: [],
      headers: {
        messageId: "<newsletter123@example.com>",
        priority: 3,
        threadTopic: "Newsletter Topic",
        listHelp: "<mailto:help@example.com>",
        listUnsubscribe: "<https://example.com/unsubscribe>",
      },
    };

    const eml = convertToEml(parsed);

    assert.ok(eml.includes("Message-ID: <newsletter123@example.com>"), "Should have Message-ID");
    assert.ok(eml.includes("X-Priority: 3"), "Should have X-Priority");
    assert.ok(eml.includes("Thread-Topic: Newsletter Topic"), "Should have Thread-Topic");
    assert.ok(eml.includes("List-Help: <mailto:help@example.com>"), "Should have List-Help");
    assert.ok(eml.includes("List-Unsubscribe: <https://example.com/unsubscribe>"), "Should have List-Unsubscribe");
  });
});
