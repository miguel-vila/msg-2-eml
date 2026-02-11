import { describe, it, before, after } from "node:test";
import assert from "node:assert";
import { readFileSync, existsSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { msgToEml, parseMsg, convertToEml } from "./msg-to-eml.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const fixturesDir = join(__dirname, "..", "..", "test", "fixtures");

/**
 * Helper to load a fixture MSG file
 */
function loadFixture(filename: string): ArrayBuffer {
  const filePath = join(fixturesDir, filename);
  if (!existsSync(filePath)) {
    throw new Error(`Fixture file not found: ${filePath}`);
  }
  const buffer = readFileSync(filePath);
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

/**
 * Validates that the EML output is a valid RFC 5322 message
 */
function validateEmlStructure(eml: string): void {
  // Must have CRLF line endings
  assert.ok(eml.includes("\r\n"), "EML should use CRLF line endings");

  // Must have required headers
  assert.ok(eml.includes("From:"), "EML must have From header");
  assert.ok(eml.includes("Date:"), "EML must have Date header");
  assert.ok(eml.includes("MIME-Version: 1.0"), "EML must have MIME-Version header");

  // Headers and body must be separated by blank line
  assert.ok(eml.includes("\r\n\r\n"), "Headers and body must be separated by blank line");
}

/**
 * Checks that no header line exceeds RFC 5322 recommended length
 */
function validateHeaderFolding(eml: string): void {
  const headerSection = eml.split("\r\n\r\n")[0];
  const lines = headerSection.split("\r\n");

  // Unfold headers to check logical header lines
  const unfoldedHeaders: string[] = [];
  for (const line of lines) {
    if (line.startsWith(" ") || line.startsWith("\t")) {
      // Continuation line
      if (unfoldedHeaders.length > 0) {
        unfoldedHeaders[unfoldedHeaders.length - 1] += line;
      }
    } else {
      unfoldedHeaders.push(line);
    }
  }

  // Physical lines should not exceed 998 characters (MUST per RFC 5322)
  for (const line of lines) {
    assert.ok(
      line.length <= 998,
      `Line exceeds 998 character limit: ${line.substring(0, 50)}...`
    );
  }
}

/**
 * Validates MIME multipart structure
 */
function validateMultipartStructure(eml: string, boundary: string): void {
  // Check boundary markers exist
  assert.ok(eml.includes(`--${boundary}`), "Missing boundary markers");
  assert.ok(eml.includes(`--${boundary}--`), "Missing closing boundary");

  // Count parts
  const parts = eml.split(`--${boundary}`).slice(1, -1);
  assert.ok(parts.length > 0, "Multipart message should have at least one part");
}

describe("Integration Tests", () => {
  describe("Full MSG to EML Conversion", () => {
    it("should convert a real MSG file to valid EML", () => {
      const msgBuffer = loadFixture("simple-email.msg");
      const eml = msgToEml(msgBuffer);

      validateEmlStructure(eml);
      validateHeaderFolding(eml);

      // Should have actual content
      assert.ok(eml.length > 100, "EML should have substantial content");
    });

    it("should preserve email metadata through conversion", () => {
      const msgBuffer = loadFixture("simple-email.msg");
      const parsed = parseMsg(msgBuffer);

      // Verify parsing extracted data
      assert.ok(parsed.subject !== undefined, "Should extract subject");
      assert.ok(parsed.from !== undefined, "Should extract sender");
      assert.ok(parsed.date instanceof Date, "Should extract date");

      // Convert and verify data is preserved
      const eml = convertToEml(parsed);
      if (parsed.subject && parsed.subject.length > 0) {
        // Subject may be encoded, but should be present
        assert.ok(
          eml.includes("Subject:"),
          "Converted EML should have Subject header"
        );
      }
    });

    it("should handle the full pipeline without errors", () => {
      const msgBuffer = loadFixture("simple-email.msg");

      // This should not throw
      const eml = msgToEml(msgBuffer);

      // Verify it returns a string
      assert.strictEqual(typeof eml, "string");
      assert.ok(eml.length > 0, "Should return non-empty EML");
    });
  });

  describe("EML Output Compliance", () => {
    it("should produce valid Content-Type headers", () => {
      const msgBuffer = loadFixture("simple-email.msg");
      const eml = msgToEml(msgBuffer);

      // Extract Content-Type header
      const contentTypeMatch = eml.match(/Content-Type:\s*([^\r\n]+(?:\r\n[ \t]+[^\r\n]+)*)/);
      assert.ok(contentTypeMatch, "Should have Content-Type header");

      const contentType = contentTypeMatch[1].replace(/\r\n[ \t]+/g, " ");
      // Should be a valid MIME type
      assert.ok(
        contentType.includes("text/") || contentType.includes("multipart/"),
        `Content-Type should be text/* or multipart/*: ${contentType}`
      );
    });

    it("should have properly formatted Date header", () => {
      const msgBuffer = loadFixture("simple-email.msg");
      const eml = msgToEml(msgBuffer);

      const dateMatch = eml.match(/Date:\s*([^\r\n]+)/);
      assert.ok(dateMatch, "Should have Date header");

      const dateValue = dateMatch[1];
      // RFC 5322 date format: day-name, day month year hour:minute:second zone
      // Example: Mon, 15 Jan 2024 10:30:00 +0000
      assert.ok(
        /[A-Z][a-z]{2},\s+\d{1,2}\s+[A-Z][a-z]{2}\s+\d{4}\s+\d{2}:\d{2}:\d{2}\s+[+-]\d{4}/.test(
          dateValue
        ),
        `Date should be in RFC 5322 format: ${dateValue}`
      );
    });

    it("should encode non-ASCII content properly if present", () => {
      const msgBuffer = loadFixture("simple-email.msg");
      const parsed = parseMsg(msgBuffer);

      // Create a test case with non-ASCII content
      const testParsed = {
        ...parsed,
        subject: "Tëst Subjëct with spëcial chäräctërs",
        recipients: [
          {
            name: "Récipiént Nàme",
            email: "test@example.com",
            type: "to" as const,
          },
        ],
      };

      const eml = convertToEml(testParsed);

      // Should use RFC 2047 encoding for non-ASCII
      assert.ok(
        eml.includes("=?UTF-8?") || eml.includes("=?utf-8?"),
        "Non-ASCII content should be RFC 2047 encoded"
      );
    });
  });

  describe("Attachment Handling", () => {
    it("should properly encode attachments as base64", () => {
      const msgBuffer = loadFixture("simple-email.msg");
      const parsed = parseMsg(msgBuffer);

      if (parsed.attachments.length > 0) {
        const eml = convertToEml(parsed);

        // If there are attachments, should have multipart structure
        assert.ok(
          eml.includes("multipart/"),
          "Messages with attachments should be multipart"
        );

        // Attachments should be base64 encoded
        assert.ok(
          eml.includes("Content-Transfer-Encoding: base64"),
          "Attachments should be base64 encoded"
        );
      } else {
        // No attachments - just verify the test runs
        assert.ok(true, "No attachments in this test file");
      }
    });

    it("should handle inline attachments with Content-ID", () => {
      const msgBuffer = loadFixture("simple-email.msg");
      const parsed = parseMsg(msgBuffer);

      // Add a test inline attachment
      const testParsed = {
        ...parsed,
        bodyHtml: '<img src="cid:image001@test.local" />',
        attachments: [
          {
            fileName: "image.png",
            content: new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
            contentType: "image/png",
            contentId: "image001@test.local",
          },
        ],
      };

      const eml = convertToEml(testParsed);

      // Should have Content-ID header for inline attachment
      assert.ok(
        eml.includes("Content-ID: <image001@test.local>"),
        "Inline attachment should have Content-ID"
      );

      // Should use multipart/related for inline images
      assert.ok(
        eml.includes("multipart/related"),
        "Inline images should use multipart/related"
      );
    });
  });

  describe("Header Extraction", () => {
    it("should extract and include Message-ID if present", () => {
      const msgBuffer = loadFixture("simple-email.msg");
      const parsed = parseMsg(msgBuffer);

      if (parsed.headers?.messageId) {
        const eml = convertToEml(parsed);
        assert.ok(
          eml.includes("Message-ID:") || eml.includes("Message-Id:"),
          "Should include Message-ID header"
        );
      } else {
        // Not all MSG files have Message-ID
        assert.ok(true, "No Message-ID in source file");
      }
    });

    it("should map priority to X-Priority header", () => {
      const msgBuffer = loadFixture("simple-email.msg");
      const parsed = parseMsg(msgBuffer);

      // Add test priority
      const testParsed = {
        ...parsed,
        headers: {
          ...parsed.headers,
          priority: 1, // High priority
        },
      };

      const eml = convertToEml(testParsed);
      assert.ok(
        eml.includes("X-Priority: 1"),
        "Should include X-Priority header"
      );
    });
  });

  describe("Special Content Types", () => {
    it("should handle calendar events with VCALENDAR output", () => {
      const msgBuffer = loadFixture("simple-email.msg");
      const parsed = parseMsg(msgBuffer);

      // Add test calendar event
      const testParsed = {
        ...parsed,
        calendarEvent: {
          startTime: new Date("2024-03-15T09:00:00Z"),
          endTime: new Date("2024-03-15T10:00:00Z"),
          location: "Conference Room A",
          organizer: "organizer@example.com",
          attendees: ["attendee1@example.com", "attendee2@example.com"],
        },
      };

      const eml = convertToEml(testParsed);

      // Should have text/calendar content type
      assert.ok(
        eml.includes("text/calendar"),
        "Calendar events should have text/calendar part"
      );

      // Should have VCALENDAR structure
      assert.ok(
        eml.includes("BEGIN:VCALENDAR"),
        "Should include VCALENDAR start"
      );
      assert.ok(eml.includes("BEGIN:VEVENT"), "Should include VEVENT start");
      assert.ok(eml.includes("END:VEVENT"), "Should include VEVENT end");
      assert.ok(eml.includes("END:VCALENDAR"), "Should include VCALENDAR end");
    });

    it("should handle embedded MSG files as message/rfc822", () => {
      const msgBuffer = loadFixture("simple-email.msg");
      const parsed = parseMsg(msgBuffer);

      // Add test embedded message
      const testParsed = {
        ...parsed,
        attachments: [
          {
            fileName: "forwarded.eml",
            content: new TextEncoder().encode(
              "From: test@example.com\r\nSubject: Embedded\r\n\r\nBody"
            ),
            contentType: "message/rfc822",
            isEmbeddedMessage: true,
          },
        ],
      };

      const eml = convertToEml(testParsed);

      // Should have message/rfc822 content type
      assert.ok(
        eml.includes("message/rfc822"),
        "Embedded messages should have message/rfc822 type"
      );
    });

    it("should handle RTF body extraction", () => {
      const msgBuffer = loadFixture("simple-email.msg");
      const parsed = parseMsg(msgBuffer);

      // The parsed message should have a body
      // (may come from plain text, HTML, or RTF)
      assert.ok(
        parsed.body !== undefined,
        "Should extract body from MSG file"
      );
    });
  });

  describe("Receipt Headers", () => {
    it("should include Disposition-Notification-To for read receipts", () => {
      const msgBuffer = loadFixture("simple-email.msg");
      const parsed = parseMsg(msgBuffer);

      const testParsed = {
        ...parsed,
        headers: {
          ...parsed.headers,
          dispositionNotificationTo: "sender@example.com",
        },
      };

      const eml = convertToEml(testParsed);
      assert.ok(
        eml.includes("Disposition-Notification-To:"),
        "Should include Disposition-Notification-To header"
      );
    });

    it("should include Return-Receipt-To for delivery receipts", () => {
      const msgBuffer = loadFixture("simple-email.msg");
      const parsed = parseMsg(msgBuffer);

      const testParsed = {
        ...parsed,
        headers: {
          ...parsed.headers,
          returnReceiptTo: "sender@example.com",
        },
      };

      const eml = convertToEml(testParsed);
      assert.ok(
        eml.includes("Return-Receipt-To:"),
        "Should include Return-Receipt-To header"
      );
    });
  });
});

describe("Error Handling", () => {
  it("should throw on invalid MSG data", () => {
    const invalidBuffer = new ArrayBuffer(10);
    new Uint8Array(invalidBuffer).fill(0);

    assert.throws(() => {
      msgToEml(invalidBuffer);
    }, "Should throw on invalid MSG data");
  });

  it("should throw on empty buffer", () => {
    const emptyBuffer = new ArrayBuffer(0);

    assert.throws(() => {
      msgToEml(emptyBuffer);
    }, "Should throw on empty buffer");
  });
});
