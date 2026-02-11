import assert from "node:assert";
import { describe, it } from "node:test";
import { foldHeader } from "../headers.js";

describe("foldHeader", () => {
  it("should not fold headers shorter than 78 characters", () => {
    const result = foldHeader("Subject", "Short subject line");
    assert.strictEqual(result, "Subject: Short subject line");
    assert.ok(!result.includes("\r\n"));
  });

  it("should fold headers longer than 78 characters", () => {
    const longSubject =
      "This is a very long subject line that exceeds the 78 character limit and needs to be folded properly";
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
    const recipients =
      '"Alice Smith" <alice@example.com>, "Bob Johnson" <bob@example.com>, "Charlie Brown" <charlie@example.com>';
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
