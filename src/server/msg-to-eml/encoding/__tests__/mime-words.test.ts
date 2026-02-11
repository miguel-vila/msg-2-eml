import assert from "node:assert";
import { describe, it } from "node:test";
import { encodeDisplayName, encodeRfc2047 } from "../mime-words.js";

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
