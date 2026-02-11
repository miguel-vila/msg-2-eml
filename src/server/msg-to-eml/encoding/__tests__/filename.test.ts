import assert from "node:assert";
import { describe, it } from "node:test";
import { encodeRfc2231, formatFilenameParams } from "../filename.js";

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

  it("should return simple quoted format for filenames with spaces", () => {
    const result = formatFilenameParams("my document.pdf");
    // ASCII with spaces should still use simple format
    assert.strictEqual(result.name, 'name="my document.pdf"');
    assert.strictEqual(result.disposition, 'filename="my document.pdf"');
  });
});
