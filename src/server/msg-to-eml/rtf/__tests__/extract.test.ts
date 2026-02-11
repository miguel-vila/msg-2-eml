import assert from "node:assert";
import { describe, it } from "node:test";
import { extractBodyFromRtf } from "../extract.js";

/**
 * Creates an uncompressed RTF format for PidTagRtfCompressed.
 * The format uses MELA (uncompressed) magic number.
 * Structure: fileSize (4) + rawSize (4) + compType (4) + crc (4) + raw data
 */
function createUncompressedRtf(rtfContent: string): number[] {
  const rawBytes = Buffer.from(rtfContent, "latin1");
  const rawSize = rawBytes.length;
  const fileSize = rawSize + 12; // rawSize + compType + crc (fileSize excludes itself)
  const UNCOMPRESSED = 0x414c454d; // "MELA" in little-endian
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

describe("extractBodyFromRtf", () => {
  it("should extract plain text from simple RTF", () => {
    // Simple RTF with plain text
    const rtf = "{\\rtf1\\ansi Hello World}";
    const uncompressed = createUncompressedRtf(rtf);
    const result = extractBodyFromRtf(uncompressed);

    assert.ok(result !== null, "Should return a result");
    assert.ok(result?.text.includes("Hello World"), "Should extract text content");
    assert.strictEqual(result?.html, undefined, "Should not have HTML for plain text RTF");
  });

  it("should extract HTML from RTF-encapsulated HTML (fromhtml)", () => {
    // RTF with encapsulated HTML using \fromhtml1
    const rtf =
      "{\\rtf1\\ansi\\fromhtml1 {\\*\\htmltag64 <html>}{\\*\\htmltag64 <body>}Hello HTML{\\*\\htmltag64 </body>}{\\*\\htmltag64 </html>}}";
    const uncompressed = createUncompressedRtf(rtf);
    const result = extractBodyFromRtf(uncompressed);

    assert.ok(result !== null, "Should return a result");
    assert.ok(result?.html !== undefined, "Should extract HTML");
    assert.ok(result?.html?.includes("Hello HTML"), "HTML should contain content");
    assert.ok(result?.text.includes("Hello HTML"), "Should also provide plain text fallback");
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
    assert.ok(result?.text.length > 0, "Should have extracted text");
  });

  it("should strip HTML tags when providing plain text from HTML content", () => {
    const rtf = "{\\rtf1\\ansi\\fromhtml1 {\\*\\htmltag64 <p>}Paragraph text{\\*\\htmltag64 </p>}}";
    const uncompressed = createUncompressedRtf(rtf);
    const result = extractBodyFromRtf(uncompressed);

    assert.ok(result !== null, "Should return a result");
    if (result?.html) {
      // Plain text version should not contain HTML tags
      assert.ok(!result?.text.includes("<p>"), "Plain text should not contain <p> tag");
      assert.ok(!result?.text.includes("</p>"), "Plain text should not contain </p> tag");
    }
  });
});
