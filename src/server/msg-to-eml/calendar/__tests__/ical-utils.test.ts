import assert from "node:assert";
import { describe, it } from "node:test";
import { escapeICalText, foldICalLine, formatICalDateTime } from "../ical-utils.js";

describe("formatICalDateTime", () => {
  it("should format date in iCalendar format", () => {
    const date = new Date("2024-03-15T14:30:45Z");
    const result = formatICalDateTime(date);
    assert.strictEqual(result, "20240315T143045Z");
  });

  it("should handle midnight", () => {
    const date = new Date("2024-01-01T00:00:00Z");
    const result = formatICalDateTime(date);
    assert.strictEqual(result, "20240101T000000Z");
  });

  it("should pad single-digit values", () => {
    const date = new Date("2024-01-05T09:05:01Z");
    const result = formatICalDateTime(date);
    assert.strictEqual(result, "20240105T090501Z");
  });
});

describe("escapeICalText", () => {
  it("should escape backslashes", () => {
    assert.strictEqual(escapeICalText("path\\to\\file"), "path\\\\to\\\\file");
  });

  it("should escape semicolons", () => {
    assert.strictEqual(escapeICalText("first;second"), "first\\;second");
  });

  it("should escape commas", () => {
    assert.strictEqual(escapeICalText("one, two"), "one\\, two");
  });

  it("should escape newlines", () => {
    assert.strictEqual(escapeICalText("line1\nline2"), "line1\\nline2");
    assert.strictEqual(escapeICalText("line1\r\nline2"), "line1\\nline2");
  });
});

describe("foldICalLine", () => {
  it("should not fold short lines", () => {
    const shortLine = "SUMMARY:Short meeting";
    const result = foldICalLine(shortLine);
    assert.strictEqual(result, shortLine);
  });

  it("should fold long lines at 75 characters", () => {
    const longLine =
      "DESCRIPTION:This is a very long description that definitely exceeds the 75 character limit for iCalendar lines";
    const result = foldICalLine(longLine);

    const lines = result.split("\r\n");
    assert.ok(lines.length > 1, "Should be folded into multiple lines");
    assert.strictEqual(lines[0].length, 75, "First line should be 75 chars");

    // Continuation lines should start with space
    for (let i = 1; i < lines.length; i++) {
      assert.ok(lines[i].startsWith(" "), "Continuation lines should start with space");
    }
  });
});
