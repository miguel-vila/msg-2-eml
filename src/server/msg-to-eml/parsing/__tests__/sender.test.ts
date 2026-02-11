import assert from "node:assert";
import { describe, it } from "node:test";
import { formatSender } from "../sender.js";

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
