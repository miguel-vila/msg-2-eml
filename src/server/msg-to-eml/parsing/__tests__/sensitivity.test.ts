import assert from "node:assert";
import { describe, it } from "node:test";
import { mapSensitivity } from "../msg.js";

describe("mapSensitivity", () => {
  it("should return undefined for 0 (Normal)", () => {
    const result = mapSensitivity(0);
    assert.strictEqual(result, undefined);
  });

  it("should return 'Personal' for 1", () => {
    const result = mapSensitivity(1);
    assert.strictEqual(result, "Personal");
  });

  it("should return 'Private' for 2", () => {
    const result = mapSensitivity(2);
    assert.strictEqual(result, "Private");
  });

  it("should return 'Company-Confidential' for 3", () => {
    const result = mapSensitivity(3);
    assert.strictEqual(result, "Company-Confidential");
  });

  it("should return undefined for undefined input", () => {
    const result = mapSensitivity(undefined);
    assert.strictEqual(result, undefined);
  });

  it("should return undefined for unknown values", () => {
    const result = mapSensitivity(999);
    assert.strictEqual(result, undefined);
  });

  it("should return undefined for negative values", () => {
    const result = mapSensitivity(-1);
    assert.strictEqual(result, undefined);
  });
});
