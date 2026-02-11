import assert from "node:assert";
import { describe, it } from "node:test";
import { getRecipientType, mapToXPriority } from "../priority.js";

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

describe("getRecipientType", () => {
  it("should return 'to' for MAPI_TO (1)", () => {
    assert.strictEqual(getRecipientType(1), "to");
  });

  it("should return 'cc' for MAPI_CC (2)", () => {
    assert.strictEqual(getRecipientType(2), "cc");
  });

  it("should return 'bcc' for MAPI_BCC (3)", () => {
    assert.strictEqual(getRecipientType(3), "bcc");
  });

  it("should return 'to' for undefined", () => {
    assert.strictEqual(getRecipientType(undefined), "to");
  });

  it("should return 'to' for unknown type", () => {
    assert.strictEqual(getRecipientType(0), "to");
    assert.strictEqual(getRecipientType(99), "to");
  });
});
