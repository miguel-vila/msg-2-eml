import assert from "node:assert";
import { describe, it } from "node:test";
import { generateVCalendar, isCalendarMessage, parseAttendeeString } from "../vcalendar.js";

describe("parseAttendeeString", () => {
  it("should parse semicolon-separated attendees", () => {
    const result = parseAttendeeString("alice@example.com; bob@example.com; charlie@example.com");
    assert.deepStrictEqual(result, ["alice@example.com", "bob@example.com", "charlie@example.com"]);
  });

  it("should trim whitespace from attendees", () => {
    const result = parseAttendeeString("  alice@example.com  ;  bob@example.com  ");
    assert.deepStrictEqual(result, ["alice@example.com", "bob@example.com"]);
  });

  it("should return empty array for undefined input", () => {
    const result = parseAttendeeString(undefined);
    assert.deepStrictEqual(result, []);
  });

  it("should filter out empty strings", () => {
    const result = parseAttendeeString("alice@example.com;;bob@example.com;");
    assert.deepStrictEqual(result, ["alice@example.com", "bob@example.com"]);
  });
});

describe("isCalendarMessage", () => {
  it("should return true for IPM.Appointment", () => {
    assert.strictEqual(isCalendarMessage("IPM.Appointment"), true);
  });

  it("should return true for lowercase ipm.appointment", () => {
    assert.strictEqual(isCalendarMessage("ipm.appointment"), true);
  });

  it("should return true for subclasses of IPM.Appointment", () => {
    assert.strictEqual(isCalendarMessage("IPM.Appointment.Occurrence"), true);
  });

  it("should return false for regular email", () => {
    assert.strictEqual(isCalendarMessage("IPM.Note"), false);
  });

  it("should return false for undefined", () => {
    assert.strictEqual(isCalendarMessage(undefined), false);
  });
});

describe("generateVCalendar", () => {
  it("should generate valid VCALENDAR structure", () => {
    const event = {
      startTime: new Date("2024-03-15T14:00:00Z"),
      endTime: new Date("2024-03-15T15:00:00Z"),
      location: "Conference Room A",
      organizer: "organizer@example.com",
      attendees: ["attendee1@example.com", "attendee2@example.com"],
    };

    const vcal = generateVCalendar(event, "Team Meeting", "Quarterly planning session");

    assert.ok(vcal.includes("BEGIN:VCALENDAR"), "Should have VCALENDAR");
    assert.ok(vcal.includes("END:VCALENDAR"), "Should end VCALENDAR");
    assert.ok(vcal.includes("VERSION:2.0"), "Should have version");
    assert.ok(vcal.includes("PRODID:"), "Should have product ID");
    assert.ok(vcal.includes("METHOD:REQUEST"), "Should have method");
    assert.ok(vcal.includes("BEGIN:VEVENT"), "Should have VEVENT");
    assert.ok(vcal.includes("END:VEVENT"), "Should end VEVENT");
    assert.ok(vcal.includes("DTSTART:20240315T140000Z"), "Should have start time");
    assert.ok(vcal.includes("DTEND:20240315T150000Z"), "Should have end time");
    assert.ok(vcal.includes("SUMMARY:Team Meeting"), "Should have summary");
    assert.ok(vcal.includes("LOCATION:Conference Room A"), "Should have location");
    assert.ok(vcal.includes("ORGANIZER:mailto:organizer@example.com"), "Should have organizer");
    assert.ok(vcal.includes("ATTENDEE:mailto:attendee1@example.com"), "Should have attendee 1");
    assert.ok(vcal.includes("ATTENDEE:mailto:attendee2@example.com"), "Should have attendee 2");
  });

  it("should escape special characters in text", () => {
    const event = {
      startTime: new Date("2024-03-15T14:00:00Z"),
      endTime: new Date("2024-03-15T15:00:00Z"),
      location: "Room A; Building 1",
      attendees: [],
    };

    const vcal = generateVCalendar(event, "Meeting, Important!", "Line1\nLine2");

    assert.ok(vcal.includes("SUMMARY:Meeting\\, Important!"), "Should escape comma in summary");
    assert.ok(vcal.includes("LOCATION:Room A\\; Building 1"), "Should escape semicolon in location");
    assert.ok(vcal.includes("DESCRIPTION:Line1\\nLine2"), "Should escape newline in description");
  });

  it("should handle event without optional fields", () => {
    const event = {
      startTime: new Date("2024-03-15T14:00:00Z"),
      endTime: new Date("2024-03-15T15:00:00Z"),
      attendees: [],
    };

    const vcal = generateVCalendar(event, "Simple Meeting", "");

    assert.ok(!vcal.includes("LOCATION:"), "Should not have location");
    assert.ok(!vcal.includes("ORGANIZER:"), "Should not have organizer");
    assert.ok(!vcal.includes("ATTENDEE:"), "Should not have attendees");
    assert.ok(!vcal.includes("DESCRIPTION:"), "Should not have empty description");
  });
});
