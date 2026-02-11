import type { CalendarEvent } from "../types/index.js";
import { escapeICalText, foldICalLine, formatICalDateTime, generateUID } from "./ical-utils.js";

/**
 * Parses an attendee string (semicolon-separated list) into individual email addresses.
 */
export function parseAttendeeString(attendeesStr: string | undefined): string[] {
  if (!attendeesStr) return [];
  return attendeesStr
    .split(";")
    .map((a) => a.trim())
    .filter((a) => a.length > 0);
}

/**
 * Checks if a message class indicates a calendar appointment.
 */
export function isCalendarMessage(messageClass: string | undefined): boolean {
  if (!messageClass) return false;
  const normalized = messageClass.toLowerCase();
  return normalized === "ipm.appointment" || normalized.startsWith("ipm.appointment.");
}

/**
 * Generates a VCALENDAR string with a VEVENT for the given calendar event.
 */
export function generateVCalendar(event: CalendarEvent, subject: string, description: string): string {
  const lines: string[] = [];

  lines.push("BEGIN:VCALENDAR");
  lines.push("VERSION:2.0");
  lines.push("PRODID:-//msg-to-eml//NONSGML v1.0//EN");
  lines.push("METHOD:REQUEST");

  lines.push("BEGIN:VEVENT");
  lines.push(foldICalLine(`UID:${generateUID()}`));
  lines.push(`DTSTAMP:${formatICalDateTime(new Date())}`);
  lines.push(`DTSTART:${formatICalDateTime(event.startTime)}`);
  lines.push(`DTEND:${formatICalDateTime(event.endTime)}`);
  lines.push(foldICalLine(`SUMMARY:${escapeICalText(subject)}`));

  if (description) {
    lines.push(foldICalLine(`DESCRIPTION:${escapeICalText(description)}`));
  }

  if (event.location) {
    lines.push(foldICalLine(`LOCATION:${escapeICalText(event.location)}`));
  }

  if (event.organizer) {
    // Format organizer - if it looks like an email, use MAILTO
    if (event.organizer.includes("@")) {
      lines.push(foldICalLine(`ORGANIZER:mailto:${event.organizer}`));
    } else {
      lines.push(foldICalLine(`ORGANIZER;CN=${escapeICalText(event.organizer)}:mailto:noreply@unknown`));
    }
  }

  // Add attendees
  for (const attendee of event.attendees) {
    if (attendee.includes("@")) {
      lines.push(foldICalLine(`ATTENDEE:mailto:${attendee}`));
    } else {
      lines.push(foldICalLine(`ATTENDEE;CN=${escapeICalText(attendee)}:mailto:noreply@unknown`));
    }
  }

  lines.push("END:VEVENT");
  lines.push("END:VCALENDAR");

  return lines.join("\r\n");
}
