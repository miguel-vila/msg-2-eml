/**
 * Formats a Date as an iCalendar date-time string in UTC.
 * Format: YYYYMMDDTHHMMSSZ
 */
export function formatICalDateTime(date: Date): string {
  const year = date.getUTCFullYear();
  const month = (date.getUTCMonth() + 1).toString().padStart(2, "0");
  const day = date.getUTCDate().toString().padStart(2, "0");
  const hours = date.getUTCHours().toString().padStart(2, "0");
  const minutes = date.getUTCMinutes().toString().padStart(2, "0");
  const seconds = date.getUTCSeconds().toString().padStart(2, "0");
  return `${year}${month}${day}T${hours}${minutes}${seconds}Z`;
}

/**
 * Generates a unique identifier for an iCalendar event.
 * Uses the current timestamp and a random component.
 */
export function generateUID(): string {
  const timestamp = Date.now();
  const random = Math.random().toString(36).substring(2, 10);
  return `${timestamp}-${random}@msg-to-eml`;
}

/**
 * Escapes special characters in iCalendar text values.
 * Backslash, semicolon, comma, and newlines need escaping.
 */
export function escapeICalText(text: string): string {
  return text
    .replace(/\\/g, "\\\\")
    .replace(/;/g, "\\;")
    .replace(/,/g, "\\,")
    .replace(/\r\n/g, "\\n")
    .replace(/\n/g, "\\n")
    .replace(/\r/g, "\\n");
}

/**
 * Folds iCalendar content lines at 75 octets as per RFC 5545.
 * Continuation lines start with a single space.
 */
export function foldICalLine(line: string): string {
  if (line.length <= 75) {
    return line;
  }

  const result: string[] = [];
  let remaining = line;

  // First line can be up to 75 chars
  result.push(remaining.substring(0, 75));
  remaining = remaining.substring(75);

  // Subsequent lines are prefixed with space, so content is up to 74 chars
  while (remaining.length > 0) {
    result.push(` ${remaining.substring(0, 74)}`);
    remaining = remaining.substring(74);
  }

  return result.join("\r\n");
}
