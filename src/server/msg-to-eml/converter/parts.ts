import { generateVCalendar } from "../calendar/index.js";
import { encodeBase64, encodeQuotedPrintable, formatFilenameParams } from "../encoding/index.js";
import type { Attachment, CalendarEvent, ParsedMsg } from "../types/index.js";

export function generateTextPart(body: string): string {
  let part = "";
  part += `Content-Type: text/plain; charset="utf-8"\r\n`;
  part += `Content-Transfer-Encoding: quoted-printable\r\n`;
  part += `\r\n`;
  part += encodeQuotedPrintable(body);
  return part;
}

export function generateHtmlPart(bodyHtml: string): string {
  let part = "";
  part += `Content-Type: text/html; charset="utf-8"\r\n`;
  part += `Content-Transfer-Encoding: quoted-printable\r\n`;
  part += `\r\n`;
  part += encodeQuotedPrintable(bodyHtml);
  return part;
}

export function generateCalendarPart(calendarEvent: CalendarEvent, subject: string, body: string): string {
  const vcalendar = generateVCalendar(calendarEvent, subject, body);
  let part = "";
  part += `Content-Type: text/calendar; charset="utf-8"; method=REQUEST\r\n`;
  part += `Content-Transfer-Encoding: 7bit\r\n`;
  part += `\r\n`;
  part += vcalendar;
  return part;
}

export function generateAlternativePart(
  boundary: string,
  parsed: ParsedMsg,
  hasHtml: boolean,
  hasCalendar: boolean,
): string {
  let part = "";
  part += `Content-Type: multipart/alternative; boundary="${boundary}"\r\n`;
  part += `\r\n`;
  part += `--${boundary}\r\n`;
  part += generateTextPart(parsed.body);
  part += `\r\n`;
  if (hasHtml && parsed.bodyHtml) {
    part += `--${boundary}\r\n`;
    part += generateHtmlPart(parsed.bodyHtml);
    part += `\r\n`;
  }
  if (hasCalendar && parsed.calendarEvent) {
    part += `--${boundary}\r\n`;
    part += generateCalendarPart(parsed.calendarEvent, parsed.subject, parsed.body);
    part += `\r\n`;
  }
  part += `--${boundary}--\r\n`;
  return part;
}

export function generateAttachmentPart(att: Attachment, inline: boolean): string {
  const filenameParams = formatFilenameParams(att.fileName);
  let part = "";
  part += `Content-Type: ${att.contentType}; ${filenameParams.name}\r\n`;
  if (inline && att.contentId) {
    part += `Content-ID: <${att.contentId}>\r\n`;
    part += `Content-Disposition: inline; ${filenameParams.disposition}\r\n`;
  } else {
    part += `Content-Disposition: attachment; ${filenameParams.disposition}\r\n`;
  }

  // For message/rfc822 (embedded emails), use 7bit encoding since the content
  // is already a valid EML file with its own encoding
  if (att.contentType === "message/rfc822") {
    part += `Content-Transfer-Encoding: 7bit\r\n`;
    part += `\r\n`;
    part += new TextDecoder().decode(att.content);
    // Ensure proper line ending
    if (!part.endsWith("\r\n")) {
      part += "\r\n";
    }
  } else {
    part += `Content-Transfer-Encoding: base64\r\n`;
    part += `\r\n`;
    const base64 = encodeBase64(att.content);
    for (let i = 0; i < base64.length; i += 76) {
      part += `${base64.slice(i, i + 76)}\r\n`;
    }
  }
  return part;
}
