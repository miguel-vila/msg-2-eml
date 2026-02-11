import {
  Msg,
  Attachment as MsgAttachment,
  Recipient as MsgRecipient,
  PidTagSubject,
  PidTagBody,
  PidTagBodyHtml,
  PidTagSenderEmailAddress,
  PidTagSenderSmtpAddress,
  PidTagSenderName,
  PidTagSentRepresentingEmailAddress,
  PidTagSentRepresentingSmtpAddress,
  PidTagSentRepresentingName,
  PidTagMessageDeliveryTime,
  PidTagDisplayName,
  PidTagEmailAddress,
  PidTagRecipientType,
  PidTagAttachLongFilename,
  PidTagAttachFilename,
  PidTagAttachMimeTag,
  PidTagInternetMessageId,
  PidTagInReplyToId,
  PidTagInternetReferences,
  PidTagReplyRecipientNames,
  PidTagPriority,
  PidTagImportance,
} from "msg-parser";

interface Attachment {
  fileName: string;
  content: Uint8Array;
  contentType: string;
}

interface ParsedRecipient {
  name: string;
  email: string;
  type: "to" | "cc" | "bcc";
}

interface MessageHeaders {
  messageId?: string;
  inReplyTo?: string;
  references?: string;
  replyTo?: string;
  priority?: number; // 1-5 scale (1=highest, 5=lowest)
}

interface ParsedMsg {
  subject: string;
  from: string;
  recipients: ParsedRecipient[];
  date: Date;
  body: string;
  bodyHtml?: string;
  attachments: Attachment[];
  headers?: MessageHeaders;
}

function generateBoundary(): string {
  return "----=_Part_" + Math.random().toString(36).substring(2, 15);
}

function formatEmailDate(date: Date): string {
  const days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

  const d = days[date.getUTCDay()];
  const day = date.getUTCDate();
  const month = months[date.getUTCMonth()];
  const year = date.getUTCFullYear();
  const hours = date.getUTCHours().toString().padStart(2, "0");
  const minutes = date.getUTCMinutes().toString().padStart(2, "0");
  const seconds = date.getUTCSeconds().toString().padStart(2, "0");

  return `${d}, ${day} ${month} ${year} ${hours}:${minutes}:${seconds} +0000`;
}

function encodeBase64(data: Uint8Array | number[]): string {
  return Buffer.from(data).toString("base64");
}

function encodeQuotedPrintable(str: string): string {
  return str.replace(/[^\x20-\x7E\r\n]|=/g, (char) => {
    return "=" + char.charCodeAt(0).toString(16).toUpperCase().padStart(2, "0");
  });
}

/**
 * RFC 5322 header folding.
 * Headers longer than 78 characters should be folded by inserting CRLF
 * followed by a space or tab (continuation).
 * This function is careful not to break:
 * - In the middle of encoded words (=?...?=)
 * - In the middle of email addresses (<...>)
 * - In the middle of quoted strings ("...")
 */
export function foldHeader(headerName: string, headerValue: string, maxLineLength: number = 78): string {
  const fullHeader = `${headerName}: ${headerValue}`;

  // If the header is already short enough, return as-is
  if (fullHeader.length <= maxLineLength) {
    return fullHeader;
  }

  const lines: string[] = [];
  let currentLine = `${headerName}:`;
  let remaining = headerValue;
  let isFirstSegment = true;

  while (remaining.length > 0) {
    // Calculate available space on current line (accounting for leading space)
    const leadingSpace = isFirstSegment ? " " : "\t";
    const availableSpace = maxLineLength - currentLine.length - (isFirstSegment ? 1 : 0);

    if (isFirstSegment) {
      currentLine += " ";
    }

    if (remaining.length <= availableSpace) {
      // Remaining content fits on current line
      currentLine += remaining;
      remaining = "";
    } else {
      // Need to find a safe break point
      const breakPoint = findSafeBreakPoint(remaining, availableSpace);

      if (breakPoint <= 0) {
        // No safe break point found within available space
        // Either the first token is too long, or we need to include it anyway
        const forcedBreak = findFirstBreakPoint(remaining);
        if (forcedBreak > 0 && forcedBreak <= remaining.length) {
          currentLine += remaining.substring(0, forcedBreak).trimEnd();
          remaining = remaining.substring(forcedBreak).trimStart();
        } else {
          // Single token that can't be broken - include it anyway
          currentLine += remaining;
          remaining = "";
        }
      } else {
        currentLine += remaining.substring(0, breakPoint).trimEnd();
        remaining = remaining.substring(breakPoint).trimStart();
      }
    }

    if (remaining.length > 0) {
      lines.push(currentLine);
      currentLine = "\t"; // Use tab for continuation lines
      isFirstSegment = false;
    }
  }

  lines.push(currentLine);
  return lines.join("\r\n");
}

/**
 * Find the first possible break point (space or comma) in the string.
 */
function findFirstBreakPoint(str: string): number {
  for (let i = 0; i < str.length; i++) {
    if (str[i] === " " || str[i] === ",") {
      // For comma, include it in current segment
      return str[i] === "," ? i + 1 : i;
    }
  }
  return -1;
}

/**
 * Find a safe break point within the given max position.
 * Avoids breaking inside:
 * - Encoded words (=?...?=)
 * - Email addresses in angle brackets (<...>)
 * - Quoted strings ("...")
 */
function findSafeBreakPoint(str: string, maxPos: number): number {
  if (maxPos >= str.length) {
    return str.length;
  }

  let bestBreak = -1;
  let inEncodedWord = false;
  let inAngleBracket = false;
  let inQuotes = false;
  let encodedWordStart = -1;

  for (let i = 0; i < maxPos && i < str.length; i++) {
    const char = str[i];
    const nextChar = str[i + 1] || "";

    // Track encoded words (=?charset?encoding?text?=)
    if (char === "=" && nextChar === "?") {
      inEncodedWord = true;
      encodedWordStart = i;
    } else if (inEncodedWord && char === "?" && nextChar === "=") {
      inEncodedWord = false;
      // Skip past the closing ?=
      i++;
      continue;
    }

    // Track angle brackets (email addresses)
    if (char === "<" && !inQuotes && !inEncodedWord) {
      inAngleBracket = true;
    } else if (char === ">" && inAngleBracket) {
      inAngleBracket = false;
    }

    // Track quoted strings
    if (char === '"' && !inEncodedWord) {
      inQuotes = !inQuotes;
    }

    // Only consider break points when not inside special constructs
    if (!inEncodedWord && !inAngleBracket && !inQuotes) {
      // Space is a good break point
      if (char === " ") {
        bestBreak = i;
      }
      // After comma+space is also good for recipient lists
      else if (char === "," && nextChar === " ") {
        bestBreak = i + 1; // Break after the comma
      }
    }
  }

  return bestBreak;
}

function getMimeType(fileName: string): string {
  const ext = fileName.split(".").pop()?.toLowerCase() || "";
  const mimeTypes: Record<string, string> = {
    pdf: "application/pdf",
    doc: "application/msword",
    docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    xls: "application/vnd.ms-excel",
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    gif: "image/gif",
    txt: "text/plain",
    html: "text/html",
    zip: "application/zip",
  };
  return mimeTypes[ext] || "application/octet-stream";
}

function getRecipientType(type: number | undefined): "to" | "cc" | "bcc" {
  // MAPI_TO = 1, MAPI_CC = 2, MAPI_BCC = 3
  if (type === 2) return "cc";
  if (type === 3) return "bcc";
  return "to";
}

/**
 * Maps MAPI priority and importance values to X-Priority scale (1-5).
 * PidTagPriority: -1 (non-urgent), 0 (normal), 1 (urgent)
 * PidTagImportance: 0 (low), 1 (normal), 2 (high)
 * X-Priority: 1 (highest), 2 (high), 3 (normal), 4 (low), 5 (lowest)
 */
export function mapToXPriority(priority: number | undefined, importance: number | undefined): number | undefined {
  // Prefer PidTagPriority if available
  if (priority !== undefined) {
    if (priority === 1) return 1;  // urgent -> highest
    if (priority === 0) return 3;  // normal -> normal
    if (priority === -1) return 5; // non-urgent -> lowest
  }
  // Fall back to PidTagImportance
  if (importance !== undefined) {
    if (importance === 2) return 1; // high -> highest
    if (importance === 1) return 3; // normal -> normal
    if (importance === 0) return 5; // low -> lowest
  }
  return undefined;
}

function parseRecipient(recipient: MsgRecipient): ParsedRecipient {
  const name = recipient.getProperty<string>(PidTagDisplayName) || "";
  const email = recipient.getProperty<string>(PidTagEmailAddress) || name;
  const type = recipient.getProperty<number>(PidTagRecipientType);
  return { name, email, type: getRecipientType(type) };
}

function parseAttachment(attachment: MsgAttachment): Attachment | null {
  const fileName =
    attachment.getProperty<string>(PidTagAttachLongFilename) ||
    attachment.getProperty<string>(PidTagAttachFilename);

  if (!fileName) return null;

  const content = attachment.content();
  if (!content || content.length === 0) return null;

  const mimeTag = attachment.getProperty<string>(PidTagAttachMimeTag);
  const contentType = mimeTag || getMimeType(fileName);

  return {
    fileName,
    content: new Uint8Array(content),
    contentType,
  };
}

function extractSenderEmail(msg: Msg): string | undefined {
  return (
    msg.getProperty<string>(PidTagSenderEmailAddress) ||
    msg.getProperty<string>(PidTagSenderSmtpAddress) ||
    msg.getProperty<string>(PidTagSentRepresentingEmailAddress) ||
    msg.getProperty<string>(PidTagSentRepresentingSmtpAddress) ||
    undefined
  );
}

function extractSenderName(msg: Msg): string | undefined {
  return (
    msg.getProperty<string>(PidTagSenderName) ||
    msg.getProperty<string>(PidTagSentRepresentingName) ||
    undefined
  );
}

export function formatSender(email: string | undefined, name: string | undefined): string {
  if (name && email && name !== email) {
    return `"${name}" <${email}>`;
  }
  return email || name || "unknown@unknown.com";
}

export function parseMsg(buffer: ArrayBuffer): ParsedMsg {
  const msg = Msg.fromUint8Array(new Uint8Array(buffer));

  const subject = msg.getProperty<string>(PidTagSubject) || "(No Subject)";
  const body = msg.getProperty<string>(PidTagBody) || "";
  const bodyHtml = msg.getProperty<string>(PidTagBodyHtml);
  const senderEmail = extractSenderEmail(msg);
  const senderName = extractSenderName(msg);
  const deliveryTime = msg.getProperty<Date>(PidTagMessageDeliveryTime);

  const from = formatSender(senderEmail, senderName);

  const recipients = msg.recipients().map(parseRecipient);
  const attachments = msg.attachments()
    .map(parseAttachment)
    .filter((a): a is Attachment => a !== null);

  // Extract additional message headers
  const messageId = msg.getProperty<string>(PidTagInternetMessageId);
  const inReplyTo = msg.getProperty<string>(PidTagInReplyToId);
  const references = msg.getProperty<string>(PidTagInternetReferences);
  const replyTo = msg.getProperty<string>(PidTagReplyRecipientNames);
  const priority = msg.getProperty<number>(PidTagPriority);
  const importance = msg.getProperty<number>(PidTagImportance);
  const xPriority = mapToXPriority(priority, importance);

  const headers: MessageHeaders = {};
  if (messageId) headers.messageId = messageId;
  if (inReplyTo) headers.inReplyTo = inReplyTo;
  if (references) headers.references = references;
  if (replyTo) headers.replyTo = replyTo;
  if (xPriority !== undefined) headers.priority = xPriority;

  return {
    subject,
    from,
    recipients,
    date: deliveryTime || new Date(),
    body,
    bodyHtml: bodyHtml || undefined,
    attachments,
    headers: Object.keys(headers).length > 0 ? headers : undefined,
  };
}

export function convertToEml(parsed: ParsedMsg): string {
  const hasAttachments = parsed.attachments.length > 0;
  const hasHtml = !!parsed.bodyHtml;
  const boundary = generateBoundary();
  const altBoundary = generateBoundary();

  const toRecipients = parsed.recipients.filter((r) => r.type === "to");
  const ccRecipients = parsed.recipients.filter((r) => r.type === "cc");
  const bccRecipients = parsed.recipients.filter((r) => r.type === "bcc");

  const formatRecipient = (r: ParsedRecipient) =>
    r.name && r.name !== r.email ? `"${r.name}" <${r.email}>` : r.email;

  let eml = "";

  // Headers (using RFC 5322 header folding for long headers)
  eml += foldHeader("From", parsed.from) + "\r\n";
  if (toRecipients.length > 0) {
    eml += foldHeader("To", toRecipients.map(formatRecipient).join(", ")) + "\r\n";
  }
  if (ccRecipients.length > 0) {
    eml += foldHeader("Cc", ccRecipients.map(formatRecipient).join(", ")) + "\r\n";
  }
  if (bccRecipients.length > 0) {
    eml += foldHeader("Bcc", bccRecipients.map(formatRecipient).join(", ")) + "\r\n";
  }
  eml += foldHeader("Subject", parsed.subject) + "\r\n";
  eml += `Date: ${formatEmailDate(parsed.date)}\r\n`;
  eml += `MIME-Version: 1.0\r\n`;

  // Add additional message headers (using header folding for potentially long headers)
  if (parsed.headers) {
    if (parsed.headers.messageId) {
      eml += foldHeader("Message-ID", parsed.headers.messageId) + "\r\n";
    }
    if (parsed.headers.inReplyTo) {
      eml += foldHeader("In-Reply-To", parsed.headers.inReplyTo) + "\r\n";
    }
    if (parsed.headers.references) {
      eml += foldHeader("References", parsed.headers.references) + "\r\n";
    }
    if (parsed.headers.replyTo) {
      eml += foldHeader("Reply-To", parsed.headers.replyTo) + "\r\n";
    }
    if (parsed.headers.priority !== undefined) {
      eml += `X-Priority: ${parsed.headers.priority}\r\n`;
    }
  }

  if (hasAttachments) {
    eml += `Content-Type: multipart/mixed; boundary="${boundary}"\r\n`;
    eml += `\r\n`;
    eml += `--${boundary}\r\n`;

    if (hasHtml) {
      eml += `Content-Type: multipart/alternative; boundary="${altBoundary}"\r\n`;
      eml += `\r\n`;
      eml += `--${altBoundary}\r\n`;
      eml += `Content-Type: text/plain; charset="utf-8"\r\n`;
      eml += `Content-Transfer-Encoding: quoted-printable\r\n`;
      eml += `\r\n`;
      eml += encodeQuotedPrintable(parsed.body);
      eml += `\r\n`;
      eml += `--${altBoundary}\r\n`;
      eml += `Content-Type: text/html; charset="utf-8"\r\n`;
      eml += `Content-Transfer-Encoding: quoted-printable\r\n`;
      eml += `\r\n`;
      eml += encodeQuotedPrintable(parsed.bodyHtml!);
      eml += `\r\n`;
      eml += `--${altBoundary}--\r\n`;
    } else {
      eml += `Content-Type: text/plain; charset="utf-8"\r\n`;
      eml += `Content-Transfer-Encoding: quoted-printable\r\n`;
      eml += `\r\n`;
      eml += encodeQuotedPrintable(parsed.body);
      eml += `\r\n`;
    }

    for (const att of parsed.attachments) {
      eml += `--${boundary}\r\n`;
      eml += `Content-Type: ${att.contentType}; name="${att.fileName}"\r\n`;
      eml += `Content-Disposition: attachment; filename="${att.fileName}"\r\n`;
      eml += `Content-Transfer-Encoding: base64\r\n`;
      eml += `\r\n`;
      const base64 = encodeBase64(att.content);
      for (let i = 0; i < base64.length; i += 76) {
        eml += base64.slice(i, i + 76) + "\r\n";
      }
    }
    eml += `--${boundary}--\r\n`;
  } else if (hasHtml) {
    eml += `Content-Type: multipart/alternative; boundary="${boundary}"\r\n`;
    eml += `\r\n`;
    eml += `--${boundary}\r\n`;
    eml += `Content-Type: text/plain; charset="utf-8"\r\n`;
    eml += `Content-Transfer-Encoding: quoted-printable\r\n`;
    eml += `\r\n`;
    eml += encodeQuotedPrintable(parsed.body);
    eml += `\r\n`;
    eml += `--${boundary}\r\n`;
    eml += `Content-Type: text/html; charset="utf-8"\r\n`;
    eml += `Content-Transfer-Encoding: quoted-printable\r\n`;
    eml += `\r\n`;
    eml += encodeQuotedPrintable(parsed.bodyHtml!);
    eml += `\r\n`;
    eml += `--${boundary}--\r\n`;
  } else {
    eml += `Content-Type: text/plain; charset="utf-8"\r\n`;
    eml += `Content-Transfer-Encoding: quoted-printable\r\n`;
    eml += `\r\n`;
    eml += encodeQuotedPrintable(parsed.body);
  }

  return eml;
}

export function msgToEml(buffer: ArrayBuffer): string {
  const parsed = parseMsg(buffer);
  return convertToEml(parsed);
}
