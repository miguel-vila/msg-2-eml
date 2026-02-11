import { Msg } from "msg-parser";
import { encodeDisplayName, encodeQuotedPrintable, encodeRfc2047, foldHeader } from "../encoding/index.js";
import { formatEmailDate, generateBoundary } from "../mime/index.js";
import { parseMsgFromMsg } from "../parsing/index.js";
import type { ParsedMsg, ParsedRecipient } from "../types/index.js";
import { generateAlternativePart, generateAttachmentPart, generateTextPart } from "./parts.js";

export function convertToEml(parsed: ParsedMsg): string {
  const hasHtml = !!parsed.bodyHtml;
  const hasCalendar = !!parsed.calendarEvent;

  // Separate inline attachments (with contentId) from regular attachments
  const inlineAttachments = parsed.attachments.filter((a) => a.contentId);
  const regularAttachments = parsed.attachments.filter((a) => !a.contentId);

  const hasInlineAttachments = inlineAttachments.length > 0;
  const hasRegularAttachments = regularAttachments.length > 0;
  const hasAttachments = hasInlineAttachments || hasRegularAttachments;

  const mixedBoundary = generateBoundary();
  const relatedBoundary = generateBoundary();
  const altBoundary = generateBoundary();

  const toRecipients = parsed.recipients.filter((r) => r.type === "to");
  const ccRecipients = parsed.recipients.filter((r) => r.type === "cc");
  const bccRecipients = parsed.recipients.filter((r) => r.type === "bcc");

  const formatRecipient = (r: ParsedRecipient) => {
    if (r.name && r.name !== r.email) {
      const encodedName = encodeDisplayName(r.name);
      return `${encodedName} <${r.email}>`;
    }
    return r.email;
  };

  let eml = "";

  // Headers (using RFC 5322 header folding for long headers)
  eml += `${foldHeader("From", parsed.from)}\r\n`;
  if (toRecipients.length > 0) {
    eml += `${foldHeader("To", toRecipients.map(formatRecipient).join(", "))}\r\n`;
  }
  if (ccRecipients.length > 0) {
    eml += `${foldHeader("Cc", ccRecipients.map(formatRecipient).join(", "))}\r\n`;
  }
  if (bccRecipients.length > 0) {
    eml += `${foldHeader("Bcc", bccRecipients.map(formatRecipient).join(", "))}\r\n`;
  }
  // Encode non-ASCII subjects using RFC 2047
  eml += `${foldHeader("Subject", encodeRfc2047(parsed.subject))}\r\n`;
  eml += `Date: ${formatEmailDate(parsed.date)}\r\n`;
  eml += `MIME-Version: 1.0\r\n`;

  // Add additional message headers (using header folding for potentially long headers)
  if (parsed.headers) {
    if (parsed.headers.messageId) {
      eml += `${foldHeader("Message-ID", parsed.headers.messageId)}\r\n`;
    }
    if (parsed.headers.inReplyTo) {
      eml += `${foldHeader("In-Reply-To", parsed.headers.inReplyTo)}\r\n`;
    }
    if (parsed.headers.references) {
      eml += `${foldHeader("References", parsed.headers.references)}\r\n`;
    }
    if (parsed.headers.replyTo) {
      eml += `${foldHeader("Reply-To", parsed.headers.replyTo)}\r\n`;
    }
    if (parsed.headers.priority !== undefined) {
      eml += `X-Priority: ${parsed.headers.priority}\r\n`;
    }
    if (parsed.headers.dispositionNotificationTo) {
      eml += `${foldHeader("Disposition-Notification-To", parsed.headers.dispositionNotificationTo)}\r\n`;
    }
    if (parsed.headers.returnReceiptTo) {
      eml += `${foldHeader("Return-Receipt-To", parsed.headers.returnReceiptTo)}\r\n`;
    }
  }

  // Check if we need multipart/alternative (HTML and/or calendar)
  const needsAlternative = hasHtml || hasCalendar;

  if (hasHtml && hasInlineAttachments && hasRegularAttachments) {
    // Case 1: multipart/mixed > multipart/related > (multipart/alternative + inline) + regular attachments
    eml += `Content-Type: multipart/mixed; boundary="${mixedBoundary}"\r\n`;
    eml += `\r\n`;
    eml += `--${mixedBoundary}\r\n`;
    eml += `Content-Type: multipart/related; boundary="${relatedBoundary}"\r\n`;
    eml += `\r\n`;
    eml += `--${relatedBoundary}\r\n`;
    eml += generateAlternativePart(altBoundary, parsed, hasHtml, hasCalendar);
    for (const att of inlineAttachments) {
      eml += `--${relatedBoundary}\r\n`;
      eml += generateAttachmentPart(att, true);
    }
    eml += `--${relatedBoundary}--\r\n`;
    for (const att of regularAttachments) {
      eml += `--${mixedBoundary}\r\n`;
      eml += generateAttachmentPart(att, false);
    }
    eml += `--${mixedBoundary}--\r\n`;
  } else if (hasHtml && hasInlineAttachments) {
    // Case 2: multipart/related > multipart/alternative + inline attachments
    eml += `Content-Type: multipart/related; boundary="${relatedBoundary}"\r\n`;
    eml += `\r\n`;
    eml += `--${relatedBoundary}\r\n`;
    eml += generateAlternativePart(altBoundary, parsed, hasHtml, hasCalendar);
    for (const att of inlineAttachments) {
      eml += `--${relatedBoundary}\r\n`;
      eml += generateAttachmentPart(att, true);
    }
    eml += `--${relatedBoundary}--\r\n`;
  } else if (needsAlternative && hasRegularAttachments) {
    // Case 3: multipart/mixed > multipart/alternative + regular attachments
    eml += `Content-Type: multipart/mixed; boundary="${mixedBoundary}"\r\n`;
    eml += `\r\n`;
    eml += `--${mixedBoundary}\r\n`;
    eml += generateAlternativePart(altBoundary, parsed, hasHtml, hasCalendar);
    for (const att of regularAttachments) {
      eml += `--${mixedBoundary}\r\n`;
      eml += generateAttachmentPart(att, false);
    }
    eml += `--${mixedBoundary}--\r\n`;
  } else if (needsAlternative) {
    // Case 4: multipart/alternative only (HTML and/or calendar)
    eml += generateAlternativePart(altBoundary, parsed, hasHtml, hasCalendar);
  } else if (hasAttachments) {
    // Case 5: multipart/mixed > text/plain + attachments
    eml += `Content-Type: multipart/mixed; boundary="${mixedBoundary}"\r\n`;
    eml += `\r\n`;
    eml += `--${mixedBoundary}\r\n`;
    eml += generateTextPart(parsed.body);
    eml += `\r\n`;
    for (const att of parsed.attachments) {
      eml += `--${mixedBoundary}\r\n`;
      eml += generateAttachmentPart(att, false);
    }
    eml += `--${mixedBoundary}--\r\n`;
  } else {
    // Case 6: text/plain only
    eml += `Content-Type: text/plain; charset="utf-8"\r\n`;
    eml += `Content-Transfer-Encoding: quoted-printable\r\n`;
    eml += `\r\n`;
    eml += encodeQuotedPrintable(parsed.body);
  }

  return eml;
}

/**
 * Converts a Msg object directly to EML format.
 * Used internally for recursive conversion of embedded messages.
 */
export function msgToEmlFromMsg(msg: Msg): string {
  const parsed = parseMsgFromMsg(msg, msgToEmlFromMsg);
  return convertToEml(parsed);
}

export function msgToEml(buffer: ArrayBuffer): string {
  const msg = Msg.fromUint8Array(new Uint8Array(buffer));
  const parsed = parseMsgFromMsg(msg, msgToEmlFromMsg);
  return convertToEml(parsed);
}

export function parseMsg(buffer: ArrayBuffer): ParsedMsg {
  const msg = Msg.fromUint8Array(new Uint8Array(buffer));
  return parseMsgFromMsg(msg, msgToEmlFromMsg);
}
