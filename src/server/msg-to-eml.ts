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

  // Headers
  eml += `From: ${parsed.from}\r\n`;
  if (toRecipients.length > 0) {
    eml += `To: ${toRecipients.map(formatRecipient).join(", ")}\r\n`;
  }
  if (ccRecipients.length > 0) {
    eml += `Cc: ${ccRecipients.map(formatRecipient).join(", ")}\r\n`;
  }
  if (bccRecipients.length > 0) {
    eml += `Bcc: ${bccRecipients.map(formatRecipient).join(", ")}\r\n`;
  }
  eml += `Subject: ${parsed.subject}\r\n`;
  eml += `Date: ${formatEmailDate(parsed.date)}\r\n`;
  eml += `MIME-Version: 1.0\r\n`;

  // Add additional message headers
  if (parsed.headers) {
    if (parsed.headers.messageId) {
      eml += `Message-ID: ${parsed.headers.messageId}\r\n`;
    }
    if (parsed.headers.inReplyTo) {
      eml += `In-Reply-To: ${parsed.headers.inReplyTo}\r\n`;
    }
    if (parsed.headers.references) {
      eml += `References: ${parsed.headers.references}\r\n`;
    }
    if (parsed.headers.replyTo) {
      eml += `Reply-To: ${parsed.headers.replyTo}\r\n`;
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
