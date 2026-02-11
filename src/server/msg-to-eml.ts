import {
  Msg,
  EmbeddedMessage,
  Attachment as MsgAttachment,
  Recipient as MsgRecipient,
  PidTagSubject,
  PidTagBody,
  PidTagBodyHtml,
  PidTagRtfCompressed,
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
  PidTagAttachContentId,
  PidTagInternetMessageId,
  PidTagInReplyToId,
  PidTagInternetReferences,
  PidTagReplyRecipientNames,
  PidTagPriority,
  PidTagImportance,
} from "msg-parser";
import { decompressRTF } from "@kenjiuno/decompressrtf";
import { deEncapsulateSync } from "rtf-stream-parser";
import * as iconvLite from "iconv-lite";

interface Attachment {
  fileName: string;
  content: Uint8Array;
  contentType: string;
  contentId?: string;
  isEmbeddedMessage?: boolean;
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

/**
 * Checks if a string contains only ASCII characters.
 */
function isAscii(str: string): boolean {
  for (let i = 0; i < str.length; i++) {
    if (str.charCodeAt(i) > 127) {
      return false;
    }
  }
  return true;
}

/**
 * RFC 2047 MIME encoded-word encoding.
 * Encodes non-ASCII text for use in email headers (Subject, From display name, etc.)
 * Format: =?charset?encoding?encoded-text?=
 *
 * We use Base64 (B) encoding as it's more compact for non-ASCII text.
 * Each encoded word is limited to 75 characters total.
 *
 * @param text The text to encode
 * @returns RFC 2047 encoded string, or original if ASCII-only
 */
export function encodeRfc2047(text: string): string {
  if (!text || isAscii(text)) {
    return text;
  }

  // Encode as UTF-8 Base64
  const utf8Bytes = new TextEncoder().encode(text);
  const base64 = Buffer.from(utf8Bytes).toString("base64");

  // Build the encoded word
  const prefix = "=?UTF-8?B?";
  const suffix = "?=";
  const overhead = prefix.length + suffix.length; // 12 characters
  const maxEncodedLength = 75 - overhead; // 63 characters for encoded text

  // If the entire encoded string fits in one encoded-word, return it directly
  if (base64.length <= maxEncodedLength) {
    return `${prefix}${base64}${suffix}`;
  }

  // Need to split into multiple encoded-words
  // Each encoded-word must be valid UTF-8, so we need to split on character boundaries
  return encodeRfc2047Chunked(text);
}

/**
 * Encodes text as multiple RFC 2047 encoded-words, properly splitting on character boundaries.
 * Each chunk is encoded separately to ensure valid UTF-8 in each encoded-word.
 */
function encodeRfc2047Chunked(text: string): string {
  const prefix = "=?UTF-8?B?";
  const suffix = "?=";
  const overhead = prefix.length + suffix.length;
  const maxEncodedLength = 75 - overhead;

  // We need to determine how many characters we can encode per chunk
  // Base64 expands 3 bytes to 4 characters
  // A UTF-8 character can be 1-4 bytes, so we need to be conservative
  // For safety, we'll aim for chunks that result in ~60 base64 chars (45 bytes before encoding)

  const result: string[] = [];
  let i = 0;

  while (i < text.length) {
    // Try to find the longest substring that fits in one encoded-word
    let chunk = "";
    let chunkBytes = 0;

    while (i < text.length) {
      const char = text[i];
      const charBytes = new TextEncoder().encode(char).length;

      // Check if adding this character would exceed the limit
      // Base64 encoding: ceil(bytes * 4 / 3)
      const newBytes = chunkBytes + charBytes;
      const newBase64Length = Math.ceil((newBytes * 4) / 3);

      if (newBase64Length > maxEncodedLength && chunk.length > 0) {
        // This character would push us over; encode what we have
        break;
      }

      chunk += char;
      chunkBytes = newBytes;
      i++;
    }

    // Encode this chunk
    const chunkUtf8 = new TextEncoder().encode(chunk);
    const chunkBase64 = Buffer.from(chunkUtf8).toString("base64");
    result.push(`${prefix}${chunkBase64}${suffix}`);
  }

  // Join with space, as per RFC 2047 section 5.3
  // When encoded-words are adjacent, they should be separated by linear whitespace
  return result.join(" ");
}

/**
 * Encodes a display name for use in email headers if it contains non-ASCII characters.
 * Returns the name in quotes if ASCII, or RFC 2047 encoded if non-ASCII.
 */
export function encodeDisplayName(name: string): string {
  if (!name) {
    return "";
  }
  if (isAscii(name)) {
    return `"${name}"`;
  }
  return encodeRfc2047(name);
}

/**
 * Encodes a filename according to RFC 2231 for non-ASCII characters.
 * Returns the encoded format: UTF-8''<percent-encoded-value>
 *
 * RFC 2231 specifies that:
 * - attr*=charset'language'encoded-value
 * - We use UTF-8 charset and leave language empty
 * - Characters are percent-encoded (similar to URL encoding)
 */
export function encodeRfc2231(filename: string): string {
  const encoder = new TextEncoder();
  const bytes = encoder.encode(filename);
  let encoded = "";

  for (const byte of bytes) {
    // RFC 2231 allows: ALPHA / DIGIT / "!" / "#" / "$" / "&" / "+" / "-" / "." /
    // "^" / "_" / "`" / "|" / "~" and attribute-char which excludes "*", "'", "%"
    // For simplicity and safety, we only allow alphanumeric, -, ., and _
    if (
      (byte >= 0x30 && byte <= 0x39) || // 0-9
      (byte >= 0x41 && byte <= 0x5A) || // A-Z
      (byte >= 0x61 && byte <= 0x7A) || // a-z
      byte === 0x2D || // -
      byte === 0x2E || // .
      byte === 0x5F    // _
    ) {
      encoded += String.fromCharCode(byte);
    } else {
      // Percent-encode the byte
      encoded += "%" + byte.toString(16).toUpperCase().padStart(2, "0");
    }
  }

  return `UTF-8''${encoded}`;
}

/**
 * Generates Content-Type and Content-Disposition parameters for a filename.
 * Uses RFC 2231 encoding (filename*=) for non-ASCII filenames,
 * otherwise uses the simple quoted form (filename=).
 */
export function formatFilenameParams(filename: string): { name: string; disposition: string } {
  if (isAscii(filename)) {
    return {
      name: `name="${filename}"`,
      disposition: `filename="${filename}"`,
    };
  } else {
    const encoded = encodeRfc2231(filename);
    return {
      name: `name*=${encoded}`,
      disposition: `filename*=${encoded}`,
    };
  }
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
  const contentId = attachment.getProperty<string>(PidTagAttachContentId);

  return {
    fileName,
    content: new Uint8Array(content),
    contentType,
    contentId: contentId || undefined,
  };
}

/**
 * Parses an embedded MSG file (forwarded email, attachment) and converts it to an EML attachment.
 * Uses msg.extractEmbeddedMessage() to get a full Msg object, then recursively converts to EML.
 */
function parseEmbeddedMessage(msg: Msg, embeddedMessage: EmbeddedMessage): Attachment {
  // Get filename from the embedded message attachment properties
  const fileName =
    embeddedMessage.getProperty<string>(PidTagAttachLongFilename) ||
    embeddedMessage.getProperty<string>(PidTagAttachFilename) ||
    "embedded.eml";

  // Ensure the filename has .eml extension
  const emlFileName = fileName.toLowerCase().endsWith(".eml")
    ? fileName
    : fileName.replace(/\.msg$/i, ".eml") || fileName + ".eml";

  // Extract the embedded message as a full Msg object
  const extractedMsg = msg.extractEmbeddedMessage(embeddedMessage);

  // Recursively convert to EML using the internal parse function
  const emlContent = msgToEmlFromMsg(extractedMsg);

  return {
    fileName: emlFileName,
    content: new TextEncoder().encode(emlContent),
    contentType: "message/rfc822",
    isEmbeddedMessage: true,
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
    const encodedName = encodeDisplayName(name);
    return `${encodedName} <${email}>`;
  }
  return email || name || "unknown@unknown.com";
}

interface RtfBodyResult {
  text: string;
  html?: string;
}

/**
 * Decompresses RTF from PidTagRtfCompressed and extracts text/HTML content.
 * RTF may contain HTML wrapped in \fromhtml1 tags (RTF-encapsulated HTML).
 */
export function extractBodyFromRtf(compressedRtf: number[]): RtfBodyResult | null {
  if (!compressedRtf || compressedRtf.length === 0) {
    return null;
  }

  try {
    // Step 1: Decompress the RTF (returns number[])
    const decompressed = decompressRTF(compressedRtf);
    const rtfString = Buffer.from(decompressed).toString("latin1");

    // Step 2: Try to de-encapsulate to extract HTML or text
    // The rtf-stream-parser library throws if RTF is not encapsulated
    try {
      const result = deEncapsulateSync(rtfString, {
        decode: iconvLite.decode,
        mode: "either",
      });

      if (result.mode === "html") {
        // RTF contained encapsulated HTML
        const html = typeof result.text === "string" ? result.text : result.text.toString("utf-8");
        // Also provide a plain text fallback by stripping HTML tags
        const text = stripHtmlTags(html);
        return { text, html };
      } else {
        // RTF contained encapsulated plain text
        const text = typeof result.text === "string" ? result.text : result.text.toString("utf-8");
        return { text };
      }
    } catch {
      // RTF is not encapsulated, extract plain text directly
      const text = extractPlainTextFromRtf(rtfString);
      return text ? { text } : null;
    }
  } catch {
    // If decompression fails, return null
    return null;
  }
}

/**
 * Extracts plain text from raw RTF content (non-encapsulated).
 * This is a simple RTF parser for basic text extraction.
 */
function extractPlainTextFromRtf(rtf: string): string {
  let text = "";
  let i = 0;
  let depth = 0;
  let skipGroup = 0;

  // Groups to skip (they don't contain visible text)
  const skipPatterns = /^\\(fonttbl|colortbl|stylesheet|info|pict|object|fldinst|fldrslt)/;

  while (i < rtf.length) {
    const char = rtf[i];

    if (char === "{") {
      depth++;
      // Check if this group should be skipped
      const remaining = rtf.slice(i + 1, i + 20);
      if (skipPatterns.test(remaining)) {
        skipGroup = depth;
      }
      i++;
    } else if (char === "}") {
      if (skipGroup === depth) {
        skipGroup = 0;
      }
      depth--;
      i++;
    } else if (skipGroup > 0) {
      // Skip content in ignored groups
      i++;
    } else if (char === "\\") {
      // Handle control words and escape sequences
      i++;
      if (i >= rtf.length) break;

      const nextChar = rtf[i];

      // Escape sequences
      if (nextChar === "'" && i + 2 < rtf.length) {
        // Hex escape like \'e9 for é
        const hex = rtf.slice(i + 1, i + 3);
        const code = parseInt(hex, 16);
        if (!isNaN(code)) {
          text += String.fromCharCode(code);
        }
        i += 3;
      } else if (nextChar === "\\" || nextChar === "{" || nextChar === "}") {
        text += nextChar;
        i++;
      } else if (nextChar === "\n" || nextChar === "\r") {
        // Line break in RTF source, continue
        i++;
      } else {
        // Control word - read until space or non-letter
        let controlWord = "";
        while (i < rtf.length && /[a-zA-Z]/.test(rtf[i])) {
          controlWord += rtf[i];
          i++;
        }
        // Skip optional numeric parameter
        while (i < rtf.length && /[-0-9]/.test(rtf[i])) {
          i++;
        }
        // Skip single space delimiter if present
        if (i < rtf.length && rtf[i] === " ") {
          i++;
        }

        // Handle special control words
        if (controlWord === "par" || controlWord === "line") {
          text += "\n";
        } else if (controlWord === "tab") {
          text += "\t";
        }
        // Other control words are ignored
      }
    } else {
      // Regular character
      if (char !== "\r" && char !== "\n") {
        text += char;
      }
      i++;
    }
  }

  return text.trim();
}

/**
 * Simple HTML tag stripper for plain text fallback.
 */
function stripHtmlTags(html: string): string {
  return html
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, "") // Remove style blocks
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, "") // Remove script blocks
    .replace(/<[^>]+>/g, "") // Remove HTML tags
    .replace(/&nbsp;/g, " ") // Replace &nbsp; with space
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\s+/g, " ") // Normalize whitespace
    .trim();
}

function parseMsgFromMsg(msg: Msg): ParsedMsg {
  const subject = msg.getProperty<string>(PidTagSubject) || "(No Subject)";
  let body = msg.getProperty<string>(PidTagBody) || "";
  let bodyHtml = msg.getProperty<string>(PidTagBodyHtml);
  const senderEmail = extractSenderEmail(msg);
  const senderName = extractSenderName(msg);
  const deliveryTime = msg.getProperty<Date>(PidTagMessageDeliveryTime);

  // If body is empty, try to extract from compressed RTF
  if (!body && !bodyHtml) {
    const compressedRtf = msg.getProperty<number[]>(PidTagRtfCompressed);
    if (compressedRtf && compressedRtf.length > 0) {
      const rtfResult = extractBodyFromRtf(compressedRtf);
      if (rtfResult) {
        body = rtfResult.text;
        if (rtfResult.html) {
          bodyHtml = rtfResult.html;
        }
      }
    }
  }

  const from = formatSender(senderEmail, senderName);

  const recipients = msg.recipients().map(parseRecipient);

  // Parse regular attachments
  const regularAttachments = msg.attachments()
    .map(parseAttachment)
    .filter((a): a is Attachment => a !== null);

  // Parse embedded messages (forwarded emails, attached emails)
  const embeddedMessages = msg.embeddedMessages();
  const embeddedAttachments = embeddedMessages.map((embedded) =>
    parseEmbeddedMessage(msg, embedded)
  );

  // Combine regular attachments and embedded message attachments
  const attachments = [...regularAttachments, ...embeddedAttachments];

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

export function parseMsg(buffer: ArrayBuffer): ParsedMsg {
  const msg = Msg.fromUint8Array(new Uint8Array(buffer));
  return parseMsgFromMsg(msg);
}

export function convertToEml(parsed: ParsedMsg): string {
  const hasHtml = !!parsed.bodyHtml;

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
  // Encode non-ASCII subjects using RFC 2047
  eml += foldHeader("Subject", encodeRfc2047(parsed.subject)) + "\r\n";
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

  // Helper to generate the text/plain part
  const generateTextPart = (): string => {
    let part = "";
    part += `Content-Type: text/plain; charset="utf-8"\r\n`;
    part += `Content-Transfer-Encoding: quoted-printable\r\n`;
    part += `\r\n`;
    part += encodeQuotedPrintable(parsed.body);
    return part;
  };

  // Helper to generate the text/html part
  const generateHtmlPart = (): string => {
    let part = "";
    part += `Content-Type: text/html; charset="utf-8"\r\n`;
    part += `Content-Transfer-Encoding: quoted-printable\r\n`;
    part += `\r\n`;
    part += encodeQuotedPrintable(parsed.bodyHtml!);
    return part;
  };

  // Helper to generate multipart/alternative content
  const generateAlternativePart = (boundary: string): string => {
    let part = "";
    part += `Content-Type: multipart/alternative; boundary="${boundary}"\r\n`;
    part += `\r\n`;
    part += `--${boundary}\r\n`;
    part += generateTextPart();
    part += `\r\n`;
    part += `--${boundary}\r\n`;
    part += generateHtmlPart();
    part += `\r\n`;
    part += `--${boundary}--\r\n`;
    return part;
  };

  // Helper to generate attachment part
  const generateAttachmentPart = (att: Attachment, inline: boolean): string => {
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
        part += base64.slice(i, i + 76) + "\r\n";
      }
    }
    return part;
  };

  // Determine the structure based on content types
  // Cases:
  // 1. HTML + inline attachments + regular attachments: multipart/mixed > (multipart/related > (multipart/alternative + inline)) + regular
  // 2. HTML + inline attachments only: multipart/related > multipart/alternative + inline
  // 3. HTML + regular attachments only: multipart/mixed > multipart/alternative + regular
  // 4. HTML only: multipart/alternative
  // 5. Plain + any attachments: multipart/mixed > text/plain + attachments
  // 6. Plain only: text/plain

  if (hasHtml && hasInlineAttachments && hasRegularAttachments) {
    // Case 1: multipart/mixed > multipart/related > (multipart/alternative + inline) + regular attachments
    eml += `Content-Type: multipart/mixed; boundary="${mixedBoundary}"\r\n`;
    eml += `\r\n`;
    eml += `--${mixedBoundary}\r\n`;
    eml += `Content-Type: multipart/related; boundary="${relatedBoundary}"\r\n`;
    eml += `\r\n`;
    eml += `--${relatedBoundary}\r\n`;
    eml += generateAlternativePart(altBoundary);
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
    eml += generateAlternativePart(altBoundary);
    for (const att of inlineAttachments) {
      eml += `--${relatedBoundary}\r\n`;
      eml += generateAttachmentPart(att, true);
    }
    eml += `--${relatedBoundary}--\r\n`;
  } else if (hasHtml && hasRegularAttachments) {
    // Case 3: multipart/mixed > multipart/alternative + regular attachments
    eml += `Content-Type: multipart/mixed; boundary="${mixedBoundary}"\r\n`;
    eml += `\r\n`;
    eml += `--${mixedBoundary}\r\n`;
    eml += generateAlternativePart(altBoundary);
    for (const att of regularAttachments) {
      eml += `--${mixedBoundary}\r\n`;
      eml += generateAttachmentPart(att, false);
    }
    eml += `--${mixedBoundary}--\r\n`;
  } else if (hasHtml) {
    // Case 4: multipart/alternative only
    eml += generateAlternativePart(altBoundary);
  } else if (hasAttachments) {
    // Case 5: multipart/mixed > text/plain + attachments
    eml += `Content-Type: multipart/mixed; boundary="${mixedBoundary}"\r\n`;
    eml += `\r\n`;
    eml += `--${mixedBoundary}\r\n`;
    eml += generateTextPart();
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
function msgToEmlFromMsg(msg: Msg): string {
  const parsed = parseMsgFromMsg(msg);
  return convertToEml(parsed);
}

export function msgToEml(buffer: ArrayBuffer): string {
  const parsed = parseMsg(buffer);
  return convertToEml(parsed);
}
