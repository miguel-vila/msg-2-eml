#!/usr/bin/env npx tsx
/**
 * Debug script for MSG to EML conversion
 * Usage: npx tsx scripts/debug-conversion.ts [options] <path-to-msg-file>
 *
 * Options:
 *   --full-eml      Output the complete EML content
 *   --save-eml      Save the EML output to a file (same name with .eml extension)
 *   --json          Output gap analysis as JSON
 *   --help          Show this help message
 *
 * Outputs structured debug information including:
 * - Input MSG data in structured format
 * - Output EML content
 * - Gap analysis identifying missing or potentially problematic fields
 */

import * as fs from "node:fs";
import * as path from "node:path";

interface CliOptions {
  filePath: string | null;
  fullEml: boolean;
  saveEml: boolean;
  json: boolean;
  help: boolean;
}

function parseArgs(args: string[]): CliOptions {
  const options: CliOptions = {
    filePath: null,
    fullEml: false,
    saveEml: false,
    json: false,
    help: false,
  };

  for (const arg of args) {
    if (arg === "--full-eml") {
      options.fullEml = true;
    } else if (arg === "--save-eml") {
      options.saveEml = true;
    } else if (arg === "--json") {
      options.json = true;
    } else if (arg === "--help" || arg === "-h") {
      options.help = true;
    } else if (!arg.startsWith("-")) {
      options.filePath = arg;
    }
  }

  return options;
}

function printHelp(): void {
  console.log(`
${colorize("MSG to EML Debug Conversion Script", "bold")}

${colorize("Usage:", "cyan")}
  npx tsx scripts/debug-conversion.ts [options] <path-to-msg-file>
  npm run debug -- [options] <path-to-msg-file>

${colorize("Options:", "cyan")}
  --full-eml    Output the complete EML content (not just headers)
  --save-eml    Save the EML output to a file (same name with .eml extension)
  --json        Output gap analysis as JSON (for scripting)
  --help, -h    Show this help message

${colorize("Examples:", "cyan")}
  npm run debug -- samples/email.msg
  npm run debug -- --full-eml samples/email.msg
  npm run debug -- --save-eml --json samples/email.msg
`);
}
import {
  Msg,
  PidLidAppointmentEndWhole,
  PidLidAppointmentStartWhole,
  PidLidCcAttendeesString,
  PidLidLocation,
  PidLidToAttendeesString,
  PidTagBody,
  PidTagBodyHtml,
  PidTagImportance,
  PidTagInReplyToId,
  PidTagInternetMessageId,
  PidTagInternetReferences,
  PidTagMessageClass,
  PidTagMessageDeliveryTime,
  PidTagOriginatorDeliveryReportRequested,
  PidTagPriority,
  PidTagReadReceiptRequested,
  PidTagReplyRecipientNames,
  PidTagRtfCompressed,
  PidTagSenderEmailAddress,
  PidTagSenderName,
  PidTagSenderSmtpAddress,
  PidTagSubject,
} from "msg-parser";
import { msgToEml, parseMsg } from "../src/server/msg-to-eml.js";
import type { ParsedMsg, ParsedRecipient, Attachment } from "../src/server/msg-to-eml/types/index.js";

// ANSI color codes
const colors = {
  reset: "\x1b[0m",
  bold: "\x1b[1m",
  dim: "\x1b[2m",
  red: "\x1b[31m",
  green: "\x1b[32m",
  yellow: "\x1b[33m",
  blue: "\x1b[34m",
  magenta: "\x1b[35m",
  cyan: "\x1b[36m",
};

function colorize(text: string, color: keyof typeof colors): string {
  return `${colors[color]}${text}${colors.reset}`;
}

function printHeader(title: string): void {
  const line = "═".repeat(60);
  console.log();
  console.log(colorize(line, "blue"));
  console.log(colorize(`  ${title}`, "bold"));
  console.log(colorize(line, "blue"));
}

function printSection(title: string): void {
  console.log();
  console.log(colorize(`▸ ${title}`, "cyan"));
  console.log(colorize("─".repeat(40), "dim"));
}

function printField(label: string, value: unknown, indent = 2): void {
  const prefix = " ".repeat(indent);
  const formattedLabel = colorize(`${label}:`, "dim");

  if (value === undefined || value === null || value === "") {
    console.log(`${prefix}${formattedLabel} ${colorize("(empty)", "yellow")}`);
  } else if (typeof value === "string") {
    // Truncate long strings
    const maxLen = 100;
    const displayValue = value.length > maxLen ? `${value.substring(0, maxLen)}...` : value;
    console.log(`${prefix}${formattedLabel} ${displayValue}`);
  } else if (value instanceof Date) {
    console.log(`${prefix}${formattedLabel} ${value.toISOString()}`);
  } else if (Array.isArray(value)) {
    if (value.length === 0) {
      console.log(`${prefix}${formattedLabel} ${colorize("(empty array)", "yellow")}`);
    } else {
      console.log(`${prefix}${formattedLabel} [${value.length} items]`);
    }
  } else if (typeof value === "object") {
    console.log(`${prefix}${formattedLabel}`);
    for (const [k, v] of Object.entries(value)) {
      printField(k, v, indent + 2);
    }
  } else {
    console.log(`${prefix}${formattedLabel} ${value}`);
  }
}

interface Gap {
  field: string;
  severity: "warning" | "error" | "info";
  message: string;
  inputValue?: unknown;
  outputValue?: unknown;
}

interface RawMsgData {
  subject: string | undefined;
  senderName: string | undefined;
  senderEmail: string | undefined;
  senderSmtp: string | undefined;
  body: string | undefined;
  bodyHtml: string | undefined;
  hasCompressedRtf: boolean;
  deliveryTime: Date | undefined;
  messageClass: string | undefined;
  messageId: string | undefined;
  inReplyTo: string | undefined;
  references: string | undefined;
  replyTo: string | undefined;
  priority: number | undefined;
  importance: number | undefined;
  readReceiptRequested: boolean | undefined;
  deliveryReceiptRequested: boolean | undefined;
  recipientCount: number;
  attachmentCount: number;
  embeddedMessageCount: number;
  // Calendar fields
  appointmentStart: Date | undefined;
  appointmentEnd: Date | undefined;
  location: string | undefined;
  toAttendees: string | undefined;
  ccAttendees: string | undefined;
}

function safeGetProperty<T>(msg: Msg, prop: unknown): T | undefined {
  try {
    return msg.getProperty<T>(prop);
  } catch {
    return undefined;
  }
}

function extractRawMsgData(msg: Msg): RawMsgData {
  const compressedRtf = safeGetProperty<number[]>(msg, PidTagRtfCompressed);

  let recipientCount = 0;
  let attachmentCount = 0;
  let embeddedMessageCount = 0;

  try {
    recipientCount = msg.recipients().length;
  } catch {
    // ignore
  }
  try {
    attachmentCount = msg.attachments().length;
  } catch {
    // ignore
  }
  try {
    embeddedMessageCount = msg.embeddedMessages().length;
  } catch {
    // ignore
  }

  return {
    subject: safeGetProperty<string>(msg, PidTagSubject),
    senderName: safeGetProperty<string>(msg, PidTagSenderName),
    senderEmail: safeGetProperty<string>(msg, PidTagSenderEmailAddress),
    senderSmtp: safeGetProperty<string>(msg, PidTagSenderSmtpAddress),
    body: safeGetProperty<string>(msg, PidTagBody),
    bodyHtml: safeGetProperty<string>(msg, PidTagBodyHtml),
    hasCompressedRtf: !!(compressedRtf && compressedRtf.length > 0),
    deliveryTime: safeGetProperty<Date>(msg, PidTagMessageDeliveryTime),
    messageClass: safeGetProperty<string>(msg, PidTagMessageClass),
    messageId: safeGetProperty<string>(msg, PidTagInternetMessageId),
    inReplyTo: safeGetProperty<string>(msg, PidTagInReplyToId),
    references: safeGetProperty<string>(msg, PidTagInternetReferences),
    replyTo: safeGetProperty<string>(msg, PidTagReplyRecipientNames),
    priority: safeGetProperty<number>(msg, PidTagPriority),
    importance: safeGetProperty<number>(msg, PidTagImportance),
    readReceiptRequested: safeGetProperty<boolean>(msg, PidTagReadReceiptRequested),
    deliveryReceiptRequested: safeGetProperty<boolean>(msg, PidTagOriginatorDeliveryReportRequested),
    recipientCount,
    attachmentCount,
    embeddedMessageCount,
    // Calendar fields (these are named properties and may fail)
    appointmentStart: safeGetProperty<Date>(msg, PidLidAppointmentStartWhole),
    appointmentEnd: safeGetProperty<Date>(msg, PidLidAppointmentEndWhole),
    location: safeGetProperty<string>(msg, PidLidLocation),
    toAttendees: safeGetProperty<string>(msg, PidLidToAttendeesString),
    ccAttendees: safeGetProperty<string>(msg, PidLidCcAttendeesString),
  };
}

function analyzeGaps(rawData: RawMsgData, parsed: ParsedMsg, emlContent: string): Gap[] {
  const gaps: Gap[] = [];

  // Check subject
  if (!rawData.subject && parsed.subject === "(No Subject)") {
    gaps.push({
      field: "subject",
      severity: "info",
      message: "No subject in source MSG, using default",
    });
  }

  // Check sender
  if (!rawData.senderEmail && !rawData.senderSmtp) {
    gaps.push({
      field: "from",
      severity: "warning",
      message: "No sender email address found in MSG properties",
      inputValue: { name: rawData.senderName, email: rawData.senderEmail, smtp: rawData.senderSmtp },
    });
  }

  // Check if From is using X500 address instead of SMTP
  const fromUsesX500 = parsed.from.includes("/O=EXCHANGELABS") || parsed.from.includes("/o=ExchangeLabs");
  if (fromUsesX500 && rawData.senderSmtp) {
    gaps.push({
      field: "from",
      severity: "warning",
      message: "From uses X500 address but SMTP address is available",
      inputValue: { x500: rawData.senderEmail, smtp: rawData.senderSmtp },
    });
  }

  // Check body content
  if (!rawData.body && !rawData.bodyHtml && !rawData.hasCompressedRtf) {
    gaps.push({
      field: "body",
      severity: "warning",
      message: "No body content found (plain text, HTML, or RTF)",
    });
  } else if (!rawData.body && !rawData.bodyHtml && rawData.hasCompressedRtf) {
    if (!parsed.body && !parsed.bodyHtml) {
      gaps.push({
        field: "body",
        severity: "error",
        message: "Failed to extract body from compressed RTF",
      });
    } else {
      gaps.push({
        field: "body",
        severity: "info",
        message: "Body extracted from compressed RTF",
      });
    }
  }

  // Check date
  if (!rawData.deliveryTime) {
    gaps.push({
      field: "date",
      severity: "warning",
      message: "No delivery time in MSG, using current date",
    });
  }

  // Check recipients
  if (rawData.recipientCount === 0) {
    gaps.push({
      field: "recipients",
      severity: "warning",
      message: "No recipients found in MSG",
    });
  } else if (parsed.recipients.length !== rawData.recipientCount) {
    gaps.push({
      field: "recipients",
      severity: "error",
      message: `Recipient count mismatch: MSG has ${rawData.recipientCount}, parsed ${parsed.recipients.length}`,
      inputValue: rawData.recipientCount,
      outputValue: parsed.recipients.length,
    });
  }

  // Check for recipients without emails
  const recipientsWithoutEmail = parsed.recipients.filter((r) => !r.email);
  if (recipientsWithoutEmail.length > 0) {
    gaps.push({
      field: "recipients",
      severity: "warning",
      message: `${recipientsWithoutEmail.length} recipient(s) without email address`,
      inputValue: recipientsWithoutEmail.map((r) => r.name),
    });
  }

  // Check for recipients using X500 addresses instead of SMTP
  const recipientsWithX500 = parsed.recipients.filter(
    (r) => r.email && (r.email.includes("/O=EXCHANGELABS") || r.email.includes("/o=ExchangeLabs") || r.email.startsWith("/"))
  );
  if (recipientsWithX500.length > 0) {
    gaps.push({
      field: "recipients",
      severity: "warning",
      message: `${recipientsWithX500.length} recipient(s) using X500/Exchange address instead of SMTP`,
      inputValue: recipientsWithX500.map((r) => ({ name: r.name, email: r.email })),
    });
  }

  // Check attachments
  const totalInputAttachments = rawData.attachmentCount + rawData.embeddedMessageCount;
  if (parsed.attachments.length !== totalInputAttachments) {
    gaps.push({
      field: "attachments",
      severity: "warning",
      message: `Attachment count: ${rawData.attachmentCount} regular + ${rawData.embeddedMessageCount} embedded = ${totalInputAttachments} total, parsed ${parsed.attachments.length}`,
      inputValue: totalInputAttachments,
      outputValue: parsed.attachments.length,
    });
  }

  // Check for attachments without content
  const emptyAttachments = parsed.attachments.filter((a) => a.content.length === 0);
  if (emptyAttachments.length > 0) {
    gaps.push({
      field: "attachments",
      severity: "error",
      message: `${emptyAttachments.length} attachment(s) with empty content`,
      inputValue: emptyAttachments.map((a) => a.fileName),
    });
  }

  // Check message headers
  if (rawData.messageId && !parsed.headers?.messageId) {
    gaps.push({
      field: "headers.messageId",
      severity: "error",
      message: "Message-ID present in MSG but not parsed",
      inputValue: rawData.messageId,
    });
  }

  if (rawData.inReplyTo && !parsed.headers?.inReplyTo) {
    gaps.push({
      field: "headers.inReplyTo",
      severity: "error",
      message: "In-Reply-To present in MSG but not parsed",
      inputValue: rawData.inReplyTo,
    });
  }

  if (rawData.references && !parsed.headers?.references) {
    gaps.push({
      field: "headers.references",
      severity: "error",
      message: "References present in MSG but not parsed",
      inputValue: rawData.references,
    });
  }

  if (rawData.replyTo && !parsed.headers?.replyTo) {
    gaps.push({
      field: "headers.replyTo",
      severity: "error",
      message: "Reply-To present in MSG but not parsed",
      inputValue: rawData.replyTo,
    });
  }

  // Check priority
  if ((rawData.priority !== undefined || rawData.importance !== undefined) && !parsed.headers?.priority) {
    gaps.push({
      field: "headers.priority",
      severity: "warning",
      message: "Priority/Importance in MSG but not mapped",
      inputValue: { priority: rawData.priority, importance: rawData.importance },
    });
  }

  // Check read receipt
  if (rawData.readReceiptRequested && !parsed.headers?.dispositionNotificationTo) {
    gaps.push({
      field: "headers.dispositionNotificationTo",
      severity: "warning",
      message: "Read receipt requested but not set in parsed headers",
    });
  }

  // Check delivery receipt
  if (rawData.deliveryReceiptRequested && !parsed.headers?.returnReceiptTo) {
    gaps.push({
      field: "headers.returnReceiptTo",
      severity: "warning",
      message: "Delivery receipt requested but not set in parsed headers",
    });
  }

  // Check calendar event
  const isCalendar = rawData.messageClass?.startsWith("IPM.Appointment");
  if (isCalendar) {
    if (!parsed.calendarEvent) {
      gaps.push({
        field: "calendarEvent",
        severity: "error",
        message: "Calendar message detected but no calendar event parsed",
        inputValue: rawData.messageClass,
      });
    } else {
      if (rawData.appointmentStart && !parsed.calendarEvent.startTime) {
        gaps.push({
          field: "calendarEvent.startTime",
          severity: "error",
          message: "Appointment start time present but not parsed",
        });
      }
      if (rawData.appointmentEnd && !parsed.calendarEvent.endTime) {
        gaps.push({
          field: "calendarEvent.endTime",
          severity: "error",
          message: "Appointment end time present but not parsed",
        });
      }
    }
  }

  // EML output validation
  if (!emlContent.includes("From:")) {
    gaps.push({
      field: "eml.from",
      severity: "error",
      message: "EML output missing From header",
    });
  }

  if (!emlContent.includes("Date:")) {
    gaps.push({
      field: "eml.date",
      severity: "error",
      message: "EML output missing Date header",
    });
  }

  if (!emlContent.includes("MIME-Version:")) {
    gaps.push({
      field: "eml.mime",
      severity: "warning",
      message: "EML output missing MIME-Version header",
    });
  }

  // Check for encoding issues (non-ASCII in headers without proper encoding)
  const headerSection = emlContent.split("\r\n\r\n")[0];
  const nonAsciiInHeaders = /[\u0080-\uffff]/.test(headerSection);
  const hasEncodedHeaders = /=\?UTF-8\?[BQ]\?/.test(headerSection);
  if (nonAsciiInHeaders && !hasEncodedHeaders) {
    gaps.push({
      field: "eml.encoding",
      severity: "warning",
      message: "Non-ASCII characters in headers without RFC 2047 encoding",
    });
  }

  return gaps;
}

function printRecipients(recipients: ParsedRecipient[]): void {
  const byType = {
    to: recipients.filter((r) => r.type === "to"),
    cc: recipients.filter((r) => r.type === "cc"),
    bcc: recipients.filter((r) => r.type === "bcc"),
  };

  for (const [type, list] of Object.entries(byType)) {
    if (list.length > 0) {
      console.log(`    ${colorize(type.toUpperCase(), "dim")}:`);
      for (const r of list) {
        const email = r.email || colorize("(no email)", "yellow");
        const name = r.name || colorize("(no name)", "dim");
        console.log(`      - ${name} <${email}>`);
      }
    }
  }
}

function printAttachments(attachments: Attachment[]): void {
  if (attachments.length === 0) {
    console.log(`    ${colorize("(none)", "dim")}`);
    return;
  }

  for (const att of attachments) {
    const size = att.content.length;
    const sizeStr = size > 1024 ? `${(size / 1024).toFixed(1)} KB` : `${size} bytes`;
    const flags: string[] = [];
    if (att.contentId) flags.push("inline");
    if (att.isEmbeddedMessage) flags.push("embedded msg");
    const flagStr = flags.length > 0 ? colorize(` [${flags.join(", ")}]`, "magenta") : "";
    console.log(`    - ${att.fileName} (${att.contentType}, ${sizeStr})${flagStr}`);
  }
}

function printGaps(gaps: Gap[]): void {
  if (gaps.length === 0) {
    console.log(`  ${colorize("✓ No gaps detected", "green")}`);
    return;
  }

  const errors = gaps.filter((g) => g.severity === "error");
  const warnings = gaps.filter((g) => g.severity === "warning");
  const infos = gaps.filter((g) => g.severity === "info");

  console.log(`  ${colorize(`${errors.length} errors`, "red")}, ${colorize(`${warnings.length} warnings`, "yellow")}, ${colorize(`${infos.length} info`, "dim")}`);
  console.log();

  for (const gap of gaps) {
    const icon = gap.severity === "error" ? colorize("✗", "red") : gap.severity === "warning" ? colorize("!", "yellow") : colorize("ℹ", "dim");
    const fieldColor = gap.severity === "error" ? "red" : gap.severity === "warning" ? "yellow" : "dim";
    console.log(`  ${icon} ${colorize(gap.field, fieldColor)}: ${gap.message}`);
    if (gap.inputValue !== undefined) {
      console.log(`      Input: ${JSON.stringify(gap.inputValue)}`);
    }
    if (gap.outputValue !== undefined) {
      console.log(`      Output: ${JSON.stringify(gap.outputValue)}`);
    }
  }
}

async function main(): Promise<void> {
  const args = process.argv.slice(2);
  const options = parseArgs(args);

  if (options.help) {
    printHelp();
    process.exit(0);
  }

  if (!options.filePath) {
    console.error(colorize("Error: No file specified", "red"));
    printHelp();
    process.exit(1);
  }

  const filePath = path.resolve(options.filePath);

  if (!fs.existsSync(filePath)) {
    console.error(colorize(`File not found: ${filePath}`, "red"));
    process.exit(1);
  }

  if (!filePath.toLowerCase().endsWith(".msg")) {
    console.error(colorize("Warning: File does not have .msg extension", "yellow"));
  }

  // Read the MSG file
  const buffer = fs.readFileSync(filePath);
  const arrayBuffer = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);

  // Parse raw MSG data
  const msg = Msg.fromUint8Array(new Uint8Array(arrayBuffer));
  const rawData = extractRawMsgData(msg);

  // Parse using our conversion logic
  const parsed = parseMsg(arrayBuffer);

  // Convert to EML
  const emlContent = msgToEml(arrayBuffer);

  // Analyze gaps
  const gaps = analyzeGaps(rawData, parsed, emlContent);

  // JSON output mode
  if (options.json) {
    const output = {
      file: path.basename(filePath),
      size: buffer.length,
      input: {
        subject: rawData.subject,
        messageClass: rawData.messageClass,
        senderName: rawData.senderName,
        senderEmail: rawData.senderEmail,
        senderSmtp: rawData.senderSmtp,
        deliveryTime: rawData.deliveryTime?.toISOString(),
        bodyLength: rawData.body?.length ?? 0,
        bodyHtmlLength: rawData.bodyHtml?.length ?? 0,
        hasCompressedRtf: rawData.hasCompressedRtf,
        recipientCount: rawData.recipientCount,
        attachmentCount: rawData.attachmentCount,
        embeddedMessageCount: rawData.embeddedMessageCount,
        messageId: rawData.messageId,
      },
      output: {
        subject: parsed.subject,
        from: parsed.from,
        date: parsed.date.toISOString(),
        bodyLength: parsed.body?.length ?? 0,
        bodyHtmlLength: parsed.bodyHtml?.length ?? 0,
        recipients: parsed.recipients,
        attachments: parsed.attachments.map((a) => ({
          fileName: a.fileName,
          contentType: a.contentType,
          size: a.content.length,
          isInline: !!a.contentId,
          isEmbeddedMessage: !!a.isEmbeddedMessage,
        })),
        headers: parsed.headers,
        hasCalendarEvent: !!parsed.calendarEvent,
      },
      eml: {
        size: emlContent.length,
        lineCount: emlContent.split("\r\n").length,
      },
      errors: gaps
        .filter((g) => g.severity === "error")
        .map((g) => ({
          field: g.field,
          message: g.message,
          inputValue: g.inputValue,
          outputValue: g.outputValue,
        })),
      warnings: gaps
        .filter((g) => g.severity === "warning")
        .map((g) => ({
          field: g.field,
          message: g.message,
          inputValue: g.inputValue,
          outputValue: g.outputValue,
        })),
      info: gaps
        .filter((g) => g.severity === "info")
        .map((g) => ({
          field: g.field,
          message: g.message,
        })),
      summary: {
        errorCount: gaps.filter((g) => g.severity === "error").length,
        warningCount: gaps.filter((g) => g.severity === "warning").length,
        infoCount: gaps.filter((g) => g.severity === "info").length,
        status: gaps.some((g) => g.severity === "error")
          ? "error"
          : gaps.some((g) => g.severity === "warning")
            ? "warning"
            : "ok",
      },
    };
    console.log(JSON.stringify(output, null, 2));

    // Still save EML if requested
    if (options.saveEml) {
      const emlPath = filePath.replace(/\.msg$/i, ".eml");
      fs.writeFileSync(emlPath, emlContent);
      console.error(colorize(`EML saved to: ${emlPath}`, "green"));
    }

    return;
  }

  printHeader("MSG TO EML DEBUG ANALYSIS");
  console.log(`  File: ${colorize(path.basename(filePath), "bold")}`);
  console.log(`  Size: ${buffer.length} bytes`);

  // ─────────────────────────────────────────────────────────────────
  // INPUT: Raw MSG Data
  // ─────────────────────────────────────────────────────────────────
  printHeader("INPUT: Raw MSG Properties");

  printSection("Basic Fields");
  printField("Subject", rawData.subject);
  printField("Message Class", rawData.messageClass);
  printField("Delivery Time", rawData.deliveryTime);

  printSection("Sender");
  printField("Name", rawData.senderName);
  printField("Email", rawData.senderEmail);
  printField("SMTP", rawData.senderSmtp);

  printSection("Body Content");
  printField("Plain Text", rawData.body ? `${rawData.body.length} chars` : undefined);
  printField("HTML", rawData.bodyHtml ? `${rawData.bodyHtml.length} chars` : undefined);
  printField("Compressed RTF", rawData.hasCompressedRtf ? "present" : undefined);

  printSection("Recipients & Attachments");
  printField("Recipients", rawData.recipientCount);
  printField("Attachments", rawData.attachmentCount);
  printField("Embedded Messages", rawData.embeddedMessageCount);

  printSection("Message Headers");
  printField("Message-ID", rawData.messageId);
  printField("In-Reply-To", rawData.inReplyTo);
  printField("References", rawData.references);
  printField("Reply-To", rawData.replyTo);
  printField("Priority", rawData.priority);
  printField("Importance", rawData.importance);
  printField("Read Receipt", rawData.readReceiptRequested);
  printField("Delivery Receipt", rawData.deliveryReceiptRequested);

  if (rawData.messageClass?.startsWith("IPM.Appointment")) {
    printSection("Calendar Properties");
    printField("Start Time", rawData.appointmentStart);
    printField("End Time", rawData.appointmentEnd);
    printField("Location", rawData.location);
    printField("To Attendees", rawData.toAttendees);
    printField("Cc Attendees", rawData.ccAttendees);
  }

  // ─────────────────────────────────────────────────────────────────
  // OUTPUT: Parsed Structure
  // ─────────────────────────────────────────────────────────────────
  printHeader("OUTPUT: Parsed MSG Structure");

  printSection("Basic Fields");
  printField("Subject", parsed.subject);
  printField("From", parsed.from);
  printField("Date", parsed.date);

  printSection("Body");
  printField("Plain Text", parsed.body ? `${parsed.body.length} chars` : undefined);
  printField("HTML", parsed.bodyHtml ? `${parsed.bodyHtml.length} chars` : undefined);

  printSection("Recipients");
  printRecipients(parsed.recipients);

  printSection("Attachments");
  printAttachments(parsed.attachments);

  if (parsed.headers && Object.keys(parsed.headers).length > 0) {
    printSection("Message Headers");
    printField("Message-ID", parsed.headers.messageId);
    printField("In-Reply-To", parsed.headers.inReplyTo);
    printField("References", parsed.headers.references);
    printField("Reply-To", parsed.headers.replyTo);
    printField("Priority (X-Priority)", parsed.headers.priority);
    printField("Disposition-Notification-To", parsed.headers.dispositionNotificationTo);
    printField("Return-Receipt-To", parsed.headers.returnReceiptTo);
  }

  if (parsed.calendarEvent) {
    printSection("Calendar Event");
    printField("Start Time", parsed.calendarEvent.startTime);
    printField("End Time", parsed.calendarEvent.endTime);
    printField("Location", parsed.calendarEvent.location);
    printField("Organizer", parsed.calendarEvent.organizer);
    printField("Attendees", parsed.calendarEvent.attendees.join(", ") || undefined);
  }

  // ─────────────────────────────────────────────────────────────────
  // EML OUTPUT
  // ─────────────────────────────────────────────────────────────────
  printHeader("EML OUTPUT");

  const emlLines = emlContent.split("\r\n");
  const headerEndIndex = emlLines.findIndex((line) => line === "");
  const emlHeaders = emlLines.slice(0, headerEndIndex).join("\n");

  printSection("EML Headers");
  console.log(colorize(emlHeaders, "dim"));

  if (options.fullEml) {
    printSection("Full EML Content");
    console.log(colorize(emlContent, "dim"));
  }

  printSection("EML Stats");
  printField("Total Size", `${emlContent.length} chars`);
  printField("Line Count", emlLines.length);
  printField("Boundary Count", (emlContent.match(/boundary=/g) || []).length);

  // Save EML if requested
  if (options.saveEml) {
    const emlPath = filePath.replace(/\.msg$/i, ".eml");
    fs.writeFileSync(emlPath, emlContent);
    console.log();
    console.log(`  ${colorize("✓", "green")} EML saved to: ${colorize(emlPath, "cyan")}`);
  }

  // ─────────────────────────────────────────────────────────────────
  // GAP ANALYSIS
  // ─────────────────────────────────────────────────────────────────
  printHeader("GAP ANALYSIS");

  printGaps(gaps);

  // ─────────────────────────────────────────────────────────────────
  // SUMMARY
  // ─────────────────────────────────────────────────────────────────
  printHeader("SUMMARY");

  const errors = gaps.filter((g) => g.severity === "error").length;
  const warnings = gaps.filter((g) => g.severity === "warning").length;

  if (errors === 0 && warnings === 0) {
    console.log(`  ${colorize("✓ Conversion looks good!", "green")}`);
  } else if (errors === 0) {
    console.log(`  ${colorize("⚠ Conversion completed with warnings", "yellow")}`);
  } else {
    console.log(`  ${colorize("✗ Conversion has issues that need attention", "red")}`);
  }

  console.log();
}

main().catch((err) => {
  console.error(colorize(`Error: ${err.message}`, "red"));
  if (err.stack) {
    console.error(colorize(err.stack, "dim"));
  }
  process.exit(1);
});
