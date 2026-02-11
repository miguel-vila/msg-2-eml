#!/usr/bin/env npx tsx
/**
 * Script to analyze all MSG files in ./samples and update the gap-analysis-tracker.json
 * with warnings and errors from each conversion.
 */

import * as fs from "node:fs";
import * as path from "node:path";
import { Msg } from "msg-parser";
import { msgToEml, parseMsg } from "../src/server/msg-to-eml.js";

// Import the gap analysis types and functions from debug-conversion
import {
  PidTagBody,
  PidTagBodyHtml,
  PidTagImportance,
  PidTagInReplyToId,
  PidTagInternetMessageId,
  PidTagInternetReferences,
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
  PidTagMessageClass,
} from "msg-parser";
import type { ParsedMsg } from "../src/server/msg-to-eml/types/index.js";

interface Gap {
  field: string;
  severity: "warning" | "error" | "info";
  message: string;
  inputValue?: unknown;
  outputValue?: unknown;
}

interface TrackerFile {
  filename: string;
  status: string;
  processedAt: string | null;
  gapFound: boolean;
  gapDescription: string | null;
  testsAdded: string[];
  commitHash: string | null;
  gaps?: {
    warnings: Array<{ field: string; message: string }>;
    errors: Array<{ field: string; message: string }>;
  };
}

interface Tracker {
  description: string;
  lastUpdated: string;
  summary: {
    totalFiles: number;
    processed: number;
    pending: number;
    gapsFound: number;
  };
  files: TrackerFile[];
}

function safeGetProperty<T>(msg: Msg, prop: unknown): T | undefined {
  try {
    return msg.getProperty<T>(prop);
  } catch {
    return undefined;
  }
}

function analyzeGaps(msg: Msg, parsed: ParsedMsg, emlContent: string): Gap[] {
  const gaps: Gap[] = [];

  // Extract raw data
  const subject = safeGetProperty<string>(msg, PidTagSubject);
  const senderEmail = safeGetProperty<string>(msg, PidTagSenderEmailAddress);
  const senderSmtp = safeGetProperty<string>(msg, PidTagSenderSmtpAddress);
  const senderName = safeGetProperty<string>(msg, PidTagSenderName);
  const body = safeGetProperty<string>(msg, PidTagBody);
  const bodyHtml = safeGetProperty<string>(msg, PidTagBodyHtml);
  const compressedRtf = safeGetProperty<number[]>(msg, PidTagRtfCompressed);
  const hasCompressedRtf = !!(compressedRtf && compressedRtf.length > 0);
  const deliveryTime = safeGetProperty<Date>(msg, PidTagMessageDeliveryTime);
  const messageId = safeGetProperty<string>(msg, PidTagInternetMessageId);
  const inReplyTo = safeGetProperty<string>(msg, PidTagInReplyToId);
  const references = safeGetProperty<string>(msg, PidTagInternetReferences);
  const replyTo = safeGetProperty<string>(msg, PidTagReplyRecipientNames);
  const priority = safeGetProperty<number>(msg, PidTagPriority);
  const importance = safeGetProperty<number>(msg, PidTagImportance);
  const readReceiptRequested = safeGetProperty<boolean>(msg, PidTagReadReceiptRequested);
  const deliveryReceiptRequested = safeGetProperty<boolean>(msg, PidTagOriginatorDeliveryReportRequested);
  const messageClass = safeGetProperty<string>(msg, PidTagMessageClass);

  let recipientCount = 0;
  let attachmentCount = 0;
  let embeddedMessageCount = 0;
  try { recipientCount = msg.recipients().length; } catch { /* ignore */ }
  try { attachmentCount = msg.attachments().length; } catch { /* ignore */ }
  try { embeddedMessageCount = msg.embeddedMessages().length; } catch { /* ignore */ }

  // Check subject
  if (!subject && parsed.subject === "(No Subject)") {
    gaps.push({
      field: "subject",
      severity: "info",
      message: "No subject in source MSG, using default",
    });
  }

  // Check sender
  if (!senderEmail && !senderSmtp) {
    gaps.push({
      field: "from",
      severity: "warning",
      message: "No sender email address found in MSG properties",
    });
  }

  // Check if From is using X500 address instead of SMTP
  const fromUsesX500 = parsed.from.includes("/O=EXCHANGELABS") || parsed.from.includes("/o=ExchangeLabs");
  if (fromUsesX500 && senderSmtp) {
    gaps.push({
      field: "from",
      severity: "warning",
      message: "From uses X500 address but SMTP address is available",
    });
  }

  // Check body content
  if (!body && !bodyHtml && !hasCompressedRtf) {
    gaps.push({
      field: "body",
      severity: "warning",
      message: "No body content found (plain text, HTML, or RTF)",
    });
  } else if (!body && !bodyHtml && hasCompressedRtf) {
    if (!parsed.body && !parsed.bodyHtml) {
      gaps.push({
        field: "body",
        severity: "error",
        message: "Failed to extract body from compressed RTF",
      });
    }
  }

  // Check date
  if (!deliveryTime) {
    gaps.push({
      field: "date",
      severity: "warning",
      message: "No delivery time in MSG, using current date",
    });
  }

  // Check recipients
  if (recipientCount === 0) {
    gaps.push({
      field: "recipients",
      severity: "warning",
      message: "No recipients found in MSG",
    });
  } else if (parsed.recipients.length !== recipientCount) {
    gaps.push({
      field: "recipients",
      severity: "error",
      message: `Recipient count mismatch: MSG has ${recipientCount}, parsed ${parsed.recipients.length}`,
    });
  }

  // Check for recipients without emails
  const recipientsWithoutEmail = parsed.recipients.filter((r) => !r.email);
  if (recipientsWithoutEmail.length > 0) {
    gaps.push({
      field: "recipients",
      severity: "warning",
      message: `${recipientsWithoutEmail.length} recipient(s) without email address`,
    });
  }

  // Check for recipients using X500 addresses
  const recipientsWithX500 = parsed.recipients.filter(
    (r) => r.email && (r.email.includes("/O=EXCHANGELABS") || r.email.includes("/o=ExchangeLabs") || r.email.startsWith("/"))
  );
  if (recipientsWithX500.length > 0) {
    gaps.push({
      field: "recipients",
      severity: "warning",
      message: `${recipientsWithX500.length} recipient(s) using X500/Exchange address instead of SMTP`,
    });
  }

  // Check attachments
  const totalInputAttachments = attachmentCount + embeddedMessageCount;
  if (parsed.attachments.length !== totalInputAttachments) {
    gaps.push({
      field: "attachments",
      severity: "warning",
      message: `Attachment count: ${attachmentCount} regular + ${embeddedMessageCount} embedded = ${totalInputAttachments} total, parsed ${parsed.attachments.length}`,
    });
  }

  // Check for attachments without content
  const emptyAttachments = parsed.attachments.filter((a) => a.content.length === 0);
  if (emptyAttachments.length > 0) {
    gaps.push({
      field: "attachments",
      severity: "error",
      message: `${emptyAttachments.length} attachment(s) with empty content`,
    });
  }

  // Check message headers
  if (messageId && !parsed.headers?.messageId) {
    gaps.push({
      field: "headers.messageId",
      severity: "error",
      message: "Message-ID present in MSG but not parsed",
    });
  }

  if (inReplyTo && !parsed.headers?.inReplyTo) {
    gaps.push({
      field: "headers.inReplyTo",
      severity: "error",
      message: "In-Reply-To present in MSG but not parsed",
    });
  }

  if (references && !parsed.headers?.references) {
    gaps.push({
      field: "headers.references",
      severity: "error",
      message: "References present in MSG but not parsed",
    });
  }

  if (replyTo && !parsed.headers?.replyTo) {
    gaps.push({
      field: "headers.replyTo",
      severity: "error",
      message: "Reply-To present in MSG but not parsed",
    });
  }

  // Check priority
  if ((priority !== undefined || importance !== undefined) && !parsed.headers?.priority) {
    gaps.push({
      field: "headers.priority",
      severity: "warning",
      message: "Priority/Importance in MSG but not mapped",
    });
  }

  // Check read receipt
  if (readReceiptRequested && !parsed.headers?.dispositionNotificationTo) {
    gaps.push({
      field: "headers.dispositionNotificationTo",
      severity: "warning",
      message: "Read receipt requested but not set in parsed headers",
    });
  }

  // Check delivery receipt
  if (deliveryReceiptRequested && !parsed.headers?.returnReceiptTo) {
    gaps.push({
      field: "headers.returnReceiptTo",
      severity: "warning",
      message: "Delivery receipt requested but not set in parsed headers",
    });
  }

  // Check calendar event
  const isCalendar = messageClass?.startsWith("IPM.Appointment");
  if (isCalendar && !parsed.calendarEvent) {
    gaps.push({
      field: "calendarEvent",
      severity: "error",
      message: "Calendar message detected but no calendar event parsed",
    });
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

  return gaps;
}

async function main(): Promise<void> {
  const projectDir = path.resolve(import.meta.dirname, "..");
  const samplesDir = path.join(projectDir, "samples");
  const trackerPath = path.join(projectDir, "scripts", "gap-analysis-tracker.json");

  // Read the tracker
  const tracker: Tracker = JSON.parse(fs.readFileSync(trackerPath, "utf-8"));

  console.log(`Processing ${tracker.files.length} files...\n`);

  let processed = 0;
  let filesWithGaps = 0;
  let totalErrors = 0;
  let totalWarnings = 0;

  for (const fileEntry of tracker.files) {
    const filePath = path.join(samplesDir, fileEntry.filename);

    if (!fs.existsSync(filePath)) {
      console.log(`⚠ File not found: ${fileEntry.filename}`);
      continue;
    }

    try {
      const buffer = fs.readFileSync(filePath);
      const arrayBuffer = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);

      const msg = Msg.fromUint8Array(new Uint8Array(arrayBuffer));
      const parsed = parseMsg(arrayBuffer);
      const emlContent = msgToEml(arrayBuffer);

      const gaps = analyzeGaps(msg, parsed, emlContent);

      const errors = gaps
        .filter((g) => g.severity === "error")
        .map((g) => ({ field: g.field, message: g.message }));

      const warnings = gaps
        .filter((g) => g.severity === "warning")
        .map((g) => ({ field: g.field, message: g.message }));

      fileEntry.gaps = { warnings, errors };
      fileEntry.status = "processed";
      fileEntry.processedAt = new Date().toISOString();
      fileEntry.gapFound = errors.length > 0 || warnings.length > 0;

      if (errors.length > 0 || warnings.length > 0) {
        filesWithGaps++;
        const gapSummary = [];
        if (errors.length > 0) gapSummary.push(`${errors.length} errors`);
        if (warnings.length > 0) gapSummary.push(`${warnings.length} warnings`);
        fileEntry.gapDescription = gapSummary.join(", ");
      }

      totalErrors += errors.length;
      totalWarnings += warnings.length;
      processed++;

      const status = errors.length > 0 ? "✗" : warnings.length > 0 ? "⚠" : "✓";
      const statusColor = errors.length > 0 ? "\x1b[31m" : warnings.length > 0 ? "\x1b[33m" : "\x1b[32m";
      console.log(`${statusColor}${status}\x1b[0m ${fileEntry.filename} (${errors.length}E/${warnings.length}W)`);

    } catch (err) {
      console.log(`\x1b[31m✗\x1b[0m ${fileEntry.filename} - Error: ${(err as Error).message}`);
      fileEntry.gaps = {
        warnings: [],
        errors: [{ field: "processing", message: (err as Error).message }]
      };
      fileEntry.status = "error";
      fileEntry.gapFound = true;
      fileEntry.gapDescription = `Processing error: ${(err as Error).message}`;
      totalErrors++;
      filesWithGaps++;
    }
  }

  // Update summary
  tracker.summary.processed = processed;
  tracker.summary.pending = tracker.files.length - processed;
  tracker.summary.gapsFound = filesWithGaps;
  tracker.lastUpdated = new Date().toISOString();

  // Write back
  fs.writeFileSync(trackerPath, JSON.stringify(tracker, null, 2) + "\n");

  console.log("\n" + "═".repeat(50));
  console.log("SUMMARY");
  console.log("═".repeat(50));
  console.log(`Files processed: ${processed}`);
  console.log(`Files with gaps: ${filesWithGaps}`);
  console.log(`Total errors: ${totalErrors}`);
  console.log(`Total warnings: ${totalWarnings}`);
  console.log(`\nTracker updated: ${trackerPath}`);
}

main().catch((err) => {
  console.error("Error:", err);
  process.exit(1);
});
