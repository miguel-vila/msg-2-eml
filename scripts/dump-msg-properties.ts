#!/usr/bin/env npx tsx
/**
 * Dumps all properties from a MSG file to see what data is available
 * vs what we're currently extracting.
 */

import * as fs from "node:fs";
import * as path from "node:path";
import { Msg } from "msg-parser";
import * as KnownProperties from "msg-parser/dist/property/KnownProperties.js";

// Properties we currently extract (by property ID or name)
const EXTRACTED_PROPERTY_IDS = new Set([
  0x0037, // PidTagSubject
  0x1000, // PidTagBody
  0x1013, // PidTagBodyHtml
  0x1009, // PidTagRtfCompressed
  0x0E06, // PidTagMessageDeliveryTime
  0x001A, // PidTagMessageClass
  0x0C1F, // PidTagSenderEmailAddress
  0x5D01, // PidTagSenderSmtpAddress
  0x0C1A, // PidTagSenderName
  0x0065, // PidTagSentRepresentingEmailAddress
  0x5D02, // PidTagSentRepresentingSmtpAddress
  0x0042, // PidTagSentRepresentingName
  0x1035, // PidTagInternetMessageId
  0x1042, // PidTagInReplyToId
  0x1039, // PidTagInternetReferences
  0x0050, // PidTagReplyRecipientNames
  0x0026, // PidTagPriority
  0x0017, // PidTagImportance
  0x0029, // PidTagReadReceiptRequested
  0x0023, // PidTagOriginatorDeliveryReportRequested
  // Recipient
  0x3001, // PidTagDisplayName
  0x3003, // PidTagEmailAddress
  0x39FE, // PidTagSmtpAddress
  0x0C15, // PidTagRecipientType
  // Attachment
  0x3707, // PidTagAttachLongFilename
  0x3704, // PidTagAttachFilename
  0x370E, // PidTagAttachMimeTag
  0x3712, // PidTagAttachContentId
]);

// Named properties we extract (by LID)
const EXTRACTED_NAMED_LIDS = new Set([
  0x820D, // PidLidAppointmentStartWhole
  0x820E, // PidLidAppointmentEndWhole
  0x8208, // PidLidLocation
  0x823B, // PidLidToAttendeesString
  0x823C, // PidLidCcAttendeesString
]);

function formatValue(value: unknown): string {
  if (value === undefined || value === null) {
    return "(empty)";
  }
  if (value instanceof Date) {
    return value.toISOString();
  }
  if (value instanceof Uint8Array || ArrayBuffer.isView(value)) {
    return `[${value.byteLength} bytes]`;
  }
  if (Array.isArray(value)) {
    if (value.length > 100) {
      return `[${value.length} items]`;
    }
    if (value.length === 0) {
      return "[]";
    }
    if (typeof value[0] === "number") {
      return `[${value.length} bytes]`;
    }
    return JSON.stringify(value).substring(0, 100);
  }
  if (typeof value === "string") {
    if (value.length > 150) {
      return `"${value.substring(0, 150)}..." (${value.length} chars)`;
    }
    return `"${value}"`;
  }
  if (typeof value === "object") {
    return JSON.stringify(value).substring(0, 100);
  }
  return String(value);
}

function getPropertyName(propId: number): string | undefined {
  // Search through known properties
  for (const [name, prop] of Object.entries(KnownProperties)) {
    if (prop && typeof prop === "object" && "propertyId" in prop) {
      if ((prop as { propertyId: number }).propertyId === propId) {
        return name;
      }
    }
  }
  return undefined;
}

async function main(): Promise<void> {
  const args = process.argv.slice(2);

  if (args.length === 0) {
    console.error("Usage: npx tsx scripts/dump-msg-properties.ts <path-to-msg-file>");
    process.exit(1);
  }

  const filePath = path.resolve(args[0]);

  if (!fs.existsSync(filePath)) {
    console.error(`File not found: ${filePath}`);
    process.exit(1);
  }

  const buffer = fs.readFileSync(filePath);
  const arrayBuffer = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
  const msg = Msg.fromUint8Array(new Uint8Array(arrayBuffer));

  console.log("═".repeat(70));
  console.log("MSG PROPERTY DUMP");
  console.log("═".repeat(70));
  console.log(`File: ${path.basename(filePath)}`);
  console.log();

  // Get all streams to find property IDs
  const streams = msg.streams();
  const propertyStreams = streams.filter((s) => s.getDirectoryEntryName().startsWith("__substg1.0_"));

  const extracted: Array<{ id: string; name: string; value: string }> = [];
  const notExtracted: Array<{ id: string; name: string; value: string }> = [];

  // Parse stream names to get property IDs
  for (const stream of propertyStreams) {
    const streamName = stream.getDirectoryEntryName();
    // Stream name format: __substg1.0_XXXXYYYY where XXXX is property ID in hex
    const match = streamName.match(/__substg1\.0_([0-9A-F]{4})([0-9A-F]{4})/i);
    if (match) {
      const propId = parseInt(match[1], 16);
      const propType = parseInt(match[2], 16);
      const propName = getPropertyName(propId) || `Unknown_0x${propId.toString(16).toUpperCase().padStart(4, "0")}`;

      // Try to get the value using known properties
      let value: unknown;
      const knownProp = Object.values(KnownProperties).find(
        (p) => p && typeof p === "object" && "propertyId" in p && (p as { propertyId: number }).propertyId === propId
      );

      if (knownProp) {
        try {
          value = msg.getProperty(knownProp as Parameters<typeof msg.getProperty>[0]);
        } catch {
          value = "(error reading)";
        }
      } else {
        // Read raw stream data
        try {
          const data = stream.getStream();
          if (propType === 0x001f || propType === 0x001e) {
            // Unicode or ASCII string
            value = new TextDecoder("utf-16le").decode(data);
          } else if (propType === 0x0003) {
            // Integer
            value = data.length >= 4 ? new DataView(data.buffer).getInt32(0, true) : "(invalid)";
          } else if (propType === 0x000b) {
            // Boolean
            value = data.length >= 2 && new DataView(data.buffer).getInt16(0, true) !== 0;
          } else {
            value = `[${data.length} bytes, type 0x${propType.toString(16)}]`;
          }
        } catch {
          value = "(error reading stream)";
        }
      }

      const isExtracted = EXTRACTED_PROPERTY_IDS.has(propId);
      const formattedValue = formatValue(value);
      const idStr = `0x${propId.toString(16).toUpperCase().padStart(4, "0")}`;

      if (value !== undefined && value !== null && formattedValue !== "(empty)") {
        if (isExtracted) {
          extracted.push({ id: idStr, name: propName, value: formattedValue });
        } else {
          notExtracted.push({ id: idStr, name: propName, value: formattedValue });
        }
      }
    }
  }

  // Print extracted properties
  console.log("\x1b[32m▸ PROPERTIES WE EXTRACT (" + extracted.length + ")\x1b[0m");
  console.log("─".repeat(70));
  for (const { id, name, value } of extracted.sort((a, b) => a.name.localeCompare(b.name))) {
    console.log(`  \x1b[32m✓\x1b[0m [${id}] ${name}: ${value}`);
  }

  // Print NOT extracted properties (potential gaps)
  console.log();
  console.log("\x1b[33m▸ PROPERTIES WE DON'T EXTRACT (" + notExtracted.length + ")\x1b[0m");
  console.log("─".repeat(70));
  for (const { id, name, value } of notExtracted.sort((a, b) => a.name.localeCompare(b.name))) {
    console.log(`  \x1b[33m!\x1b[0m [${id}] ${name}: ${value}`);
  }

  // Print recipients
  console.log();
  console.log("\x1b[36m▸ RECIPIENTS\x1b[0m");
  console.log("─".repeat(70));
  const recipients = msg.recipients();
  for (let i = 0; i < recipients.length; i++) {
    const r = recipients[i];
    console.log(`  Recipient ${i + 1}:`);
    const recipientStreams = r.streams();
    for (const stream of recipientStreams) {
      const streamName = stream.getDirectoryEntryName();
      const match = streamName.match(/__substg1\.0_([0-9A-F]{4})([0-9A-F]{4})/i);
      if (match) {
        const propId = parseInt(match[1], 16);
        const propType = parseInt(match[2], 16);
        const propName = getPropertyName(propId) || `Unknown_0x${propId.toString(16).toUpperCase()}`;

        let value: unknown;
        try {
          const data = stream.getStream();
          if (propType === 0x001f || propType === 0x001e) {
            value = new TextDecoder("utf-16le").decode(data);
          } else {
            value = `[${data.length} bytes]`;
          }
        } catch {
          value = "(error)";
        }

        if (value && formatValue(value) !== "(empty)") {
          const isExtracted = EXTRACTED_PROPERTY_IDS.has(propId);
          const marker = isExtracted ? "\x1b[32m✓\x1b[0m" : "\x1b[33m!\x1b[0m";
          console.log(`    ${marker} [0x${propId.toString(16).toUpperCase().padStart(4, "0")}] ${propName}: ${formatValue(value)}`);
        }
      }
    }
  }

  // Print attachments
  console.log();
  console.log("\x1b[36m▸ ATTACHMENTS\x1b[0m");
  console.log("─".repeat(70));
  const attachments = msg.attachments();
  for (let i = 0; i < attachments.length; i++) {
    const a = attachments[i];
    console.log(`  Attachment ${i + 1}:`);
    const attachmentStreams = a.streams();
    for (const stream of attachmentStreams) {
      const streamName = stream.getDirectoryEntryName();
      const match = streamName.match(/__substg1\.0_([0-9A-F]{4})([0-9A-F]{4})/i);
      if (match) {
        const propId = parseInt(match[1], 16);
        const propType = parseInt(match[2], 16);
        const propName = getPropertyName(propId) || `Unknown_0x${propId.toString(16).toUpperCase()}`;

        let value: unknown;
        try {
          const data = stream.getStream();
          if (propType === 0x001f || propType === 0x001e) {
            value = new TextDecoder("utf-16le").decode(data);
          } else {
            value = `[${data.length} bytes]`;
          }
        } catch {
          value = "(error)";
        }

        if (value && formatValue(value) !== "(empty)") {
          const isExtracted = EXTRACTED_PROPERTY_IDS.has(propId);
          const marker = isExtracted ? "\x1b[32m✓\x1b[0m" : "\x1b[33m!\x1b[0m";
          console.log(`    ${marker} [0x${propId.toString(16).toUpperCase().padStart(4, "0")}] ${propName}: ${formatValue(value)}`);
        }
      }
    }
  }

  // Summary
  console.log();
  console.log("═".repeat(70));
  console.log("SUMMARY");
  console.log("═".repeat(70));
  console.log(`  Properties with values: ${extracted.length + notExtracted.length}`);
  console.log(`  \x1b[32mExtracted: ${extracted.length}\x1b[0m`);
  console.log(`  \x1b[33mNot extracted: ${notExtracted.length}\x1b[0m`);

  if (notExtracted.length > 0) {
    console.log();
    console.log("\x1b[33mPotential gaps - properties with data we're not using:\x1b[0m");
    for (const { id, name } of notExtracted) {
      console.log(`  - [${id}] ${name}`);
    }
  }
}

main().catch((err) => {
  console.error("Error:", err);
  process.exit(1);
});
