import {
  type Recipient as MsgRecipient,
  PidTagDisplayName,
  PidTagEmailAddress,
  PidTagRecipientType,
  PidTagSmtpAddress,
} from "msg-parser";
import { getRecipientType } from "../mime/index.js";
import type { ParsedRecipient } from "../types/index.js";

/**
 * Parsed address entry from a transport header (To:/Cc:/Bcc:)
 */
interface TransportAddress {
  name?: string;
  email?: string;
}

/**
 * Parsed recipient addresses from transport headers, grouped by type
 */
export interface TransportRecipients {
  to: TransportAddress[];
  cc: TransportAddress[];
  bcc: TransportAddress[];
}

/**
 * Parses a single address header value (e.g. the value portion of "To: ...") into
 * individual address entries. Handles formats:
 *   - user@example.com
 *   - "Display Name" <user@example.com>
 *   - Display Name <user@example.com>
 *   - Multiple addresses separated by commas
 */
function parseAddressHeaderValue(headerValue: string): TransportAddress[] {
  const addresses: TransportAddress[] = [];
  if (!headerValue.trim()) return addresses;

  // Split on commas that are not inside angle brackets or quotes
  const parts: string[] = [];
  let current = "";
  let inAngle = 0;
  let inQuote = false;

  for (const char of headerValue) {
    if (char === '"' && !inAngle) {
      inQuote = !inQuote;
    } else if (char === "<" && !inQuote) {
      inAngle++;
    } else if (char === ">" && !inQuote) {
      inAngle = Math.max(0, inAngle - 1);
    }

    if (char === "," && !inAngle && !inQuote) {
      parts.push(current.trim());
      current = "";
    } else {
      current += char;
    }
  }
  if (current.trim()) {
    parts.push(current.trim());
  }

  for (const part of parts) {
    const trimmed = part.trim();
    if (!trimmed) continue;

    // Try "Name <email>" or "<email>" format
    const angleMatch = trimmed.match(/^(.*?)\s*<([^>]+)>\s*$/);
    if (angleMatch) {
      const rawName = angleMatch[1].replace(/^["']|["']$/g, "").trim();
      const email = angleMatch[2].trim();
      addresses.push({ name: rawName || undefined, email });
      continue;
    }

    // Bare email address
    if (trimmed.includes("@")) {
      addresses.push({ email: trimmed, name: undefined });
      continue;
    }

    // Just a name, no email
    addresses.push({ name: trimmed, email: undefined });
  }

  return addresses;
}

/**
 * Extracts a single header's full value (handling folded lines) from transport headers.
 */
function extractHeaderValue(transportHeaders: string, headerName: string): string | undefined {
  const normalized = transportHeaders.replace(/\r?\n/g, "\n");
  const lines = normalized.split("\n");
  const pattern = new RegExp(`^${headerName}:\\s*(.*)`, "i");

  let value = "";
  let found = false;

  for (const line of lines) {
    if (!found) {
      const match = line.match(pattern);
      if (match) {
        value = match[1];
        found = true;
      }
    } else if (line.startsWith(" ") || line.startsWith("\t")) {
      value += ` ${line.trim()}`;
    } else {
      break;
    }
  }

  return found && value.trim() ? value.trim() : undefined;
}

/**
 * Parses To:, Cc:, and Bcc: headers from raw transport message headers.
 * Returns grouped recipient addresses by type.
 */
export function parseRecipientsFromTransportHeaders(transportHeaders: string): TransportRecipients {
  const toValue = extractHeaderValue(transportHeaders, "To");
  const ccValue = extractHeaderValue(transportHeaders, "Cc");
  const bccValue = extractHeaderValue(transportHeaders, "Bcc");

  return {
    to: toValue ? parseAddressHeaderValue(toValue) : [],
    cc: ccValue ? parseAddressHeaderValue(ccValue) : [],
    bcc: bccValue ? parseAddressHeaderValue(bccValue) : [],
  };
}

/**
 * Tries to resolve a recipient's email from transport header addresses by matching display name.
 * Returns the email if a match is found, undefined otherwise.
 */
function resolveEmailFromTransportHeaders(
  displayName: string,
  recipientType: "to" | "cc" | "bcc",
  transportRecipients: TransportRecipients,
): string | undefined {
  if (!displayName) return undefined;

  const normalizedName = displayName.toLowerCase().trim();

  // Search in the matching recipient type first, then fall back to all types
  const typedAddresses = transportRecipients[recipientType];
  const allAddresses = [...transportRecipients.to, ...transportRecipients.cc, ...transportRecipients.bcc];

  for (const addresses of [typedAddresses, allAddresses]) {
    for (const addr of addresses) {
      if (addr.email && addr.name && addr.name.toLowerCase().trim() === normalizedName) {
        return addr.email;
      }
    }
  }

  return undefined;
}

export function parseRecipient(recipient: MsgRecipient, transportRecipients?: TransportRecipients): ParsedRecipient {
  const name = recipient.getProperty<string>(PidTagDisplayName) || "";
  // Prefer SMTP address over X500/Exchange address
  const smtpAddress = recipient.getProperty<string>(PidTagSmtpAddress);
  const emailAddress = recipient.getProperty<string>(PidTagEmailAddress);
  const type = getRecipientType(recipient.getProperty<number>(PidTagRecipientType));

  let email: string | undefined = smtpAddress || emailAddress;

  // If no email found, try to resolve from transport headers by matching display name
  if (!email && name && transportRecipients) {
    email = resolveEmailFromTransportHeaders(name, type, transportRecipients);
  }

  return { name, email: email ?? name, type };
}
