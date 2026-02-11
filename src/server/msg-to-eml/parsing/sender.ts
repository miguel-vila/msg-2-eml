import {
  type Msg,
  PidTagSenderEmailAddress,
  PidTagSenderName,
  PidTagSenderSmtpAddress,
  PidTagSentRepresentingEmailAddress,
  PidTagSentRepresentingName,
  PidTagSentRepresentingSmtpAddress,
} from "msg-parser";
import { encodeDisplayName } from "../encoding/index.js";

/**
 * Information about the actual sender (the person who sent the message)
 */
export interface SenderInfo {
  email: string | undefined;
  name: string | undefined;
}

/**
 * Result of extracting sender information, including "on behalf of" scenario detection
 */
export interface SenderResult {
  /** The "From" address - the person being represented (or the sender if no delegation) */
  from: SenderInfo;
  /** The "Sender" address - only set when sender is acting on behalf of another user */
  sender: SenderInfo | undefined;
  /** True if this is an "on behalf of" scenario */
  isOnBehalfOf: boolean;
}

/**
 * Extracts the actual sender's email (PidTagSender* properties)
 */
function extractActualSenderEmail(msg: Msg): string | undefined {
  // Prefer SMTP address over X500/Exchange address
  return msg.getProperty<string>(PidTagSenderSmtpAddress) || msg.getProperty<string>(PidTagSenderEmailAddress);
}

/**
 * Extracts the actual sender's name (PidTagSenderName)
 */
function extractActualSenderName(msg: Msg): string | undefined {
  return msg.getProperty<string>(PidTagSenderName);
}

/**
 * Extracts the represented sender's email (PidTagSentRepresenting* properties)
 */
function extractRepresentedSenderEmail(msg: Msg): string | undefined {
  // Prefer SMTP address over X500/Exchange address
  return (
    msg.getProperty<string>(PidTagSentRepresentingSmtpAddress) ||
    msg.getProperty<string>(PidTagSentRepresentingEmailAddress)
  );
}

/**
 * Extracts the represented sender's name (PidTagSentRepresentingName)
 */
function extractRepresentedSenderName(msg: Msg): string | undefined {
  return msg.getProperty<string>(PidTagSentRepresentingName);
}

/**
 * Parses the From: header from raw transport message headers string.
 * Returns the parsed email and display name, or undefined if not found.
 *
 * Handles formats like:
 *   From: user@example.com
 *   From: "Display Name" <user@example.com>
 *   From: Display Name <user@example.com>
 */
export function parseFromTransportHeaders(transportHeaders: string): SenderInfo | undefined {
  // Normalize line endings and unfold continuation lines
  const normalized = transportHeaders.replace(/\r?\n/g, "\n");
  const lines = normalized.split("\n");

  let fromValue = "";
  let foundFrom = false;

  for (const line of lines) {
    if (!foundFrom) {
      // Match "From:" header (case-insensitive)
      const match = line.match(/^from:\s*(.*)/i);
      if (match) {
        fromValue = match[1];
        foundFrom = true;
      }
    } else if (line.startsWith(" ") || line.startsWith("\t")) {
      // Continuation line (folded header)
      fromValue += ` ${line.trim()}`;
    } else {
      // Next header starts, stop collecting
      break;
    }
  }

  if (!foundFrom || !fromValue.trim()) {
    return undefined;
  }

  fromValue = fromValue.trim();

  // Try to parse "Name <email>" or "<email>" format
  const angleMatch = fromValue.match(/^(.*?)\s*<([^>]+)>\s*$/);
  if (angleMatch) {
    const rawName = angleMatch[1].replace(/^["']|["']$/g, "").trim();
    const email = angleMatch[2].trim();
    return {
      email,
      name: rawName || undefined,
    };
  }

  // Check if the value looks like a bare email address
  if (fromValue.includes("@")) {
    return {
      email: fromValue,
      name: undefined,
    };
  }

  // Only a display name, no email
  return {
    email: undefined,
    name: fromValue,
  };
}

/**
 * Normalizes an email address for comparison (lowercase, trim whitespace)
 */
function normalizeEmail(email: string | undefined): string {
  return (email || "").toLowerCase().trim();
}

/**
 * Determines if this is an "on behalf of" scenario by comparing sender and represented sender.
 * Returns true if the actual sender differs from the represented sender.
 */
function isOnBehalfOfScenario(actual: SenderInfo, represented: SenderInfo): boolean {
  // If no actual sender info, not an "on behalf of" scenario
  if (!actual.email && !actual.name) {
    return false;
  }

  // If no represented sender info, not an "on behalf of" scenario
  if (!represented.email && !represented.name) {
    return false;
  }

  // Compare emails if both are available
  if (actual.email && represented.email) {
    return normalizeEmail(actual.email) !== normalizeEmail(represented.email);
  }

  // Compare names if only names are available
  if (actual.name && represented.name && !actual.email && !represented.email) {
    return actual.name.trim() !== represented.name.trim();
  }

  // If one has email and the other only has name, they're different
  if ((actual.email && !represented.email) || (!actual.email && represented.email)) {
    return true;
  }

  return false;
}

/**
 * Extracts complete sender information, detecting "on behalf of" scenarios.
 *
 * In Exchange/Outlook, when a user sends mail "on behalf of" another user:
 * - PidTagSender* contains the actual sender (the person who sent the message)
 * - PidTagSentRepresenting* contains the person being represented (shows in "From")
 *
 * RFC 5322 handles this with:
 * - "From" header: The person being represented (author)
 * - "Sender" header: The actual sender (when different from From)
 */
export function extractSenderInfo(msg: Msg, transportMessageHeaders?: string): SenderResult {
  const actualSender: SenderInfo = {
    email: extractActualSenderEmail(msg),
    name: extractActualSenderName(msg),
  };

  const representedSender: SenderInfo = {
    email: extractRepresentedSenderEmail(msg),
    name: extractRepresentedSenderName(msg),
  };

  const isOnBehalf = isOnBehalfOfScenario(actualSender, representedSender);

  if (isOnBehalf) {
    // "On behalf of" scenario: From = represented, Sender = actual
    return {
      from: representedSender,
      sender: actualSender,
      isOnBehalfOf: true,
    };
  }

  // Normal scenario: use best available sender info (prefer actual, fall back to represented)
  let bestEmail = actualSender.email || representedSender.email;
  let bestName = actualSender.name || representedSender.name;

  // Fallback: parse From: header from transport message headers when no MAPI sender properties are available
  if (!bestEmail && !bestName && transportMessageHeaders) {
    const fromTransport = parseFromTransportHeaders(transportMessageHeaders);
    if (fromTransport) {
      bestEmail = fromTransport.email;
      bestName = fromTransport.name;
    }
  }

  return {
    from: { email: bestEmail, name: bestName },
    sender: undefined,
    isOnBehalfOf: false,
  };
}

// Legacy functions for backward compatibility
export function extractSenderEmail(msg: Msg): string | undefined {
  const result = extractSenderInfo(msg);
  return result.from.email;
}

export function extractSenderName(msg: Msg): string | undefined {
  const result = extractSenderInfo(msg);
  return result.from.name;
}

export function formatSender(email: string | undefined, name: string | undefined): string {
  if (name && email && name !== email) {
    const encodedName = encodeDisplayName(name);
    return `${encodedName} <${email}>`;
  }
  return email || name || "unknown@unknown.com";
}
