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
export function extractSenderInfo(msg: Msg): SenderResult {
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
  const bestEmail = actualSender.email || representedSender.email;
  const bestName = actualSender.name || representedSender.name;

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
