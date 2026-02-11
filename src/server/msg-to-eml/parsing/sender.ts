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

export function extractSenderEmail(msg: Msg): string | undefined {
  // Prefer SMTP addresses over X500/Exchange addresses
  const smtpAddress =
    msg.getProperty<string>(PidTagSenderSmtpAddress) || msg.getProperty<string>(PidTagSentRepresentingSmtpAddress);

  if (smtpAddress) {
    return smtpAddress;
  }

  // Fall back to email address properties (may be X500)
  const emailAddress =
    msg.getProperty<string>(PidTagSenderEmailAddress) || msg.getProperty<string>(PidTagSentRepresentingEmailAddress);

  return emailAddress || undefined;
}

export function extractSenderName(msg: Msg): string | undefined {
  return msg.getProperty<string>(PidTagSenderName) || msg.getProperty<string>(PidTagSentRepresentingName) || undefined;
}

export function formatSender(email: string | undefined, name: string | undefined): string {
  if (name && email && name !== email) {
    const encodedName = encodeDisplayName(name);
    return `${encodedName} <${email}>`;
  }
  return email || name || "unknown@unknown.com";
}
