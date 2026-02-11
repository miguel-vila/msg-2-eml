import {
  type Recipient as MsgRecipient,
  PidTagDisplayName,
  PidTagEmailAddress,
  PidTagRecipientType,
  PidTagSmtpAddress,
} from "msg-parser";
import { getRecipientType } from "../mime/index.js";
import type { ParsedRecipient } from "../types/index.js";

export function parseRecipient(recipient: MsgRecipient): ParsedRecipient {
  const name = recipient.getProperty<string>(PidTagDisplayName) || "";
  // Prefer SMTP address over X500/Exchange address
  const smtpAddress = recipient.getProperty<string>(PidTagSmtpAddress);
  const emailAddress = recipient.getProperty<string>(PidTagEmailAddress);
  const email = smtpAddress || emailAddress || name;
  const type = recipient.getProperty<number>(PidTagRecipientType);
  return { name, email, type: getRecipientType(type) };
}
