import { type Recipient as MsgRecipient, PidTagDisplayName, PidTagEmailAddress, PidTagRecipientType } from "msg-parser";
import { getRecipientType } from "../mime/index.js";
import type { ParsedRecipient } from "../types/index.js";

export function parseRecipient(recipient: MsgRecipient): ParsedRecipient {
  const name = recipient.getProperty<string>(PidTagDisplayName) || "";
  const email = recipient.getProperty<string>(PidTagEmailAddress) || name;
  const type = recipient.getProperty<number>(PidTagRecipientType);
  return { name, email, type: getRecipientType(type) };
}
