export { parseAttachment, parseEmbeddedMessage } from "./attachment.js";
export { mapSensitivity, parseMsg, parseMsgFromMsg } from "./msg.js";
export { parseRecipient } from "./recipient.js";
export {
  extractSenderEmail,
  extractSenderInfo,
  extractSenderName,
  formatSender,
  parseFromTransportHeaders,
  type SenderInfo,
  type SenderResult,
} from "./sender.js";
