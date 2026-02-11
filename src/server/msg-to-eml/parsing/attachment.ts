import {
  type EmbeddedMessage,
  type Msg,
  type Attachment as MsgAttachment,
  PidTagAttachContentId,
  PidTagAttachFilename,
  PidTagAttachLongFilename,
  PidTagAttachMimeTag,
} from "msg-parser";
import { getMimeType } from "../mime/index.js";
import type { Attachment } from "../types/index.js";

export function parseAttachment(attachment: MsgAttachment): Attachment | null {
  const fileName =
    attachment.getProperty<string>(PidTagAttachLongFilename) || attachment.getProperty<string>(PidTagAttachFilename);

  if (!fileName) return null;

  const content = attachment.content();
  if (!content || content.length === 0) return null;

  const mimeTag = attachment.getProperty<string>(PidTagAttachMimeTag);
  const contentType = mimeTag || getMimeType(fileName);
  const contentId = attachment.getProperty<string>(PidTagAttachContentId);

  return {
    fileName,
    content: new Uint8Array(content),
    contentType,
    contentId: contentId || undefined,
  };
}

/**
 * Parses an embedded MSG file (forwarded email, attachment) and converts it to an EML attachment.
 * Uses msg.extractEmbeddedMessage() to get a full Msg object, then recursively converts to EML.
 */
export function parseEmbeddedMessage(
  msg: Msg,
  embeddedMessage: EmbeddedMessage,
  msgToEmlFromMsg: (msg: Msg) => string,
): Attachment {
  // Get filename from the embedded message attachment properties
  const fileName =
    embeddedMessage.getProperty<string>(PidTagAttachLongFilename) ||
    embeddedMessage.getProperty<string>(PidTagAttachFilename) ||
    "embedded.eml";

  // Ensure the filename has .eml extension
  const emlFileName = fileName.toLowerCase().endsWith(".eml")
    ? fileName
    : fileName.replace(/\.msg$/i, ".eml") || `${fileName}.eml`;

  // Extract the embedded message as a full Msg object
  const extractedMsg = msg.extractEmbeddedMessage(embeddedMessage);

  // Recursively convert to EML using the provided conversion function
  const emlContent = msgToEmlFromMsg(extractedMsg);

  return {
    fileName: emlFileName,
    content: new TextEncoder().encode(emlContent),
    contentType: "message/rfc822",
    isEmbeddedMessage: true,
  };
}
