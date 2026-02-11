import {
  Msg,
  PidLidAppointmentEndWhole,
  PidLidAppointmentStartWhole,
  PidLidCategories,
  PidLidCcAttendeesString,
  PidLidLocation,
  PidLidToAttendeesString,
  PidTagBody,
  PidTagBodyHtml,
  PidTagConversationIndex,
  PidTagConversationTopic,
  PidTagImportance,
  PidTagInReplyToId,
  PidTagInternetMessageId,
  PidTagInternetReferences,
  PidTagListHelp,
  PidTagListSubscribe,
  PidTagListUnsubscribe,
  PidTagMessageClass,
  PidTagMessageDeliveryTime,
  PidTagOriginatorDeliveryReportRequested,
  PidTagPriority,
  PidTagReadReceiptRequested,
  PidTagReceivedByName,
  PidTagReceivedBySmtpAddress,
  PidTagReplyRecipientNames,
  PidTagRtfCompressed,
  PidTagSensitivity,
  PidTagSubject,
  PidTagTransportMessageHeaders,
} from "msg-parser";
import { isCalendarMessage, parseAttendeeString } from "../calendar/index.js";
import { mapToXPriority } from "../mime/index.js";
import { extractBodyFromRtf } from "../rtf/index.js";
import type { Attachment, CalendarEvent, MessageHeaders, ParsedMsg } from "../types/index.js";
import { parseAttachment, parseEmbeddedMessage } from "./attachment.js";
import { parseRecipient } from "./recipient.js";
import { extractSenderInfo, formatSender } from "./sender.js";

/**
 * Maps PidTagSensitivity value to Sensitivity header string.
 * Values: 0=Normal (returns undefined, header should be omitted),
 * 1='Personal', 2='Private', 3='Company-Confidential'
 */
export function mapSensitivity(
  sensitivityValue: number | undefined,
): "Personal" | "Private" | "Company-Confidential" | undefined {
  switch (sensitivityValue) {
    case 1:
      return "Personal";
    case 2:
      return "Private";
    case 3:
      return "Company-Confidential";
    default:
      // 0 (Normal) or undefined - omit the header
      return undefined;
  }
}

export function parseMsgFromMsg(msg: Msg, msgToEmlFromMsg: (msg: Msg) => string): ParsedMsg {
  const subject = msg.getProperty<string>(PidTagSubject) || "(No Subject)";
  let body = msg.getProperty<string>(PidTagBody) || "";
  let bodyHtml = msg.getProperty<string>(PidTagBodyHtml);
  const senderInfo = extractSenderInfo(msg);
  const senderEmail = senderInfo.from.email;
  const senderName = senderInfo.from.name;
  const deliveryTime = msg.getProperty<Date>(PidTagMessageDeliveryTime);

  // If body is empty, try to extract from compressed RTF
  if (!body && !bodyHtml) {
    const compressedRtf = msg.getProperty<number[]>(PidTagRtfCompressed);
    if (compressedRtf && compressedRtf.length > 0) {
      const rtfResult = extractBodyFromRtf(compressedRtf);
      if (rtfResult) {
        body = rtfResult.text;
        if (rtfResult.html) {
          bodyHtml = rtfResult.html;
        }
      }
    }
  }

  const from = formatSender(senderEmail, senderName);

  const recipients = msg.recipients().map(parseRecipient);

  // Parse regular attachments
  const regularAttachments = msg
    .attachments()
    .map(parseAttachment)
    .filter((a): a is Attachment => a !== null);

  // Parse embedded messages (forwarded emails, attached emails)
  const embeddedMessages = msg.embeddedMessages();
  const embeddedAttachments = embeddedMessages.map((embedded) => parseEmbeddedMessage(msg, embedded, msgToEmlFromMsg));

  // Combine regular attachments and embedded message attachments
  const attachments = [...regularAttachments, ...embeddedAttachments];

  // Extract additional message headers
  const messageId = msg.getProperty<string>(PidTagInternetMessageId);
  const inReplyTo = msg.getProperty<string>(PidTagInReplyToId);
  const references = msg.getProperty<string>(PidTagInternetReferences);
  const replyTo = msg.getProperty<string>(PidTagReplyRecipientNames);
  const priority = msg.getProperty<number>(PidTagPriority);
  const importance = msg.getProperty<number>(PidTagImportance);
  const xPriority = mapToXPriority(priority, importance);

  // Extract sensitivity level (PidTagSensitivity)
  // Values: 0=Normal (omit header), 1='Personal', 2='Private', 3='Company-Confidential'
  const sensitivityValue = msg.getProperty<number>(PidTagSensitivity);
  const sensitivity = mapSensitivity(sensitivityValue);

  const headers: MessageHeaders = {};
  if (messageId) headers.messageId = messageId;
  if (inReplyTo) headers.inReplyTo = inReplyTo;
  if (references) headers.references = references;
  if (replyTo) headers.replyTo = replyTo;
  if (xPriority !== undefined) headers.priority = xPriority;
  if (sensitivity) headers.sensitivity = sensitivity;

  // Check for read receipt and delivery receipt requests
  const readReceiptRequested = msg.getProperty<boolean>(PidTagReadReceiptRequested);
  const deliveryReceiptRequested = msg.getProperty<boolean>(PidTagOriginatorDeliveryReportRequested);

  // If read receipt is requested, use sender's email for Disposition-Notification-To header
  if (readReceiptRequested && senderEmail) {
    headers.dispositionNotificationTo = senderEmail;
  }

  // If delivery receipt is requested, use sender's email for Return-Receipt-To header
  if (deliveryReceiptRequested && senderEmail) {
    headers.returnReceiptTo = senderEmail;
  }

  // Extract original transport headers (Received, DKIM-Signature, SPF, Authentication-Results, etc.)
  const transportMessageHeaders = msg.getProperty<string>(PidTagTransportMessageHeaders);
  if (transportMessageHeaders) {
    headers.transportMessageHeaders = transportMessageHeaders;
  }

  // Extract conversation threading metadata
  const conversationIndex = msg.getProperty<number[]>(PidTagConversationIndex);
  if (conversationIndex && conversationIndex.length > 0) {
    // Convert binary array to base64 for Thread-Index header
    const bytes = new Uint8Array(conversationIndex);
    const binary = String.fromCharCode(...bytes);
    headers.threadIndex = btoa(binary);
  }

  const conversationTopic = msg.getProperty<string>(PidTagConversationTopic);
  if (conversationTopic) {
    headers.threadTopic = conversationTopic;
  }

  // Extract received-by information for the final recipient
  const receivedBySmtpAddress = msg.getProperty<string>(PidTagReceivedBySmtpAddress);
  if (receivedBySmtpAddress) {
    headers.receivedByEmail = receivedBySmtpAddress;
  }

  const receivedByName = msg.getProperty<string>(PidTagReceivedByName);
  if (receivedByName) {
    headers.receivedByName = receivedByName;
  }

  // Extract mailing list headers (RFC 2369)
  const listHelp = msg.getProperty<string>(PidTagListHelp);
  if (listHelp) {
    headers.listHelp = listHelp;
  }

  const listSubscribe = msg.getProperty<string>(PidTagListSubscribe);
  if (listSubscribe) {
    headers.listSubscribe = listSubscribe;
  }

  const listUnsubscribe = msg.getProperty<string>(PidTagListUnsubscribe);
  if (listUnsubscribe) {
    headers.listUnsubscribe = listUnsubscribe;
  }

  // Extract Outlook categories (PidLidCategories) for Keywords header
  // PidLidCategories uses named properties which may not be available in embedded messages,
  // so we wrap in try-catch to handle gracefully.
  try {
    const categories = msg.getProperty<string[]>(PidLidCategories);
    if (categories && categories.length > 0) {
      headers.keywords = categories;
    }
  } catch {
    // Named property mapping not available (e.g., embedded messages) - skip categories
  }

  // Check if this is a calendar appointment
  const messageClass = msg.getProperty<string>(PidTagMessageClass);
  let calendarEvent: CalendarEvent | undefined;

  if (isCalendarMessage(messageClass)) {
    const startTime = msg.getProperty<Date>(PidLidAppointmentStartWhole);
    const endTime = msg.getProperty<Date>(PidLidAppointmentEndWhole);

    if (startTime && endTime) {
      const location = msg.getProperty<string>(PidLidLocation);
      const toAttendees = msg.getProperty<string>(PidLidToAttendeesString);
      const ccAttendees = msg.getProperty<string>(PidLidCcAttendeesString);

      // Combine To and Cc attendees
      const attendees = [...parseAttendeeString(toAttendees), ...parseAttendeeString(ccAttendees)];

      // Use sender as organizer
      const organizer = senderEmail || senderName;

      calendarEvent = {
        startTime,
        endTime,
        location: location || undefined,
        organizer: organizer || undefined,
        attendees,
      };
    }
  }

  // Build the sender header if this is an "on behalf of" scenario
  const senderHeader =
    senderInfo.isOnBehalfOf && senderInfo.sender
      ? formatSender(senderInfo.sender.email, senderInfo.sender.name)
      : undefined;

  return {
    subject,
    from,
    sender: senderHeader,
    recipients,
    date: deliveryTime || new Date(),
    body,
    bodyHtml: bodyHtml || undefined,
    attachments,
    headers: Object.keys(headers).length > 0 ? headers : undefined,
    calendarEvent,
  };
}

export function parseMsg(buffer: ArrayBuffer, msgToEmlFromMsg: (msg: Msg) => string): ParsedMsg {
  const msg = Msg.fromUint8Array(new Uint8Array(buffer));
  return parseMsgFromMsg(msg, msgToEmlFromMsg);
}
