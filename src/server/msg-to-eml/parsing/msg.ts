import {
  Msg,
  PidLidAppointmentEndWhole,
  PidLidAppointmentStartWhole,
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
  PidTagMessageClass,
  PidTagMessageDeliveryTime,
  PidTagOriginatorDeliveryReportRequested,
  PidTagPriority,
  PidTagReadReceiptRequested,
  PidTagReplyRecipientNames,
  PidTagRtfCompressed,
  PidTagSubject,
  PidTagTransportMessageHeaders,
} from "msg-parser";
import { isCalendarMessage, parseAttendeeString } from "../calendar/index.js";
import { mapToXPriority } from "../mime/index.js";
import { extractBodyFromRtf } from "../rtf/index.js";
import type { Attachment, CalendarEvent, MessageHeaders, ParsedMsg } from "../types/index.js";
import { parseAttachment, parseEmbeddedMessage } from "./attachment.js";
import { parseRecipient } from "./recipient.js";
import { extractSenderEmail, extractSenderName, formatSender } from "./sender.js";

export function parseMsgFromMsg(msg: Msg, msgToEmlFromMsg: (msg: Msg) => string): ParsedMsg {
  const subject = msg.getProperty<string>(PidTagSubject) || "(No Subject)";
  let body = msg.getProperty<string>(PidTagBody) || "";
  let bodyHtml = msg.getProperty<string>(PidTagBodyHtml);
  const senderEmail = extractSenderEmail(msg);
  const senderName = extractSenderName(msg);
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

  const headers: MessageHeaders = {};
  if (messageId) headers.messageId = messageId;
  if (inReplyTo) headers.inReplyTo = inReplyTo;
  if (references) headers.references = references;
  if (replyTo) headers.replyTo = replyTo;
  if (xPriority !== undefined) headers.priority = xPriority;

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

  return {
    subject,
    from,
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
