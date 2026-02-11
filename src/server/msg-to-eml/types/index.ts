export interface Attachment {
  fileName: string;
  content: Uint8Array;
  contentType: string;
  contentId?: string;
  isEmbeddedMessage?: boolean;
}

export interface ParsedRecipient {
  name: string;
  email: string;
  type: "to" | "cc" | "bcc";
}

export interface MessageHeaders {
  messageId?: string;
  inReplyTo?: string;
  references?: string;
  replyTo?: string;
  priority?: number; // 1-5 scale (1=highest, 5=lowest)
  dispositionNotificationTo?: string; // For read receipt requests
  returnReceiptTo?: string; // For delivery receipt requests
  transportMessageHeaders?: string; // Original transport headers (Received, DKIM, SPF, etc.)
  threadIndex?: string; // Base64-encoded thread index for email threading (Thread-Index header)
  threadTopic?: string; // Conversation topic/subject for threading (Thread-Topic header)
}

export interface CalendarEvent {
  startTime: Date;
  endTime: Date;
  location?: string;
  organizer?: string;
  attendees: string[];
}

export interface ParsedMsg {
  subject: string;
  from: string;
  recipients: ParsedRecipient[];
  date: Date;
  body: string;
  bodyHtml?: string;
  attachments: Attachment[];
  headers?: MessageHeaders;
  calendarEvent?: CalendarEvent;
}
