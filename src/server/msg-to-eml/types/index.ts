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
  sensitivity?: "Personal" | "Private" | "Company-Confidential"; // Email sensitivity/confidentiality level (Sensitivity header)
  dispositionNotificationTo?: string; // For read receipt requests
  returnReceiptTo?: string; // For delivery receipt requests
  transportMessageHeaders?: string; // Original transport headers (Received, DKIM, SPF, etc.)
  threadIndex?: string; // Base64-encoded thread index for email threading (Thread-Index header)
  threadTopic?: string; // Conversation topic/subject for threading (Thread-Topic header)
  receivedByEmail?: string; // SMTP address of the final recipient (from PidTagReceivedBySmtpAddress)
  receivedByName?: string; // Display name of the final recipient (from PidTagReceivedByName)
  // Mailing list headers (RFC 2369)
  listHelp?: string; // List-Help header URL (from PidTagListHelp)
  listSubscribe?: string; // List-Subscribe header URL (from PidTagListSubscribe)
  listUnsubscribe?: string; // List-Unsubscribe header URL (from PidTagListUnsubscribe)
  // Categories / Keywords (RFC 5322)
  keywords?: string[]; // Outlook categories mapped to Keywords header (from PidLidCategories)
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
  /** RFC 5322 Sender header - only present when actual sender differs from From (on behalf of scenarios) */
  sender?: string;
  recipients: ParsedRecipient[];
  date: Date;
  body: string;
  bodyHtml?: string;
  attachments: Attachment[];
  headers?: MessageHeaders;
  calendarEvent?: CalendarEvent;
}
