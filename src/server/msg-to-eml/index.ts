// Types

// Calendar utilities
export {
  escapeICalText,
  foldICalLine,
  formatICalDateTime,
  generateVCalendar,
  isCalendarMessage,
  parseAttendeeString,
} from "./calendar/index.js";
// Main conversion functions
export {
  convertToEml,
  msgToEml,
  parseMsg,
} from "./converter/index.js";
// Encoding utilities
export {
  encodeBase64,
  encodeDisplayName,
  encodeQuotedPrintable,
  encodeRfc2047,
  encodeRfc2231,
  foldHeader,
  formatFilenameParams,
  isAscii,
} from "./encoding/index.js";
// MIME utilities
export {
  formatEmailDate,
  generateBoundary,
  getMimeType,
  getRecipientType,
  mapToXPriority,
} from "./mime/index.js";
// Parsing utilities
export {
  extractSenderEmail,
  extractSenderInfo,
  extractSenderName,
  formatSender,
  parseAttachment,
  parseEmbeddedMessage,
  parseRecipient,
  type SenderInfo,
  type SenderResult,
} from "./parsing/index.js";
export type { RtfBodyResult } from "./rtf/index.js";
// RTF utilities
export { extractBodyFromRtf } from "./rtf/index.js";
export type {
  Attachment,
  CalendarEvent,
  MessageHeaders,
  ParsedMsg,
  ParsedRecipient,
} from "./types/index.js";
