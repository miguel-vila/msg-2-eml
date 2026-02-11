/**
 * Gets the recipient type string from MAPI recipient type number.
 * MAPI_TO = 1, MAPI_CC = 2, MAPI_BCC = 3
 */
export function getRecipientType(type: number | undefined): "to" | "cc" | "bcc" {
  if (type === 2) return "cc";
  if (type === 3) return "bcc";
  return "to";
}

/**
 * Maps MAPI priority and importance values to X-Priority scale (1-5).
 * PidTagPriority: -1 (non-urgent), 0 (normal), 1 (urgent)
 * PidTagImportance: 0 (low), 1 (normal), 2 (high)
 * X-Priority: 1 (highest), 2 (high), 3 (normal), 4 (low), 5 (lowest)
 */
export function mapToXPriority(priority: number | undefined, importance: number | undefined): number | undefined {
  // Prefer PidTagPriority if available
  if (priority !== undefined) {
    if (priority === 1) return 1; // urgent -> highest
    if (priority === 0) return 3; // normal -> normal
    if (priority === -1) return 5; // non-urgent -> lowest
  }
  // Fall back to PidTagImportance
  if (importance !== undefined) {
    if (importance === 2) return 1; // high -> highest
    if (importance === 1) return 3; // normal -> normal
    if (importance === 0) return 5; // low -> lowest
  }
  return undefined;
}
