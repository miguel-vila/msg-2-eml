/**
 * RFC 5322 header folding.
 * Headers longer than 78 characters should be folded by inserting CRLF
 * followed by a space or tab (continuation).
 * This function is careful not to break:
 * - In the middle of encoded words (=?...?=)
 * - In the middle of email addresses (<...>)
 * - In the middle of quoted strings ("...")
 */
export function foldHeader(headerName: string, headerValue: string, maxLineLength: number = 78): string {
  const fullHeader = `${headerName}: ${headerValue}`;

  // If the header is already short enough, return as-is
  if (fullHeader.length <= maxLineLength) {
    return fullHeader;
  }

  const lines: string[] = [];
  let currentLine = `${headerName}:`;
  let remaining = headerValue;
  let isFirstSegment = true;

  while (remaining.length > 0) {
    // Calculate available space on current line (accounting for leading space)
    const _leadingSpace = isFirstSegment ? " " : "\t";
    const availableSpace = maxLineLength - currentLine.length - (isFirstSegment ? 1 : 0);

    if (isFirstSegment) {
      currentLine += " ";
    }

    if (remaining.length <= availableSpace) {
      // Remaining content fits on current line
      currentLine += remaining;
      remaining = "";
    } else {
      // Need to find a safe break point
      const breakPoint = findSafeBreakPoint(remaining, availableSpace);

      if (breakPoint <= 0) {
        // No safe break point found within available space
        // Either the first token is too long, or we need to include it anyway
        const forcedBreak = findFirstBreakPoint(remaining);
        if (forcedBreak > 0 && forcedBreak <= remaining.length) {
          currentLine += remaining.substring(0, forcedBreak).trimEnd();
          remaining = remaining.substring(forcedBreak).trimStart();
        } else {
          // Single token that can't be broken - include it anyway
          currentLine += remaining;
          remaining = "";
        }
      } else {
        currentLine += remaining.substring(0, breakPoint).trimEnd();
        remaining = remaining.substring(breakPoint).trimStart();
      }
    }

    if (remaining.length > 0) {
      lines.push(currentLine);
      currentLine = "\t"; // Use tab for continuation lines
      isFirstSegment = false;
    }
  }

  lines.push(currentLine);
  return lines.join("\r\n");
}

/**
 * Find the first possible break point (space or comma) in the string.
 */
function findFirstBreakPoint(str: string): number {
  for (let i = 0; i < str.length; i++) {
    if (str[i] === " " || str[i] === ",") {
      // For comma, include it in current segment
      return str[i] === "," ? i + 1 : i;
    }
  }
  return -1;
}

/**
 * Find a safe break point within the given max position.
 * Avoids breaking inside:
 * - Encoded words (=?...?=)
 * - Email addresses in angle brackets (<...>)
 * - Quoted strings ("...")
 */
function findSafeBreakPoint(str: string, maxPos: number): number {
  if (maxPos >= str.length) {
    return str.length;
  }

  let bestBreak = -1;
  let inEncodedWord = false;
  let inAngleBracket = false;
  let inQuotes = false;
  let _encodedWordStart = -1;

  for (let i = 0; i < maxPos && i < str.length; i++) {
    const char = str[i];
    const nextChar = str[i + 1] || "";

    // Track encoded words (=?charset?encoding?text?=)
    if (char === "=" && nextChar === "?") {
      inEncodedWord = true;
      _encodedWordStart = i;
    } else if (inEncodedWord && char === "?" && nextChar === "=") {
      inEncodedWord = false;
      // Skip past the closing ?=
      i++;
      continue;
    }

    // Track angle brackets (email addresses)
    if (char === "<" && !inQuotes && !inEncodedWord) {
      inAngleBracket = true;
    } else if (char === ">" && inAngleBracket) {
      inAngleBracket = false;
    }

    // Track quoted strings
    if (char === '"' && !inEncodedWord) {
      inQuotes = !inQuotes;
    }

    // Only consider break points when not inside special constructs
    if (!inEncodedWord && !inAngleBracket && !inQuotes) {
      // Space is a good break point
      if (char === " ") {
        bestBreak = i;
      }
      // After comma+space is also good for recipient lists
      else if (char === "," && nextChar === " ") {
        bestBreak = i + 1; // Break after the comma
      }
    }
  }

  return bestBreak;
}
