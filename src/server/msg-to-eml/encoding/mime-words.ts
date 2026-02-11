/**
 * Checks if a string contains only ASCII characters.
 */
export function isAscii(str: string): boolean {
  for (let i = 0; i < str.length; i++) {
    if (str.charCodeAt(i) > 127) {
      return false;
    }
  }
  return true;
}

/**
 * RFC 2047 MIME encoded-word encoding.
 * Encodes non-ASCII text for use in email headers (Subject, From display name, etc.)
 * Format: =?charset?encoding?encoded-text?=
 *
 * We use Base64 (B) encoding as it's more compact for non-ASCII text.
 * Each encoded word is limited to 75 characters total.
 *
 * @param text The text to encode
 * @returns RFC 2047 encoded string, or original if ASCII-only
 */
export function encodeRfc2047(text: string): string {
  if (!text || isAscii(text)) {
    return text;
  }

  // Encode as UTF-8 Base64
  const utf8Bytes = new TextEncoder().encode(text);
  const base64 = Buffer.from(utf8Bytes).toString("base64");

  // Build the encoded word
  const prefix = "=?UTF-8?B?";
  const suffix = "?=";
  const overhead = prefix.length + suffix.length; // 12 characters
  const maxEncodedLength = 75 - overhead; // 63 characters for encoded text

  // If the entire encoded string fits in one encoded-word, return it directly
  if (base64.length <= maxEncodedLength) {
    return `${prefix}${base64}${suffix}`;
  }

  // Need to split into multiple encoded-words
  // Each encoded-word must be valid UTF-8, so we need to split on character boundaries
  return encodeRfc2047Chunked(text);
}

/**
 * Encodes text as multiple RFC 2047 encoded-words, properly splitting on character boundaries.
 * Each chunk is encoded separately to ensure valid UTF-8 in each encoded-word.
 */
function encodeRfc2047Chunked(text: string): string {
  const prefix = "=?UTF-8?B?";
  const suffix = "?=";
  const overhead = prefix.length + suffix.length;
  const maxEncodedLength = 75 - overhead;

  // We need to determine how many characters we can encode per chunk
  // Base64 expands 3 bytes to 4 characters
  // A UTF-8 character can be 1-4 bytes, so we need to be conservative
  // For safety, we'll aim for chunks that result in ~60 base64 chars (45 bytes before encoding)

  const result: string[] = [];
  let i = 0;

  while (i < text.length) {
    // Try to find the longest substring that fits in one encoded-word
    let chunk = "";
    let chunkBytes = 0;

    while (i < text.length) {
      const char = text[i];
      const charBytes = new TextEncoder().encode(char).length;

      // Check if adding this character would exceed the limit
      // Base64 encoding: ceil(bytes * 4 / 3)
      const newBytes = chunkBytes + charBytes;
      const newBase64Length = Math.ceil((newBytes * 4) / 3);

      if (newBase64Length > maxEncodedLength && chunk.length > 0) {
        // This character would push us over; encode what we have
        break;
      }

      chunk += char;
      chunkBytes = newBytes;
      i++;
    }

    // Encode this chunk
    const chunkUtf8 = new TextEncoder().encode(chunk);
    const chunkBase64 = Buffer.from(chunkUtf8).toString("base64");
    result.push(`${prefix}${chunkBase64}${suffix}`);
  }

  // Join with space, as per RFC 2047 section 5.3
  // When encoded-words are adjacent, they should be separated by linear whitespace
  return result.join(" ");
}

/**
 * Encodes a display name for use in email headers if it contains non-ASCII characters.
 * Returns the name in quotes if ASCII, or RFC 2047 encoded if non-ASCII.
 */
export function encodeDisplayName(name: string): string {
  if (!name) {
    return "";
  }
  if (isAscii(name)) {
    return `"${name}"`;
  }
  return encodeRfc2047(name);
}
