import { isAscii } from "./mime-words.js";

/**
 * Encodes a filename according to RFC 2231 for non-ASCII characters.
 * Returns the encoded format: UTF-8''<percent-encoded-value>
 *
 * RFC 2231 specifies that:
 * - attr*=charset'language'encoded-value
 * - We use UTF-8 charset and leave language empty
 * - Characters are percent-encoded (similar to URL encoding)
 */
export function encodeRfc2231(filename: string): string {
  const encoder = new TextEncoder();
  const bytes = encoder.encode(filename);
  let encoded = "";

  for (const byte of bytes) {
    // RFC 2231 allows: ALPHA / DIGIT / "!" / "#" / "$" / "&" / "+" / "-" / "." /
    // "^" / "_" / "`" / "|" / "~" and attribute-char which excludes "*", "'", "%"
    // For simplicity and safety, we only allow alphanumeric, -, ., and _
    if (
      (byte >= 0x30 && byte <= 0x39) || // 0-9
      (byte >= 0x41 && byte <= 0x5a) || // A-Z
      (byte >= 0x61 && byte <= 0x7a) || // a-z
      byte === 0x2d || // -
      byte === 0x2e || // .
      byte === 0x5f // _
    ) {
      encoded += String.fromCharCode(byte);
    } else {
      // Percent-encode the byte
      encoded += `%${byte.toString(16).toUpperCase().padStart(2, "0")}`;
    }
  }

  return `UTF-8''${encoded}`;
}

/**
 * Generates Content-Type and Content-Disposition parameters for a filename.
 * Uses RFC 2231 encoding (filename*=) for non-ASCII filenames,
 * otherwise uses the simple quoted form (filename=).
 */
export function formatFilenameParams(filename: string): { name: string; disposition: string } {
  if (isAscii(filename)) {
    return {
      name: `name="${filename}"`,
      disposition: `filename="${filename}"`,
    };
  } else {
    const encoded = encodeRfc2231(filename);
    return {
      name: `name*=${encoded}`,
      disposition: `filename*=${encoded}`,
    };
  }
}
