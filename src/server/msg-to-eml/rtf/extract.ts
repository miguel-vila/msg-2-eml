import { decompressRTF } from "@kenjiuno/decompressrtf";
import * as iconvLite from "iconv-lite";
import { deEncapsulateSync } from "rtf-stream-parser";

export interface RtfBodyResult {
  text: string;
  html?: string;
}

/**
 * Decompresses RTF from PidTagRtfCompressed and extracts text/HTML content.
 * RTF may contain HTML wrapped in \fromhtml1 tags (RTF-encapsulated HTML).
 */
export function extractBodyFromRtf(compressedRtf: number[]): RtfBodyResult | null {
  if (!compressedRtf || compressedRtf.length === 0) {
    return null;
  }

  try {
    // Step 1: Decompress the RTF (returns number[])
    const decompressed = decompressRTF(compressedRtf);
    const rtfString = Buffer.from(decompressed).toString("latin1");

    // Step 2: Try to de-encapsulate to extract HTML or text
    // The rtf-stream-parser library throws if RTF is not encapsulated
    try {
      const result = deEncapsulateSync(rtfString, {
        decode: iconvLite.decode,
        mode: "either",
      });

      if (result.mode === "html") {
        // RTF contained encapsulated HTML
        const html = typeof result.text === "string" ? result.text : result.text.toString("utf-8");
        // Also provide a plain text fallback by stripping HTML tags
        const text = stripHtmlTags(html);
        return { text, html };
      } else {
        // RTF contained encapsulated plain text
        const text = typeof result.text === "string" ? result.text : result.text.toString("utf-8");
        return { text };
      }
    } catch {
      // RTF is not encapsulated, extract plain text directly
      const text = extractPlainTextFromRtf(rtfString);
      return text ? { text } : null;
    }
  } catch {
    // If decompression fails, return null
    return null;
  }
}

/**
 * Extracts plain text from raw RTF content (non-encapsulated).
 * This is a simple RTF parser for basic text extraction.
 */
function extractPlainTextFromRtf(rtf: string): string {
  let text = "";
  let i = 0;
  let depth = 0;
  let skipGroup = 0;

  // Groups to skip (they don't contain visible text)
  const skipPatterns = /^\\(fonttbl|colortbl|stylesheet|info|pict|object|fldinst|fldrslt)/;

  while (i < rtf.length) {
    const char = rtf[i];

    if (char === "{") {
      depth++;
      // Check if this group should be skipped
      const remaining = rtf.slice(i + 1, i + 20);
      if (skipPatterns.test(remaining)) {
        skipGroup = depth;
      }
      i++;
    } else if (char === "}") {
      if (skipGroup === depth) {
        skipGroup = 0;
      }
      depth--;
      i++;
    } else if (skipGroup > 0) {
      // Skip content in ignored groups
      i++;
    } else if (char === "\\") {
      // Handle control words and escape sequences
      i++;
      if (i >= rtf.length) break;

      const nextChar = rtf[i];

      // Escape sequences
      if (nextChar === "'" && i + 2 < rtf.length) {
        // Hex escape like \'e9 for é
        const hex = rtf.slice(i + 1, i + 3);
        const code = parseInt(hex, 16);
        if (!Number.isNaN(code)) {
          text += String.fromCharCode(code);
        }
        i += 3;
      } else if (nextChar === "\\" || nextChar === "{" || nextChar === "}") {
        text += nextChar;
        i++;
      } else if (nextChar === "\n" || nextChar === "\r") {
        // Line break in RTF source, continue
        i++;
      } else {
        // Control word - read until space or non-letter
        let controlWord = "";
        while (i < rtf.length && /[a-zA-Z]/.test(rtf[i])) {
          controlWord += rtf[i];
          i++;
        }
        // Skip optional numeric parameter
        while (i < rtf.length && /[-0-9]/.test(rtf[i])) {
          i++;
        }
        // Skip single space delimiter if present
        if (i < rtf.length && rtf[i] === " ") {
          i++;
        }

        // Handle special control words
        if (controlWord === "par" || controlWord === "line") {
          text += "\n";
        } else if (controlWord === "tab") {
          text += "\t";
        }
        // Other control words are ignored
      }
    } else {
      // Regular character
      if (char !== "\r" && char !== "\n") {
        text += char;
      }
      i++;
    }
  }

  return text.trim();
}

/**
 * Simple HTML tag stripper for plain text fallback.
 */
function stripHtmlTags(html: string): string {
  return html
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, "") // Remove style blocks
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, "") // Remove script blocks
    .replace(/<[^>]+>/g, "") // Remove HTML tags
    .replace(/&nbsp;/g, " ") // Replace &nbsp; with space
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\s+/g, " ") // Normalize whitespace
    .trim();
}
