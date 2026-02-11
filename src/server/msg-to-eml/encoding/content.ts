export function encodeBase64(data: Uint8Array | number[]): string {
  return Buffer.from(data).toString("base64");
}

export function encodeQuotedPrintable(str: string): string {
  return str.replace(/[^\x20-\x7E\r\n]|=/g, (char) => {
    return `=${char.charCodeAt(0).toString(16).toUpperCase().padStart(2, "0")}`;
  });
}
