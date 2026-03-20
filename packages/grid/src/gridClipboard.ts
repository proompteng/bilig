export function serializeClipboardMatrix(values: readonly (readonly string[])[]): string {
  return values.map((row) => row.join("\u001f")).join("\u001e");
}

export function serializeClipboardPlainText(values: readonly (readonly string[])[]): string {
  return values.map((row) => row.join("\t")).join("\n");
}

export function parseClipboardPlainText(rawText: string): readonly (readonly string[])[] {
  if (rawText.length === 0) {
    return [];
  }
  return rawText
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .split("\n")
    .map((row) => row.split("\t"));
}
