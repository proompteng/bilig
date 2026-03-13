const CELL_RE = /^\$?([A-Z]+)\$?([1-9][0-9]*)$/;
const QUALIFIED_RE = /^(?:(?:'((?:[^']|'')+)'|([^!]+))!)?(.+)$/;

export interface CellAddress {
  sheetName?: string;
  row: number;
  col: number;
  text: string;
}

export interface RangeAddress {
  sheetName?: string;
  start: CellAddress;
  end: CellAddress;
}

export function columnToIndex(column: string): number {
  let value = 0;
  for (const char of column) {
    value = value * 26 + (char.charCodeAt(0) - 64);
  }
  return value - 1;
}

export function indexToColumn(index: number): string {
  let current = index + 1;
  let output = "";
  while (current > 0) {
    const rem = (current - 1) % 26;
    output = String.fromCharCode(65 + rem) + output;
    current = Math.floor((current - 1) / 26);
  }
  return output;
}

export function formatAddress(row: number, col: number): string {
  return `${indexToColumn(col)}${row + 1}`;
}

export function parseCellAddress(raw: string, defaultSheetName?: string): CellAddress {
  const trimmed = raw.trim();
  const qualified = QUALIFIED_RE.exec(trimmed);
  if (!qualified) {
    throw new Error(`Invalid cell address: ${raw}`);
  }

  const [, quotedSheet, plainSheet, cellPart] = qualified;
  const normalizedCellPart = cellPart!;
  const match = CELL_RE.exec(normalizedCellPart.toUpperCase());
  if (!match) {
    throw new Error(`Invalid cell address: ${raw}`);
  }

  const sheetName = quotedSheet?.replaceAll("''", "'") ?? plainSheet ?? defaultSheetName;
  const result: CellAddress = {
    col: columnToIndex(match[1]!),
    row: Number.parseInt(match[2]!, 10) - 1,
    text: `${match[1]!}${match[2]!}`
  };
  if (sheetName !== undefined) {
    result.sheetName = sheetName;
  }
  return result;
}

export function parseRangeAddress(raw: string, defaultSheetName?: string): RangeAddress {
  const [left, right] = raw.split(":");
  if (!left || !right) {
    throw new Error(`Invalid range address: ${raw}`);
  }

  const start = parseCellAddress(left, defaultSheetName);
  const end = parseCellAddress(right, start.sheetName ?? defaultSheetName);
  const row1 = Math.min(start.row, end.row);
  const row2 = Math.max(start.row, end.row);
  const col1 = Math.min(start.col, end.col);
  const col2 = Math.max(start.col, end.col);

  const sheetName = start.sheetName ?? end.sheetName;
  const result: RangeAddress = {
    start: { ...start, row: row1, col: col1, text: formatAddress(row1, col1) },
    end: { ...end, row: row2, col: col2, text: formatAddress(row2, col2) }
  };
  if (sheetName !== undefined) {
    result.sheetName = sheetName;
  }
  return result;
}

export function toQualifiedAddress(sheetName: string, addr: string): string {
  return `${sheetName}!${parseCellAddress(addr, sheetName).text}`;
}
