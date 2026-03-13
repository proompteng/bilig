import { ErrorCode, ValueTag } from "@bilig/protocol";
import type { CellSnapshot, LiteralInput } from "@bilig/protocol";

const NUMERIC_RE = /^-?\d+(\.\d+)?$/;

interface CsvCellInput {
  formula?: string;
  value?: LiteralInput;
}

function escapeCsvValue(value: string): string {
  if (!/[",\n\r]/.test(value)) {
    return value;
  }
  return `"${value.replaceAll("\"", "\"\"")}"`;
}

export function cellToCsvValue(cell: CellSnapshot): string {
  if (cell.formula !== undefined) {
    return `=${cell.formula}`;
  }

  switch (cell.value.tag) {
    case ValueTag.Empty:
      return "";
    case ValueTag.Number:
      return String(cell.value.value);
    case ValueTag.Boolean:
      return cell.value.value ? "TRUE" : "FALSE";
    case ValueTag.String:
      return cell.value.value;
    case ValueTag.Error:
      return `#${ErrorCode[cell.value.code] ?? cell.value.code}`;
  }
}

export function serializeCsv(rows: string[][]): string {
  return rows.map((row) => row.map((value) => escapeCsvValue(value)).join(",")).join("\n");
}

export function parseCsv(csv: string): string[][] {
  const rows: string[][] = [];
  let currentRow: string[] = [];
  let currentValue = "";
  let index = 0;
  let inQuotes = false;

  while (index < csv.length) {
    const char = csv[index]!;
    const nextChar = csv[index + 1];

    if (inQuotes) {
      if (char === "\"" && nextChar === "\"") {
        currentValue += "\"";
        index += 2;
        continue;
      }
      if (char === "\"") {
        inQuotes = false;
        index += 1;
        continue;
      }
      currentValue += char;
      index += 1;
      continue;
    }

    if (char === "\"") {
      inQuotes = true;
      index += 1;
      continue;
    }

    if (char === ",") {
      currentRow.push(currentValue);
      currentValue = "";
      index += 1;
      continue;
    }

    if (char === "\r" || char === "\n") {
      currentRow.push(currentValue);
      currentValue = "";
      rows.push(currentRow);
      currentRow = [];
      if (char === "\r" && nextChar === "\n") {
        index += 2;
      } else {
        index += 1;
      }
      continue;
    }

    currentValue += char;
    index += 1;
  }

  currentRow.push(currentValue);
  if (currentRow.length > 1 || currentRow[0] !== "" || rows.length > 0) {
    rows.push(currentRow);
  }

  return rows;
}

export function parseCsvCellInput(raw: string): CsvCellInput | undefined {
  const normalized = raw.trim();
  if (normalized === "") {
    return undefined;
  }
  if (normalized.startsWith("=")) {
    return { formula: normalized.slice(1) };
  }
  if (normalized === "TRUE" || normalized === "FALSE") {
    return { value: normalized === "TRUE" };
  }
  if (NUMERIC_RE.test(normalized)) {
    return { value: Number(normalized) };
  }
  return { value: raw };
}
