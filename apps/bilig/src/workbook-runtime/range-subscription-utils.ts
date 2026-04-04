import { SpreadsheetEngine } from "@bilig/core";
import type { CellRangeRef, EngineEvent } from "@bilig/protocol";

export function iterateRange(range: CellRangeRef): string[] {
  const [startColPart, startRowPart] = splitAddress(range.startAddress);
  const [endColPart, endRowPart] = splitAddress(range.endAddress);
  const startCol = decodeColumn(startColPart);
  const endCol = decodeColumn(endColPart);
  const startRow = Number.parseInt(startRowPart, 10);
  const endRow = Number.parseInt(endRowPart, 10);
  const addresses: string[] = [];
  for (let row = startRow; row <= endRow; row += 1) {
    for (let col = startCol; col <= endCol; col += 1) {
      addresses.push(`${encodeColumn(col)}${row}`);
    }
  }
  return addresses;
}

export function splitAddress(address: string): [string, string] {
  const match = /^([A-Z]+)(\d+)$/i.exec(address.trim());
  if (!match) {
    throw new Error(`Invalid cell address: ${address}`);
  }
  return [match[1]!.toUpperCase(), match[2]!];
}

export function decodeColumn(column: string): number {
  let value = 0;
  for (let index = 0; index < column.length; index += 1) {
    value = value * 26 + (column.charCodeAt(index) - 64);
  }
  return value;
}

export function encodeColumn(value: number): string {
  let next = value;
  let output = "";
  while (next > 0) {
    const remainder = (next - 1) % 26;
    output = String.fromCharCode(65 + remainder) + output;
    next = Math.floor((next - 1) / 26);
  }
  return output;
}

export function cellCountForRange(range: CellRangeRef): number {
  const [startColPart, startRowPart] = splitAddress(range.startAddress);
  const [endColPart, endRowPart] = splitAddress(range.endAddress);
  const width = decodeColumn(endColPart) - decodeColumn(startColPart) + 1;
  const height = Number.parseInt(endRowPart, 10) - Number.parseInt(startRowPart, 10) + 1;
  return width * height;
}

function collectChangedAddressesInRange(
  engine: SpreadsheetEngine,
  range: CellRangeRef,
  changedCellIndices: readonly number[] | Uint32Array,
): string[] {
  const [startColPart, startRowPart] = splitAddress(range.startAddress);
  const [endColPart, endRowPart] = splitAddress(range.endAddress);
  const startCol = decodeColumn(startColPart);
  const endCol = decodeColumn(endColPart);
  const startRow = Number.parseInt(startRowPart, 10);
  const endRow = Number.parseInt(endRowPart, 10);
  const changedAddresses: string[] = [];

  for (let index = 0; index < changedCellIndices.length; index += 1) {
    const qualifiedAddress = engine.workbook.getQualifiedAddress(changedCellIndices[index]!);
    if (!qualifiedAddress.startsWith(`${range.sheetName}!`)) {
      continue;
    }
    const address = qualifiedAddress.slice(range.sheetName.length + 1);
    const parsed = splitAddress(address);
    const col = decodeColumn(parsed[0]);
    const row = Number.parseInt(parsed[1], 10);
    if (col < startCol || col > endCol || row < startRow || row > endRow) {
      continue;
    }
    changedAddresses.push(address);
  }

  return changedAddresses;
}

function collectAddressesForIntersection(
  range: CellRangeRef,
  startAddress: string,
  endAddress: string,
): string[] {
  const rangeStart = splitAddress(range.startAddress);
  const rangeEnd = splitAddress(range.endAddress);
  const eventStart = splitAddress(startAddress);
  const eventEnd = splitAddress(endAddress);

  const startCol = Math.max(decodeColumn(rangeStart[0]), decodeColumn(eventStart[0]));
  const endCol = Math.min(decodeColumn(rangeEnd[0]), decodeColumn(eventEnd[0]));
  const startRow = Math.max(Number.parseInt(rangeStart[1], 10), Number.parseInt(eventStart[1], 10));
  const endRow = Math.min(Number.parseInt(rangeEnd[1], 10), Number.parseInt(eventEnd[1], 10));

  if (startCol > endCol || startRow > endRow) {
    return [];
  }

  const changedAddresses: string[] = [];
  for (let row = startRow; row <= endRow; row += 1) {
    for (let col = startCol; col <= endCol; col += 1) {
      changedAddresses.push(`${encodeColumn(col)}${row}`);
    }
  }
  return changedAddresses;
}

export function collectChangedAddressesForEvent(
  engine: SpreadsheetEngine,
  range: CellRangeRef,
  event: EngineEvent,
): string[] {
  if (event.invalidation === "full") {
    return iterateRange(range);
  }

  const changedAddresses = new Set(
    collectChangedAddressesInRange(engine, range, event.changedCellIndices),
  );
  for (let index = 0; index < event.invalidatedRanges.length; index += 1) {
    const invalidatedRange = event.invalidatedRanges[index]!;
    if (invalidatedRange.sheetName !== range.sheetName) {
      continue;
    }
    collectAddressesForIntersection(
      range,
      invalidatedRange.startAddress,
      invalidatedRange.endAddress,
    ).forEach((address) => {
      changedAddresses.add(address);
    });
  }
  return [...changedAddresses];
}
