import type { WorkPaperCellAddress, WorkPaperCellRange, WorkPaperSheet } from "@bilig/headless";

export function buildDenseLiteralSheet(rows: number, cols: number): WorkPaperSheet {
  return Array.from({ length: rows }, (_rowValue, rowIndex) =>
    Array.from({ length: cols }, (_colValue, colIndex) => rowIndex * cols + colIndex + 1),
  );
}

export function buildFormulaChainRow(downstreamCount: number): readonly (number | string)[] {
  const row: Array<number | string> = [1];
  for (let col = 1; col <= downstreamCount; col += 1) {
    row.push(`=${columnLabel(col - 1)}1+1`);
  }
  return row;
}

export function buildValueFormulaRows(rowCount: number): WorkPaperSheet {
  return Array.from({ length: rowCount }, (_, index) => [index + 1, `=A${index + 1}*2`]);
}

export function buildLookupSheet(rowCount: number): WorkPaperSheet {
  const rows: Array<Array<number | string>> = [
    ["Key", "Value", "", Math.floor(rowCount / 2), "=MATCH(D1,A2:A20001,0)"],
  ];
  for (let index = 1; index <= rowCount; index += 1) {
    rows.push([index, index * 10]);
  }
  const headerRow = rows[0];
  if (!headerRow) {
    throw new Error("lookup benchmark header row is missing");
  }
  headerRow[4] = `=MATCH(D1,A2:A${rowCount + 1},0)`;
  return rows;
}

export function buildDynamicArraySheet(rowCount: number): WorkPaperSheet {
  const rows: Array<Array<number | string>> = [
    [0, Math.floor(rowCount / 2), `=FILTER(A2:A${rowCount + 1},A2:A${rowCount + 1}>B1)`],
  ];
  for (let index = 1; index <= rowCount; index += 1) {
    rows.push([index]);
  }
  return rows;
}

export function address(sheet: number, row: number, col: number): WorkPaperCellAddress {
  return { sheet, row, col };
}

export function range(
  sheet: number,
  startRow: number,
  startCol: number,
  endRow: number,
  endCol: number,
): WorkPaperCellRange {
  return {
    start: address(sheet, startRow, startCol),
    end: address(sheet, endRow, endCol),
  };
}

export function columnLabel(index: number): string {
  let current = index;
  let label = "";
  while (current >= 0) {
    label = String.fromCharCode(65 + (current % 26)) + label;
    current = Math.floor(current / 26) - 1;
  }
  return label;
}
