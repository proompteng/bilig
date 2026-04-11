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

export function buildBatchMultiColumnRows(rowCount: number): WorkPaperSheet {
  return Array.from({ length: rowCount }, (_, index) => [
    index + 1,
    index + 2,
    `=A${index + 1}+B${index + 1}`,
    `=A${index + 1}*B${index + 1}`,
  ]);
}

export function buildMixedContentSheet(rowCount: number): WorkPaperSheet {
  return Array.from({ length: rowCount }, (_, index) => {
    const rowNumber = index + 1;
    return [
      rowNumber,
      rowNumber * 10,
      `label-${rowNumber}`,
      rowNumber % 2 === 0,
      `=A${rowNumber}+B${rowNumber}`,
      `=E${rowNumber}*2`,
    ];
  });
}

export function buildMultiSheetLiteralSheets(
  sheetCount: number,
  rows: number,
  cols: number,
): Record<string, WorkPaperSheet> {
  return Object.fromEntries(
    Array.from({ length: sheetCount }, (_, sheetIndex) => [
      `Sheet${sheetIndex + 1}`,
      Array.from({ length: rows }, (_rowValue, rowIndex) =>
        Array.from(
          { length: cols },
          (_colValue, colIndex) => sheetIndex * rows * cols + rowIndex * cols + colIndex + 1,
        ),
      ),
    ]),
  );
}

export function buildFormulaFanoutRow(fanoutCount: number): readonly (number | string)[] {
  const row: Array<number | string> = [1];
  for (let col = 1; col <= fanoutCount; col += 1) {
    row.push(`=$A$1+${col}`);
  }
  return row;
}

export function buildFormulaEditChainRow(downstreamCount: number): readonly (number | string)[] {
  const row: Array<number | string> = [1, 2, "=A1+B1"];
  for (let col = 3; col <= downstreamCount + 2; col += 1) {
    row.push(`=${columnLabel(col - 1)}1+1`);
  }
  return row;
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

export function buildApproxLookupSheet(rowCount: number): WorkPaperSheet {
  const rows: Array<Array<number | string>> = [
    ["Key", "Value", "", Math.floor(rowCount / 2) + 0.5, ""],
  ];
  for (let index = 1; index <= rowCount; index += 1) {
    rows.push([index, index * 10]);
  }
  const headerRow = rows[0];
  if (!headerRow) {
    throw new Error("approx lookup benchmark header row is missing");
  }
  headerRow[4] = `=MATCH(D1,A2:A${rowCount + 1},1)`;
  return rows;
}

export function buildTextLookupSheet(rowCount: number): WorkPaperSheet {
  const midpoint = Math.floor(rowCount / 2);
  const rows: Array<Array<number | string>> = [["Key", "Value", "", textLookupKey(midpoint), ""]];
  for (let index = 1; index <= rowCount; index += 1) {
    rows.push([textLookupKey(index), index * 10]);
  }
  const headerRow = rows[0];
  if (!headerRow) {
    throw new Error("text lookup benchmark header row is missing");
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

export function textLookupKey(index: number): string {
  return `KEY-${String(index).padStart(5, "0")}`;
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
