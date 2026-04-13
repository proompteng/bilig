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

export function buildOverlappingAggregateSheet(rowCount: number): WorkPaperSheet {
  return Array.from({ length: rowCount }, (_, index) => {
    const rowNumber = index + 1;
    return [rowNumber, `=SUM(A1:A${rowNumber})`];
  });
}

export function buildConditionalAggregationSheet(
  rowCount: number,
  formulaCopies: number,
): WorkPaperSheet {
  const rows: Array<Array<number | string>> = [
    [
      "Group",
      "Value",
      "",
      "A",
      ...Array.from(
        { length: formulaCopies },
        () => `=SUMIF(A2:A${rowCount + 1},D1,B2:B${rowCount + 1})`,
      ),
      ...Array.from({ length: formulaCopies }, () => `=COUNTIF(A2:A${rowCount + 1},D1)`),
    ],
  ];
  for (let index = 1; index <= rowCount; index += 1) {
    rows.push([index % 2 === 0 ? "A" : "B", index]);
  }
  return rows;
}

export function buildParserCacheTemplateSheet(rowCount: number): WorkPaperSheet {
  return Array.from({ length: rowCount }, (_, index) => {
    const rowNumber = index + 1;
    return [
      rowNumber,
      rowNumber * 2,
      `=A${rowNumber}+B${rowNumber}`,
      `=C${rowNumber}*2`,
      `=SUM(A1:A${rowNumber})`,
      `=D${rowNumber}+E${rowNumber}`,
    ];
  });
}

export function buildMixedFrontierSheet(rowCount: number): WorkPaperSheet {
  return Array.from({ length: rowCount }, (_, index) => {
    const rowNumber = index + 1;
    return [rowNumber, `=$A$1+${rowNumber}`, `=SUM(A1:A${rowNumber})`];
  });
}

export function buildStructuralColumnSheet(rowCount: number): WorkPaperSheet {
  return Array.from({ length: rowCount }, (_, index) => {
    const rowNumber = index + 1;
    return [rowNumber, rowNumber * 2, `=A${rowNumber}+B${rowNumber}`, `=C${rowNumber}*2`];
  });
}

export function buildRenameDependencySheets(): Record<string, WorkPaperSheet> {
  return {
    Data: [[1], [2], [3]],
    Summary: [["=Data!A1+1", "=SUM(Data!A1:A3)"]],
  };
}

export function buildNamedExpressionBenchSheet(): WorkPaperSheet {
  return [[1, "=Rate+1", "=Rate*2"], [2]];
}

export function buildLookupSearchModeReverseSheet(rowCount: number): WorkPaperSheet {
  const target = Math.floor(rowCount / 2);
  const rows: Array<Array<number | string>> = [["Key", "Value", "", target, ""]];
  for (let index = 1; index <= rowCount; index += 1) {
    const key = index === target + 1 ? target : index;
    rows.push([key, key * 10]);
  }
  rows[0]![4] = `=XMATCH(D1,A2:A${rowCount + 1},0,-1)`;
  return rows;
}

export function buildApproxLookupDescendingSheet(rowCount: number): WorkPaperSheet {
  const rows: Array<Array<number | string>> = [
    ["Key", "Value", "", Math.floor(rowCount / 2) + 0.5, ""],
  ];
  for (let index = rowCount; index >= 1; index -= 1) {
    rows.push([index, index * 10]);
  }
  rows[0]![4] = `=MATCH(D1,A2:A${rowCount + 1},-1)`;
  return rows;
}

export function buildApproxLookupDuplicateSheet(rowCount: number): WorkPaperSheet {
  const rows: Array<Array<number | string>> = [["Key", "Value", "", Math.floor(rowCount / 4), ""]];
  for (let index = 1; index <= rowCount; index += 1) {
    const key = Math.ceil(index / 2);
    rows.push([key, key * 10]);
  }
  rows[0]![4] = `=MATCH(D1,A2:A${rowCount + 1},1)`;
  return rows;
}

export function buildConditionalAggregationSharedCriteriaSheet(
  rowCount: number,
  criteriaCount: number,
): WorkPaperSheet {
  const criteriaCells = Array.from({ length: criteriaCount }, (_, index) =>
    String.fromCharCode(65 + (index % 4)),
  );
  const rows: Array<Array<number | string>> = [
    [
      "Group",
      "Value",
      "",
      ...criteriaCells,
      ...criteriaCells.map(
        (_, index) => `=SUMIF(A2:A${rowCount + 1},${columnLabel(3 + index)}1,B2:B${rowCount + 1})`,
      ),
    ],
  ];
  for (let index = 1; index <= rowCount; index += 1) {
    rows.push([String.fromCharCode(65 + (index % 4)), index]);
  }
  return rows;
}

export function buildConditionalAggregationMixedSheet(
  rowCount: number,
  formulaCopies: number,
): WorkPaperSheet {
  const rows: Array<Array<number | string>> = [
    [
      "Group",
      "Amount",
      "Flag",
      "A",
      10,
      ...Array.from(
        { length: formulaCopies },
        () =>
          `=COUNTIFS(A2:A${rowCount + 1},D1,B2:B${rowCount + 1},">="&E1,C2:C${rowCount + 1},"x")`,
      ),
      ...Array.from(
        { length: formulaCopies },
        () =>
          `=SUMIFS(B2:B${rowCount + 1},A2:A${rowCount + 1},D1,B2:B${rowCount + 1},">="&E1,C2:C${rowCount + 1},"x")`,
      ),
    ],
  ];
  for (let index = 1; index <= rowCount; index += 1) {
    rows.push([index % 2 === 0 ? "A" : "B", index, index % 3 === 0 ? "x" : "y"]);
  }
  return rows;
}

export function buildParserCacheUniqueFormulaSheet(rowCount: number): WorkPaperSheet {
  return Array.from({ length: rowCount }, (_, index) => {
    const rowNumber = index + 1;
    return [
      rowNumber,
      rowNumber * 2,
      `=A${rowNumber}+B${rowNumber}+${rowNumber}`,
      `=C${rowNumber}*2+${rowNumber}`,
      `=SUM(A1:A${rowNumber})+${rowNumber}`,
      `=D${rowNumber}+E${rowNumber}+${rowNumber}`,
    ];
  });
}

export function buildParserCacheMixedTemplateSheet(rowCount: number): WorkPaperSheet {
  return Array.from({ length: rowCount }, (_, index) => {
    const rowNumber = index + 1;
    if (rowNumber % 3 === 1) {
      return [
        rowNumber,
        rowNumber * 2,
        `=A${rowNumber}+B${rowNumber}`,
        `=C${rowNumber}*2`,
        `=SUM(A1:A${rowNumber})`,
        `=D${rowNumber}+E${rowNumber}`,
      ];
    }
    if (rowNumber % 3 === 2) {
      return [
        rowNumber,
        rowNumber * 3,
        `=A${rowNumber}*B${rowNumber}`,
        `=C${rowNumber}+A${rowNumber}`,
        `=AVERAGE(A1:A${rowNumber})`,
        `=D${rowNumber}-E${rowNumber}`,
      ];
    }
    return [
      rowNumber,
      rowNumber * 4,
      `=B${rowNumber}-A${rowNumber}`,
      `=ABS(C${rowNumber})`,
      `=MAX(A1:A${rowNumber})`,
      `=D${rowNumber}+E${rowNumber}`,
    ];
  });
}

export function buildSlidingAggregateSheet(rowCount: number, window: number): WorkPaperSheet {
  return Array.from({ length: rowCount }, (_, index) => {
    const rowNumber = index + 1;
    const endRow = Math.min(rowCount, rowNumber + window - 1);
    return [rowNumber, `=SUM(A${rowNumber}:A${endRow})`];
  });
}

export function build2dAggregateSheet(rowCount: number): WorkPaperSheet {
  return Array.from({ length: rowCount }, (_, index) => {
    const rowNumber = index + 1;
    return [rowNumber, rowNumber * 2, `=SUM(A1:B${rowNumber})`];
  });
}

export function buildDynamicArraySortSheet(rowCount: number): WorkPaperSheet {
  const rows: Array<Array<number | string>> = [[`=SORT(A2:A${rowCount + 1})`]];
  for (let index = 0; index < rowCount; index += 1) {
    rows.push([rowCount - index]);
  }
  return rows;
}

export function buildDynamicArrayUniqueSheet(rowCount: number): WorkPaperSheet {
  const rows: Array<Array<number | string>> = [[`=UNIQUE(A2:A${rowCount + 1})`]];
  for (let index = 0; index < rowCount; index += 1) {
    rows.push([Math.floor(index / 2) + 1]);
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
