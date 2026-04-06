import type { WorkbookSnapshot } from "@bilig/protocol";

export type WorkbookBenchmarkCorpusId =
  | "dense-mixed-100k"
  | "dense-mixed-250k"
  | "analysis-multisheet-100k"
  | "analysis-multisheet-250k";

export type WorkbookBenchmarkCorpusFamily = "dense-mixed" | "analysis-multisheet";

export interface WorkbookBenchmarkCorpusViewport {
  readonly sheetName: string;
  readonly rowStart: number;
  readonly rowEnd: number;
  readonly colStart: number;
  readonly colEnd: number;
}

export interface WorkbookBenchmarkCorpusDefinition {
  readonly id: WorkbookBenchmarkCorpusId;
  readonly family: WorkbookBenchmarkCorpusFamily;
  readonly label: string;
  readonly description: string;
  readonly materializedCellCount: number;
  readonly sheetCount: number;
  readonly primaryViewport: WorkbookBenchmarkCorpusViewport;
}

export interface WorkbookBenchmarkCorpusCase extends WorkbookBenchmarkCorpusDefinition {
  readonly snapshot: WorkbookSnapshot;
}

type WorkbookBenchmarkSheet = WorkbookSnapshot["sheets"][number];
type WorkbookBenchmarkCell = WorkbookBenchmarkSheet["cells"][number];

interface WorkbookBenchmarkCorpusDescriptor extends WorkbookBenchmarkCorpusDefinition {
  readonly buildSnapshot: (materializedCellCount: number, workbookName: string) => WorkbookSnapshot;
}

function createSheet(id: number, name: string, order: number): WorkbookBenchmarkSheet {
  return {
    id,
    name,
    order,
    cells: [],
  };
}

function formatColumnName(columnIndex: number): string {
  let column = "";
  let value = columnIndex;
  do {
    column = String.fromCharCode(65 + (value % 26)) + column;
    value = Math.floor(value / 26) - 1;
  } while (value >= 0);
  return column;
}

function formatCellAddress(rowIndex: number, columnIndex: number): string {
  return `${formatColumnName(columnIndex)}${String(rowIndex + 1)}`;
}

function countSheetCells(sheet: WorkbookBenchmarkSheet): number {
  return sheet.cells.length;
}

export function countWorkbookSnapshotCells(snapshot: WorkbookSnapshot): number {
  return snapshot.sheets.reduce((sum, sheet) => sum + sheet.cells.length, 0);
}

export function buildDenseMixedWorkbookSnapshot(
  materializedCellCount = 1_000,
  workbookName = "benchmark-dense-mixed",
  sheetName = "Grid",
): WorkbookSnapshot {
  const sheet = createSheet(1, sheetName, 0);
  let remaining = Math.max(1, Math.trunc(materializedCellCount));

  const appendCell = (cell: WorkbookBenchmarkCell): boolean => {
    if (remaining === 0) {
      return false;
    }
    sheet.cells.push(cell);
    remaining -= 1;
    return true;
  };

  for (let rowIndex = 0; ; rowIndex += 1) {
    if (remaining === 0) {
      break;
    }
    const rowNumber = rowIndex + 1;
    appendCell({
      address: formatCellAddress(rowIndex, 0),
      value: rowNumber,
    });
    appendCell({
      address: formatCellAddress(rowIndex, 1),
      value: (rowIndex % 97) + 3,
    });
    appendCell({
      address: formatCellAddress(rowIndex, 2),
      formula: `A${String(rowNumber)}*B${String(rowNumber)}`,
    });
    appendCell({
      address: formatCellAddress(rowIndex, 3),
      value: `segment-${String((rowIndex % 24) + 1)}`,
    });
  }

  return {
    version: 1,
    workbook: { name: workbookName },
    sheets: [sheet],
  };
}

export function buildAnalysisMultisheetWorkbookSnapshot(
  materializedCellCount = 1_000,
  workbookName = "benchmark-analysis-multisheet",
): WorkbookSnapshot {
  const inputs = createSheet(1, "Inputs", 0);
  const ledger = createSheet(2, "Ledger", 1);
  const summary = createSheet(3, "Summary", 2);
  let remaining = Math.max(1, Math.trunc(materializedCellCount));

  const appendCell = (sheet: WorkbookBenchmarkSheet, cell: WorkbookBenchmarkCell): boolean => {
    if (remaining === 0) {
      return false;
    }
    sheet.cells.push(cell);
    remaining -= 1;
    return true;
  };

  const inputRowTarget = Math.max(64, Math.floor((materializedCellCount * 0.12) / 3));
  for (let rowIndex = 0; rowIndex < inputRowTarget; rowIndex += 1) {
    if (remaining === 0) {
      break;
    }
    const rowNumber = rowIndex + 1;
    appendCell(inputs, {
      address: formatCellAddress(rowIndex, 0),
      value: (rowIndex % 31) + 1,
    });
    appendCell(inputs, {
      address: formatCellAddress(rowIndex, 1),
      value: ((rowIndex * 7) % 19) + 2,
    });
    appendCell(inputs, {
      address: formatCellAddress(rowIndex, 2),
      formula: `A${String(rowNumber)}*B${String(rowNumber)}`,
    });
  }

  const inputRowCount = Math.max(1, Math.ceil(countSheetCells(inputs) / 3));
  const ledgerRowTarget = Math.max(256, Math.floor((materializedCellCount * 0.72) / 5));
  for (let rowIndex = 0; rowIndex < ledgerRowTarget; rowIndex += 1) {
    if (remaining === 0) {
      break;
    }
    const rowNumber = rowIndex + 1;
    const inputRowNumber = (rowIndex % inputRowCount) + 1;
    appendCell(ledger, {
      address: formatCellAddress(rowIndex, 0),
      value: (rowIndex % 41) + 1,
    });
    appendCell(ledger, {
      address: formatCellAddress(rowIndex, 1),
      value: ((rowIndex * 3) % 23) + 5,
    });
    appendCell(ledger, {
      address: formatCellAddress(rowIndex, 2),
      value: (rowIndex % 11) + 1,
    });
    appendCell(ledger, {
      address: formatCellAddress(rowIndex, 3),
      formula: `A${String(rowNumber)}*B${String(rowNumber)}`,
    });
    appendCell(ledger, {
      address: formatCellAddress(rowIndex, 4),
      formula: `D${String(rowNumber)}+C${String(rowNumber)}+Inputs!C${String(inputRowNumber)}`,
    });
  }

  const ledgerRowCount = Math.max(1, Math.ceil(countSheetCells(ledger) / 5));
  const summaryWindowSize = Math.max(32, Math.min(256, Math.ceil(ledgerRowCount / 48)));
  for (let rowIndex = 0; ; rowIndex += 1) {
    if (remaining === 0) {
      break;
    }
    const rowNumber = rowIndex + 1;
    const startLedgerRow = ((rowIndex * summaryWindowSize) % ledgerRowCount) + 1;
    const endLedgerRow = Math.min(ledgerRowCount, startLedgerRow + summaryWindowSize - 1);
    const inputRowNumber = (rowIndex % inputRowCount) + 1;
    appendCell(summary, {
      address: formatCellAddress(rowIndex, 0),
      value: `cluster-${String((rowIndex % 64) + 1)}`,
    });
    appendCell(summary, {
      address: formatCellAddress(rowIndex, 1),
      formula: `SUM(Ledger!E${String(startLedgerRow)}:E${String(endLedgerRow)})`,
    });
    appendCell(summary, {
      address: formatCellAddress(rowIndex, 2),
      formula: `B${String(rowNumber)}/Inputs!C${String(inputRowNumber)}`,
    });
    appendCell(summary, {
      address: formatCellAddress(rowIndex, 3),
      formula: `SUM(Inputs!C1:C${String(inputRowCount)})`,
    });
  }

  return {
    version: 1,
    workbook: { name: workbookName },
    sheets: [inputs, ledger, summary],
  };
}

const workbookBenchmarkCorpusIds = [
  "dense-mixed-100k",
  "dense-mixed-250k",
  "analysis-multisheet-100k",
  "analysis-multisheet-250k",
] as const satisfies readonly WorkbookBenchmarkCorpusId[];

const workbookBenchmarkCorpusDescriptors = {
  "dense-mixed-100k": {
    id: "dense-mixed-100k",
    family: "dense-mixed",
    label: "Dense mixed 100k",
    description:
      "Single-sheet dense workbook with numeric, string, and per-row formula cells for baseline giant-workbook load and restore.",
    materializedCellCount: 100_000,
    sheetCount: 1,
    primaryViewport: {
      sheetName: "Grid",
      rowStart: 0,
      rowEnd: 39,
      colStart: 0,
      colEnd: 1,
    },
    buildSnapshot: (materializedCellCount: number, workbookName: string) =>
      buildDenseMixedWorkbookSnapshot(materializedCellCount, workbookName),
  },
  "dense-mixed-250k": {
    id: "dense-mixed-250k",
    family: "dense-mixed",
    label: "Dense mixed 250k",
    description:
      "Single-sheet dense workbook at 250k cells for warm-start and large-restore contracts.",
    materializedCellCount: 250_000,
    sheetCount: 1,
    primaryViewport: {
      sheetName: "Grid",
      rowStart: 0,
      rowEnd: 39,
      colStart: 0,
      colEnd: 1,
    },
    buildSnapshot: (materializedCellCount: number, workbookName: string) =>
      buildDenseMixedWorkbookSnapshot(materializedCellCount, workbookName),
  },
  "analysis-multisheet-100k": {
    id: "analysis-multisheet-100k",
    family: "analysis-multisheet",
    label: "Analysis multisheet 100k",
    description:
      "Three-sheet workbook with cross-sheet formulas and range aggregates for realistic workbook-comprehension and warm-start coverage.",
    materializedCellCount: 100_000,
    sheetCount: 3,
    primaryViewport: {
      sheetName: "Ledger",
      rowStart: 0,
      rowEnd: 39,
      colStart: 0,
      colEnd: 4,
    },
    buildSnapshot: (materializedCellCount: number, workbookName: string) =>
      buildAnalysisMultisheetWorkbookSnapshot(materializedCellCount, workbookName),
  },
  "analysis-multisheet-250k": {
    id: "analysis-multisheet-250k",
    family: "analysis-multisheet",
    label: "Analysis multisheet 250k",
    description:
      "Three-sheet 250k-cell workbook with cross-sheet formulas and range aggregates for giant-workbook restore contracts.",
    materializedCellCount: 250_000,
    sheetCount: 3,
    primaryViewport: {
      sheetName: "Ledger",
      rowStart: 0,
      rowEnd: 39,
      colStart: 0,
      colEnd: 4,
    },
    buildSnapshot: (materializedCellCount: number, workbookName: string) =>
      buildAnalysisMultisheetWorkbookSnapshot(materializedCellCount, workbookName),
  },
} as const satisfies Record<WorkbookBenchmarkCorpusId, WorkbookBenchmarkCorpusDescriptor>;

function toCorpusDefinition(
  descriptor: WorkbookBenchmarkCorpusDescriptor,
): WorkbookBenchmarkCorpusDefinition {
  return {
    id: descriptor.id,
    family: descriptor.family,
    label: descriptor.label,
    description: descriptor.description,
    materializedCellCount: descriptor.materializedCellCount,
    sheetCount: descriptor.sheetCount,
    primaryViewport: descriptor.primaryViewport,
  };
}

export function isWorkbookBenchmarkCorpusId(value: string): value is WorkbookBenchmarkCorpusId {
  return (workbookBenchmarkCorpusIds as readonly string[]).includes(value);
}

export function listWorkbookBenchmarkCorpusDefinitions(): readonly WorkbookBenchmarkCorpusDefinition[] {
  return workbookBenchmarkCorpusIds.map((corpusId) =>
    toCorpusDefinition(workbookBenchmarkCorpusDescriptors[corpusId]),
  );
}

export function getWorkbookBenchmarkCorpusDefinition(
  corpusId: WorkbookBenchmarkCorpusId,
): WorkbookBenchmarkCorpusDefinition {
  return toCorpusDefinition(workbookBenchmarkCorpusDescriptors[corpusId]);
}

export function buildWorkbookBenchmarkCorpus(
  corpusId: WorkbookBenchmarkCorpusId,
): WorkbookBenchmarkCorpusCase {
  const descriptor = workbookBenchmarkCorpusDescriptors[corpusId];
  const snapshot = descriptor.buildSnapshot(
    descriptor.materializedCellCount,
    `benchmark-${descriptor.id}`,
  );
  const actualCellCount = countWorkbookSnapshotCells(snapshot);
  if (actualCellCount !== descriptor.materializedCellCount) {
    throw new Error(
      `Workbook corpus ${descriptor.id} built ${String(actualCellCount)} cells, expected ${String(descriptor.materializedCellCount)}`,
    );
  }
  if (snapshot.sheets.length !== descriptor.sheetCount) {
    throw new Error(
      `Workbook corpus ${descriptor.id} built ${String(snapshot.sheets.length)} sheets, expected ${String(descriptor.sheetCount)}`,
    );
  }
  return {
    ...toCorpusDefinition(descriptor),
    snapshot,
  };
}
