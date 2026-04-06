import { SpreadsheetEngine } from "@bilig/core";
import {
  applyWorkbookAgentCommandBundle,
  isWorkbookAgentCommandBundle,
  type WorkbookAgentCommandBundle,
  type WorkbookAgentPreviewCellDiff,
  type WorkbookAgentPreviewSummary,
} from "@bilig/agent-api";
import { formatAddress, parseCellAddress } from "@bilig/formula";
import type { CellSnapshot, WorkbookSnapshot } from "@bilig/protocol";

const MAX_PREVIEW_DIFFS = 64;

function cloneCellDiff(
  beforeCell: CellSnapshot,
  afterCell: CellSnapshot,
): WorkbookAgentPreviewCellDiff | null {
  const beforeFormula = beforeCell.formula ? `=${beforeCell.formula}` : null;
  const afterFormula = afterCell.formula ? `=${afterCell.formula}` : null;
  const beforeInput = beforeCell.input ?? null;
  const afterInput = afterCell.input ?? null;
  if (
    beforeFormula === afterFormula &&
    beforeInput === afterInput &&
    beforeCell.styleId === afterCell.styleId &&
    beforeCell.format === afterCell.format
  ) {
    return null;
  }
  return {
    sheetName: afterCell.sheetName,
    address: afterCell.address,
    beforeInput,
    beforeFormula,
    afterInput,
    afterFormula,
  };
}

function collectTargetAddresses(bundle: WorkbookAgentCommandBundle): readonly {
  sheetName: string;
  address: string;
}[] {
  const addresses: Array<{ sheetName: string; address: string }> = [];
  bundle.affectedRanges
    .filter((range) => range.role === "target")
    .forEach((range) => {
      const start = parseCellAddress(range.startAddress, range.sheetName);
      const end = parseCellAddress(range.endAddress, range.sheetName);
      const rowStart = Math.min(start.row, end.row);
      const rowEnd = Math.max(start.row, end.row);
      const colStart = Math.min(start.col, end.col);
      const colEnd = Math.max(start.col, end.col);
      for (let row = rowStart; row <= rowEnd && addresses.length < MAX_PREVIEW_DIFFS; row += 1) {
        for (let col = colStart; col <= colEnd && addresses.length < MAX_PREVIEW_DIFFS; col += 1) {
          const address = formatAddress(row, col);
          if (
            !addresses.some(
              (entry) => entry.sheetName === range.sheetName && entry.address === address,
            )
          ) {
            addresses.push({ sheetName: range.sheetName, address });
          }
        }
      }
    });
  return addresses;
}

function buildStructuralChanges(
  beforeSnapshot: WorkbookSnapshot,
  afterSnapshot: WorkbookSnapshot,
): string[] {
  const beforeSheets = beforeSnapshot.sheets.map((sheet) => sheet.name);
  const afterSheets = afterSnapshot.sheets.map((sheet) => sheet.name);
  const structuralChanges: string[] = [];
  afterSheets.forEach((name) => {
    if (!beforeSheets.includes(name)) {
      structuralChanges.push(`Create sheet ${name}`);
    }
  });
  beforeSheets.forEach((name) => {
    if (!afterSheets.includes(name)) {
      structuralChanges.push(`Remove or rename sheet ${name}`);
    }
  });
  return structuralChanges;
}

export async function buildWorkbookAgentPreview(input: {
  snapshot: WorkbookSnapshot;
  replicaId: string;
  bundle: WorkbookAgentCommandBundle;
}): Promise<WorkbookAgentPreviewSummary> {
  if (!isWorkbookAgentCommandBundle(input.bundle)) {
    throw new Error("Invalid workbook agent preview bundle");
  }
  const previewEngine = new SpreadsheetEngine({
    workbookName: input.snapshot.workbook.name,
    replicaId: `${input.replicaId}:agent-preview`,
  });
  await previewEngine.ready();
  previewEngine.importSnapshot(input.snapshot);
  applyWorkbookAgentCommandBundle(previewEngine, input.bundle);
  const afterSnapshot = previewEngine.exportSnapshot();
  const beforeEngine = new SpreadsheetEngine({
    workbookName: input.snapshot.workbook.name,
    replicaId: `${input.replicaId}:agent-preview-base`,
  });
  await beforeEngine.ready();
  beforeEngine.importSnapshot(input.snapshot);
  const resolvedCellDiffs = collectTargetAddresses(input.bundle)
    .flatMap(({ sheetName, address }) => {
      const beforeCell = beforeEngine.getCell(sheetName, address);
      const afterCell = previewEngine.getCell(sheetName, address);
      const diff = cloneCellDiff(beforeCell, afterCell);
      return diff ? [diff] : [];
    })
    .slice(0, MAX_PREVIEW_DIFFS);
  return {
    ranges: input.bundle.affectedRanges.map((range) => ({ ...range })),
    structuralChanges: buildStructuralChanges(input.snapshot, afterSnapshot),
    cellDiffs: resolvedCellDiffs,
  };
}
