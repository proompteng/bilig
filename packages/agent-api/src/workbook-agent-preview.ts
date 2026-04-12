import { SpreadsheetEngine } from "@bilig/core";
import type { CellSnapshot, WorkbookSnapshot } from "@bilig/protocol";
import {
  applyWorkbookAgentCommandBundle,
  describeWorkbookAgentCommand,
  isWorkbookAgentCommandBundle,
  type WorkbookAgentCommandBundle,
  type WorkbookAgentPreviewCellDiff,
  type WorkbookAgentPreviewSummary,
} from "./workbook-agent-bundles.js";
import { formatAddress, parseCellAddress } from "@bilig/formula";

const MAX_PREVIEW_DIFFS = 64;

function buildChangeKinds(
  beforeCell: CellSnapshot,
  afterCell: CellSnapshot,
): WorkbookAgentPreviewCellDiff["changeKinds"] {
  return [
    ...(beforeCell.formula !== afterCell.formula ? (["formula"] as const) : []),
    ...(beforeCell.input !== afterCell.input ? (["input"] as const) : []),
    ...(beforeCell.styleId !== afterCell.styleId ? (["style"] as const) : []),
    ...(beforeCell.format !== afterCell.format ? (["numberFormat"] as const) : []),
  ];
}

function cloneCellDiff(
  beforeCell: CellSnapshot,
  afterCell: CellSnapshot,
): WorkbookAgentPreviewCellDiff | null {
  const changeKinds = buildChangeKinds(beforeCell, afterCell);
  if (changeKinds.length === 0) {
    return null;
  }
  return {
    sheetName: afterCell.sheetName,
    address: afterCell.address,
    beforeInput: beforeCell.input ?? null,
    beforeFormula: beforeCell.formula ? `=${beforeCell.formula}` : null,
    afterInput: afterCell.input ?? null,
    afterFormula: afterCell.formula ? `=${afterCell.formula}` : null,
    changeKinds,
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

function buildStructuralChanges(bundle: WorkbookAgentCommandBundle): string[] {
  const structuralChanges: string[] = [];
  bundle.commands.forEach((command) => {
    if (
      command.kind === "createSheet" ||
      command.kind === "renameSheet" ||
      command.kind === "updateRowMetadata" ||
      command.kind === "updateColumnMetadata"
    ) {
      const description = describeWorkbookAgentCommand(command);
      if (!structuralChanges.includes(description)) {
        structuralChanges.push(description);
      }
    }
  });
  return structuralChanges;
}

function buildEffectSummary(input: {
  cellDiffs: readonly WorkbookAgentPreviewCellDiff[];
  structuralChanges: readonly string[];
  truncatedCellDiffs: boolean;
}): WorkbookAgentPreviewSummary["effectSummary"] {
  return {
    displayedCellDiffCount: input.cellDiffs.length,
    truncatedCellDiffs: input.truncatedCellDiffs,
    inputChangeCount: input.cellDiffs.filter((diff) => diff.changeKinds.includes("input")).length,
    formulaChangeCount: input.cellDiffs.filter((diff) => diff.changeKinds.includes("formula"))
      .length,
    styleChangeCount: input.cellDiffs.filter((diff) => diff.changeKinds.includes("style")).length,
    numberFormatChangeCount: input.cellDiffs.filter((diff) =>
      diff.changeKinds.includes("numberFormat"),
    ).length,
    structuralChangeCount: input.structuralChanges.length,
  };
}

export async function buildWorkbookAgentPreview(input: {
  snapshot: WorkbookSnapshot;
  replicaId: string;
  bundle: WorkbookAgentCommandBundle;
}): Promise<WorkbookAgentPreviewSummary> {
  if (!isWorkbookAgentCommandBundle(input.bundle)) {
    throw new Error("Invalid workbook agent command bundle");
  }
  const previewEngine = new SpreadsheetEngine({
    workbookName: input.snapshot.workbook.name,
    replicaId: `${input.replicaId}:agent-preview`,
  });
  await previewEngine.ready();
  previewEngine.importSnapshot(input.snapshot);
  applyWorkbookAgentCommandBundle(previewEngine, input.bundle);
  const beforeEngine = new SpreadsheetEngine({
    workbookName: input.snapshot.workbook.name,
    replicaId: `${input.replicaId}:agent-preview-base`,
  });
  await beforeEngine.ready();
  beforeEngine.importSnapshot(input.snapshot);
  const targetAddresses = collectTargetAddresses(input.bundle);
  const resolvedCellDiffs = targetAddresses
    .flatMap(({ sheetName, address }) => {
      const beforeCell = beforeEngine.getCell(sheetName, address);
      const afterCell = previewEngine.getCell(sheetName, address);
      const diff = cloneCellDiff(beforeCell, afterCell);
      return diff ? [diff] : [];
    })
    .slice(0, MAX_PREVIEW_DIFFS);
  const structuralChanges = buildStructuralChanges(input.bundle);
  const truncatedCellDiffs = input.bundle.affectedRanges.some((range) => {
    if (range.role !== "target") {
      return false;
    }
    const start = parseCellAddress(range.startAddress, range.sheetName);
    const end = parseCellAddress(range.endAddress, range.sheetName);
    const rowCount = Math.abs(end.row - start.row) + 1;
    const colCount = Math.abs(end.col - start.col) + 1;
    return rowCount * colCount > MAX_PREVIEW_DIFFS;
  });
  return {
    ranges: input.bundle.affectedRanges.map((range) => ({ ...range })),
    structuralChanges,
    cellDiffs: resolvedCellDiffs,
    effectSummary: buildEffectSummary({
      cellDiffs: resolvedCellDiffs,
      structuralChanges,
      truncatedCellDiffs,
    }),
  };
}
