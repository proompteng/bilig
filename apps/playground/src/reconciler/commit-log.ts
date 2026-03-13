import type { CommitOp } from "@bilig/core";
import type { RenderModel } from "./descriptors.js";

function sameCell(
  left: { value?: string | number | boolean | null; formula?: string },
  right: { value?: string | number | boolean | null; formula?: string }
): boolean {
  return left.value === right.value && left.formula === right.formula;
}

export function diffModels(previous: RenderModel, next: RenderModel): CommitOp[] {
  const ops: CommitOp[] = [];
  if (previous.workbookName !== next.workbookName) {
    ops.push({ kind: "upsertWorkbook", name: next.workbookName });
  }

  for (const sheetName of previous.sheets.keys()) {
    if (!next.sheets.has(sheetName)) {
      ops.push({ kind: "deleteSheet", name: sheetName });
    }
  }

  for (const [sheetName, nextSheet] of next.sheets) {
    const previousSheet = previous.sheets.get(sheetName);
    ops.push({ kind: "upsertSheet", name: sheetName, order: nextSheet.order });

    if (!previousSheet) {
      nextSheet.cells.forEach((cell) => {
        const op: CommitOp = { kind: "upsertCell", sheetName, addr: cell.addr };
        if (cell.formula !== undefined) op.formula = cell.formula;
        if (cell.value !== undefined) op.value = cell.value;
        ops.push(op);
      });
      continue;
    }

    for (const [addr] of previousSheet.cells) {
      if (!nextSheet.cells.has(addr)) {
        ops.push({ kind: "deleteCell", sheetName, addr });
      }
    }

    for (const [addr, cell] of nextSheet.cells) {
      const previousCell = previousSheet.cells.get(addr);
      if (!previousCell || !sameCell(previousCell, cell)) {
        const op: CommitOp = { kind: "upsertCell", sheetName, addr };
        if (cell.formula !== undefined) op.formula = cell.formula;
        if (cell.value !== undefined) op.value = cell.value;
        ops.push(op);
      }
    }
  }

  return ops;
}
