import { SpreadsheetEngine } from "@bilig/core";
import { formatAddress, parseCellAddress } from "@bilig/formula";
import type { WorkbookAgentCommand, WorkbookAgentCommandBundle } from "@bilig/agent-api";
import type { WorkbookAxisEntrySnapshot } from "@bilig/protocol";
import type { EngineOp } from "@bilig/workbook-domain";
import type { WorkbookChangeUndoBundle } from "@bilig/zero-sync";

function normalizeFormula(formula: string): string {
  return formula.startsWith("=") ? formula.slice(1) : formula;
}

function toEngineUndoBundle(undoOps: readonly EngineOp[] | null): WorkbookChangeUndoBundle | null {
  if (!undoOps || undoOps.length === 0) {
    return null;
  }
  return {
    kind: "engineOps",
    ops: structuredClone([...undoOps]),
  };
}

function buildWriteRangeTransaction(
  command: Extract<WorkbookAgentCommand, { kind: "writeRange" }>,
): {
  readonly ops: EngineOp[];
  readonly potentialNewCells: number;
} {
  const start = parseCellAddress(command.startAddress, command.sheetName);
  const ops: EngineOp[] = [];
  let potentialNewCells = 0;
  command.values.forEach((rowValues, rowOffset) => {
    rowValues.forEach((cellInput, colOffset) => {
      const address = formatAddress(start.row + rowOffset, start.col + colOffset);
      if (cellInput === null) {
        ops.push({ kind: "clearCell", sheetName: command.sheetName, address });
        return;
      }
      if (
        typeof cellInput === "string" ||
        typeof cellInput === "number" ||
        typeof cellInput === "boolean"
      ) {
        ops.push({ kind: "setCellValue", sheetName: command.sheetName, address, value: cellInput });
        potentialNewCells += 1;
        return;
      }
      if ("formula" in cellInput) {
        ops.push({
          kind: "setCellFormula",
          sheetName: command.sheetName,
          address,
          formula: normalizeFormula(cellInput.formula),
        });
        potentialNewCells += 1;
        return;
      }
      ops.push({
        kind: "setCellValue",
        sheetName: command.sheetName,
        address,
        value: cellInput.value,
      });
      potentialNewCells += 1;
    });
  });
  return { ops, potentialNewCells };
}

function hasOwnProperty(value: object, key: string): boolean {
  return Object.prototype.hasOwnProperty.call(value, key);
}

function formatColumnLabel(index: number): string {
  return formatAddress(0, index).replace(/\d+/gu, "");
}

function formatRowSpanLabel(startRow: number, count: number): string {
  const first = startRow + 1;
  const last = startRow + count;
  return count === 1 ? `row ${String(first)}` : `rows ${String(first)}-${String(last)}`;
}

function formatColumnSpanLabel(startCol: number, count: number): string {
  const first = formatColumnLabel(startCol);
  const last = formatColumnLabel(startCol + count - 1);
  return count === 1 ? `column ${first}` : `columns ${first}-${last}`;
}

function getConsistentAxisEntrySize(input: {
  entries: readonly WorkbookAxisEntrySnapshot[];
  start: number;
  count: number;
  spanLabel: string;
  propertyLabel: string;
}): number | null {
  const entryByIndex = new Map(input.entries.map((entry) => [entry.index, entry]));
  let resolved: number | null | undefined;
  for (let index = input.start; index < input.start + input.count; index += 1) {
    const next = entryByIndex.get(index)?.size;
    const normalized = typeof next === "number" ? next : null;
    if (resolved === undefined) {
      resolved = normalized;
      continue;
    }
    if (resolved !== normalized) {
      throw new Error(
        `Cannot preserve ${input.propertyLabel} for ${input.spanLabel} because the existing ${input.propertyLabel} state is mixed. Specify ${input.propertyLabel} explicitly.`,
      );
    }
  }
  return resolved ?? null;
}

function getConsistentAxisEntryHidden(input: {
  entries: readonly WorkbookAxisEntrySnapshot[];
  start: number;
  count: number;
  spanLabel: string;
  propertyLabel: string;
}): boolean | null {
  const entryByIndex = new Map(input.entries.map((entry) => [entry.index, entry]));
  let resolved: boolean | null | undefined;
  for (let index = input.start; index < input.start + input.count; index += 1) {
    const next = entryByIndex.get(index)?.hidden;
    const normalized = typeof next === "boolean" ? next : null;
    if (resolved === undefined) {
      resolved = normalized;
      continue;
    }
    if (resolved !== normalized) {
      throw new Error(
        `Cannot preserve ${input.propertyLabel} for ${input.spanLabel} because the existing ${input.propertyLabel} state is mixed. Specify ${input.propertyLabel} explicitly.`,
      );
    }
  }
  return resolved ?? null;
}

function resolveRowMetadataCommandState(
  engine: SpreadsheetEngine,
  command: Extract<WorkbookAgentCommand, { kind: "updateRowMetadata" }>,
): {
  height: number | null;
  hidden: boolean | null;
} {
  const spanLabel = formatRowSpanLabel(command.startRow, command.count);
  const entries = engine.getRowAxisEntries(command.sheetName);
  return {
    height: hasOwnProperty(command, "height")
      ? (command.height ?? null)
      : getConsistentAxisEntrySize({
          entries,
          start: command.startRow,
          count: command.count,
          spanLabel,
          propertyLabel: "row height",
        }),
    hidden: hasOwnProperty(command, "hidden")
      ? (command.hidden ?? null)
      : getConsistentAxisEntryHidden({
          entries,
          start: command.startRow,
          count: command.count,
          spanLabel,
          propertyLabel: "row visibility",
        }),
  };
}

function resolveColumnMetadataCommandState(
  engine: SpreadsheetEngine,
  command: Extract<WorkbookAgentCommand, { kind: "updateColumnMetadata" }>,
): {
  width: number | null;
  hidden: boolean | null;
} {
  const spanLabel = formatColumnSpanLabel(command.startCol, command.count);
  const entries = engine.getColumnAxisEntries(command.sheetName);
  return {
    width: hasOwnProperty(command, "width")
      ? (command.width ?? null)
      : getConsistentAxisEntrySize({
          entries,
          start: command.startCol,
          count: command.count,
          spanLabel,
          propertyLabel: "column width",
        }),
    hidden: hasOwnProperty(command, "hidden")
      ? (command.hidden ?? null)
      : getConsistentAxisEntryHidden({
          entries,
          start: command.startCol,
          count: command.count,
          spanLabel,
          propertyLabel: "column visibility",
        }),
  };
}

function applyWorkbookAgentCommandWithUndoCapture(
  engine: SpreadsheetEngine,
  command: WorkbookAgentCommand,
): readonly EngineOp[] | null {
  switch (command.kind) {
    case "writeRange": {
      const transaction = buildWriteRangeTransaction(command);
      return engine.applyOps(transaction.ops, {
        captureUndo: true,
        potentialNewCells: transaction.potentialNewCells,
      });
    }
    case "formatRange": {
      const aggregatedUndoOps: EngineOp[] = [];
      if (command.patch !== undefined) {
        const undoOps = engine.captureUndoOps(() => {
          engine.setRangeStyle(command.range, command.patch!);
        }).undoOps;
        if (undoOps?.length) {
          aggregatedUndoOps.unshift(...undoOps);
        }
      }
      if (command.numberFormat !== undefined) {
        const undoOps = engine.captureUndoOps(() => {
          engine.setRangeNumberFormat(command.range, command.numberFormat!);
        }).undoOps;
        if (undoOps?.length) {
          aggregatedUndoOps.unshift(...undoOps);
        }
      }
      return aggregatedUndoOps;
    }
    case "clearRange":
      return engine.captureUndoOps(() => {
        engine.clearRange(command.range);
      }).undoOps;
    case "fillRange":
      return engine.captureUndoOps(() => {
        engine.fillRange(command.source, command.target);
      }).undoOps;
    case "copyRange":
      return engine.captureUndoOps(() => {
        engine.copyRange(command.source, command.target);
      }).undoOps;
    case "moveRange":
      return engine.captureUndoOps(() => {
        engine.moveRange(command.source, command.target);
      }).undoOps;
    case "createSheet":
      return engine.captureUndoOps(() => {
        engine.createSheet(command.name);
      }).undoOps;
    case "renameSheet":
      return engine.captureUndoOps(() => {
        engine.renameSheet(command.currentName, command.nextName);
      }).undoOps;
    case "updateRowMetadata": {
      const resolved = resolveRowMetadataCommandState(engine, command);
      return engine.captureUndoOps(() => {
        engine.updateRowMetadata(
          command.sheetName,
          command.startRow,
          command.count,
          resolved.height,
          resolved.hidden,
        );
      }).undoOps;
    }
    case "updateColumnMetadata": {
      const resolved = resolveColumnMetadataCommandState(engine, command);
      return engine.captureUndoOps(() => {
        engine.updateColumnMetadata(
          command.sheetName,
          command.startCol,
          command.count,
          resolved.width,
          resolved.hidden,
        );
      }).undoOps;
    }
    default: {
      const exhaustive: never = command;
      throw new Error(`Unhandled workbook agent command: ${JSON.stringify(exhaustive)}`);
    }
  }
}

export function applyWorkbookAgentCommandBundleWithUndoCapture(
  engine: SpreadsheetEngine,
  bundle: WorkbookAgentCommandBundle,
): WorkbookChangeUndoBundle | null {
  const aggregatedUndoOps: EngineOp[] = [];
  for (const command of bundle.commands) {
    const undoOps = applyWorkbookAgentCommandWithUndoCapture(engine, command);
    if (!undoOps || undoOps.length === 0) {
      continue;
    }
    aggregatedUndoOps.unshift(...undoOps);
  }
  return toEngineUndoBundle(aggregatedUndoOps);
}
