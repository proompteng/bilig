import { SpreadsheetEngine } from "@bilig/core";
import { formatAddress, parseCellAddress } from "@bilig/formula";
import {
  applyWorkbookAgentAnnotationCommand,
  isWorkbookAgentAnnotationCommand,
  applyWorkbookAgentConditionalFormatCommand,
  isWorkbookAgentConditionalFormatCommand,
  applyWorkbookAgentMediaCommand,
  isWorkbookAgentMediaCommand,
  applyWorkbookAgentObjectCommand,
  isWorkbookAgentObjectCommand,
  applyWorkbookAgentProtectionCommand,
  isWorkbookAgentProtectionCommand,
  applyWorkbookAgentStructuralCommand,
  isWorkbookAgentStructuralCommand,
  applyWorkbookAgentValidationCommand,
  isWorkbookAgentValidationCommand,
  type WorkbookAgentCommand,
  type WorkbookAgentCommandBundle,
} from "@bilig/agent-api";
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

function applyWorkbookAgentCommandWithUndoCapture(
  engine: SpreadsheetEngine,
  command: WorkbookAgentCommand,
): readonly EngineOp[] | null {
  if (isWorkbookAgentStructuralCommand(command)) {
    return engine.captureUndoOps(() => {
      applyWorkbookAgentStructuralCommand(engine, command);
    }).undoOps;
  }
  if (isWorkbookAgentObjectCommand(command)) {
    return engine.captureUndoOps(() => {
      applyWorkbookAgentObjectCommand(engine, command);
    }).undoOps;
  }
  if (isWorkbookAgentMediaCommand(command)) {
    return engine.captureUndoOps(() => {
      applyWorkbookAgentMediaCommand(engine, command);
    }).undoOps;
  }
  if (isWorkbookAgentProtectionCommand(command)) {
    return engine.captureUndoOps(() => {
      applyWorkbookAgentProtectionCommand(engine, command);
    }).undoOps;
  }
  if (isWorkbookAgentConditionalFormatCommand(command)) {
    return engine.captureUndoOps(() => {
      applyWorkbookAgentConditionalFormatCommand(engine, command);
    }).undoOps;
  }
  if (isWorkbookAgentAnnotationCommand(command)) {
    return engine.captureUndoOps(() => {
      applyWorkbookAgentAnnotationCommand(engine, command);
    }).undoOps;
  }
  if (isWorkbookAgentValidationCommand(command)) {
    return engine.captureUndoOps(() => {
      applyWorkbookAgentValidationCommand(engine, command);
    }).undoOps;
  }
  switch (command.kind) {
    case "writeRange": {
      const transaction = buildWriteRangeTransaction(command);
      return engine.applyOps(transaction.ops, {
        captureUndo: true,
        potentialNewCells: transaction.potentialNewCells,
      });
    }
    case "setRangeFormulas":
      return engine.captureUndoOps(() => {
        engine.setRangeFormulas(
          command.range,
          command.formulas.map((row) => row.map((formula) => normalizeFormula(formula))),
        );
      }).undoOps;
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
