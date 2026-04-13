import { deepStrictEqual } from "node:assert";
import { describe, expect, it } from "vitest";
import fc from "fast-check";
import type { AsyncCommand } from "fast-check";
import { formatAddress } from "@bilig/formula";
import type {
  CellNumberFormatInput,
  CellRangeRef,
  CellStylePatch,
  LiteralInput,
  WorkbookSnapshot,
} from "@bilig/protocol";
import { SpreadsheetEngine } from "../engine.js";
import { EngineMutationError } from "../engine/errors.js";
import { runModelProperty, runProperty } from "@bilig/test-fuzz";

type CoreAction =
  | { kind: "values"; range: CellRangeRef; values: LiteralInput[][] }
  | { kind: "formula"; address: string; formula: string }
  | { kind: "style"; range: CellRangeRef; patch: CellStylePatch }
  | { kind: "format"; range: CellRangeRef; format: CellNumberFormatInput }
  | { kind: "clear"; range: CellRangeRef }
  | { kind: "fill"; source: CellRangeRef; target: CellRangeRef }
  | { kind: "copy"; source: CellRangeRef; target: CellRangeRef }
  | { kind: "move"; source: CellRangeRef; target: CellRangeRef }
  | { kind: "insertRows"; start: number; count: number }
  | { kind: "deleteRows"; start: number; count: number }
  | { kind: "insertColumns"; start: number; count: number }
  | { kind: "deleteColumns"; start: number; count: number };

const sheetName = "Sheet1";
const workbookName = "fuzz-core-book";

function toRangeRef(
  targetSheetName: string,
  startRow: number,
  startCol: number,
  endRow: number,
  endCol: number,
): CellRangeRef {
  return {
    sheetName: targetSheetName,
    startAddress: formatAddress(startRow, startCol),
    endAddress: formatAddress(endRow, endCol),
  };
}

function buildValueMatrix(
  height: number,
  width: number,
  values: readonly LiteralInput[],
): LiteralInput[][] {
  const rows: LiteralInput[][] = [];
  let offset = 0;
  for (let row = 0; row < height; row += 1) {
    const nextRow: LiteralInput[] = [];
    for (let col = 0; col < width; col += 1) {
      nextRow.push(values[offset] ?? null);
      offset += 1;
    }
    rows.push(nextRow);
  }
  return rows;
}

function assertSnapshotInvariants(snapshot: WorkbookSnapshot): void {
  const sheetNames = snapshot.sheets.map((sheet) => sheet.name);
  expect(new Set(sheetNames).size).toBe(sheetNames.length);
  snapshot.sheets.forEach((sheet) => {
    const addresses = sheet.cells.map((cell) => cell.address);
    expect(new Set(addresses).size).toBe(addresses.length);
  });
}

function normalizeSnapshotForSemanticComparison(snapshot: WorkbookSnapshot): WorkbookSnapshot {
  const clone = structuredClone(snapshot);
  clone.sheets = clone.sheets.map((sheet) => {
    if (!sheet.metadata) {
      return sheet;
    }
    const metadata = { ...sheet.metadata };
    if (metadata.rows) {
      metadata.rows = metadata.rows.map(({ id: _id, ...rest }) => rest);
    }
    if (metadata.columns) {
      metadata.columns = metadata.columns.map(({ id: _id, ...rest }) => rest);
    }
    return {
      ...sheet,
      metadata,
    };
  });
  return clone;
}

const literalInputArbitrary = fc.oneof<LiteralInput>(
  fc.integer({ min: -10_000, max: 10_000 }),
  fc.boolean(),
  fc.constantFrom("north", "south", "ready", "done"),
  fc.constant(null),
);
const rangeSeedArbitrary = fc.record({
  startRow: fc.integer({ min: 0, max: 5 }),
  startCol: fc.integer({ min: 0, max: 5 }),
  height: fc.integer({ min: 1, max: 2 }),
  width: fc.integer({ min: 1, max: 2 }),
});
const rangeArbitrary = rangeSeedArbitrary.map((value) =>
  toRangeRef(
    sheetName,
    value.startRow,
    value.startCol,
    value.startRow + value.height - 1,
    value.startCol + value.width - 1,
  ),
);
const sameSizedRangePairArbitrary = rangeSeedArbitrary.chain((source) =>
  fc
    .record({
      targetStartRow: fc.integer({ min: 0, max: 6 - source.height }),
      targetStartCol: fc.integer({ min: 0, max: 6 - source.width }),
    })
    .map(({ targetStartRow, targetStartCol }) => ({
      source: toRangeRef(
        sheetName,
        source.startRow,
        source.startCol,
        source.startRow + source.height - 1,
        source.startCol + source.width - 1,
      ),
      target: toRangeRef(
        sheetName,
        targetStartRow,
        targetStartCol,
        targetStartRow + source.height - 1,
        targetStartCol + source.width - 1,
      ),
    })),
);
const formulaArbitrary = fc
  .tuple(
    fc.constantFrom("A1", "B2", "C3", "D4", "E5", "F6"),
    fc.constantFrom("+", "-", "*", "/"),
    fc.constantFrom("A1", "B2", "C3", "D4", "E5", "F6"),
  )
  .map(([left, operator, right]) => `${left}${operator}${right}`);
const stylePatchArbitrary = fc.constantFrom<CellStylePatch>(
  { fill: { backgroundColor: "#dbeafe" } },
  { font: { bold: true } },
  { alignment: { horizontal: "right", wrap: true } },
);
const formatInputArbitrary = fc.constantFrom<CellNumberFormatInput>(
  "0.00",
  { kind: "currency", currency: "USD", decimals: 2 },
  { kind: "percent", decimals: 1 },
  { kind: "text" },
);
const valuesActionArbitrary = rangeSeedArbitrary.chain((range) =>
  fc
    .array(literalInputArbitrary, {
      minLength: range.height * range.width,
      maxLength: range.height * range.width,
    })
    .map((values) => ({
      kind: "values" as const,
      range: toRangeRef(
        sheetName,
        range.startRow,
        range.startCol,
        range.startRow + range.height - 1,
        range.startCol + range.width - 1,
      ),
      values: buildValueMatrix(range.height, range.width, values),
    })),
);
const formulaActionArbitrary = fc
  .record({
    row: fc.integer({ min: 0, max: 5 }),
    col: fc.integer({ min: 0, max: 5 }),
    formula: formulaArbitrary,
  })
  .map(({ row, col, formula }) => ({
    kind: "formula" as const,
    address: formatAddress(row, col),
    formula,
  }));
const styleActionArbitrary = fc
  .record({ range: rangeArbitrary, patch: stylePatchArbitrary })
  .map(({ range, patch }) => ({ kind: "style" as const, range, patch }));
const formatActionArbitrary = fc
  .record({ range: rangeArbitrary, format: formatInputArbitrary })
  .map(({ range, format }) => ({ kind: "format" as const, range, format }));
const clearActionArbitrary = rangeArbitrary.map((range) => ({ kind: "clear" as const, range }));
const fillActionArbitrary = sameSizedRangePairArbitrary.map(({ source, target }) => ({
  kind: "fill" as const,
  source,
  target,
}));
const copyActionArbitrary = sameSizedRangePairArbitrary.map(({ source, target }) => ({
  kind: "copy" as const,
  source,
  target,
}));
const moveActionArbitrary = sameSizedRangePairArbitrary.map(({ source, target }) => ({
  kind: "move" as const,
  source,
  target,
}));
const axisMutationArbitrary = fc.record({
  start: fc.integer({ min: 0, max: 4 }),
  count: fc.integer({ min: 1, max: 2 }),
});
const insertRowsActionArbitrary = axisMutationArbitrary.map(({ start, count }) => ({
  kind: "insertRows" as const,
  start,
  count,
}));
const deleteRowsActionArbitrary = axisMutationArbitrary.map(({ start, count }) => ({
  kind: "deleteRows" as const,
  start,
  count,
}));
const insertColumnsActionArbitrary = axisMutationArbitrary.map(({ start, count }) => ({
  kind: "insertColumns" as const,
  start,
  count,
}));
const deleteColumnsActionArbitrary = axisMutationArbitrary.map(({ start, count }) => ({
  kind: "deleteColumns" as const,
  start,
  count,
}));
const coreActionArbitrary = fc.oneof<CoreAction>(
  valuesActionArbitrary,
  formulaActionArbitrary,
  styleActionArbitrary,
  formatActionArbitrary,
  clearActionArbitrary,
  fillActionArbitrary,
  copyActionArbitrary,
  moveActionArbitrary,
  insertRowsActionArbitrary,
  deleteRowsActionArbitrary,
  insertColumnsActionArbitrary,
  deleteColumnsActionArbitrary,
);

function applyCoreAction(engine: SpreadsheetEngine, action: CoreAction): void {
  switch (action.kind) {
    case "values":
      engine.setRangeValues(action.range, action.values);
      break;
    case "formula":
      engine.setCellFormula(sheetName, action.address, action.formula);
      break;
    case "style":
      engine.setRangeStyle(action.range, action.patch);
      break;
    case "format":
      engine.setRangeNumberFormat(action.range, action.format);
      break;
    case "clear":
      engine.clearRange(action.range);
      break;
    case "fill":
      engine.fillRange(action.source, action.target);
      break;
    case "copy":
      engine.copyRange(action.source, action.target);
      break;
    case "move":
      engine.moveRange(action.source, action.target);
      break;
    case "insertRows":
      engine.insertRows(sheetName, action.start, action.count);
      break;
    case "deleteRows":
      engine.deleteRows(sheetName, action.start, action.count);
      break;
    case "insertColumns":
      engine.insertColumns(sheetName, action.start, action.count);
      break;
    case "deleteColumns":
      engine.deleteColumns(sheetName, action.start, action.count);
      break;
  }
}

function applyActionAndCaptureResult(
  engine: SpreadsheetEngine,
  action: CoreAction,
): { accepted: boolean; before: WorkbookSnapshot; after: WorkbookSnapshot } {
  const before = engine.exportSnapshot();
  try {
    applyCoreAction(engine, action);
    return {
      accepted: true,
      before,
      after: engine.exportSnapshot(),
    };
  } catch (error) {
    if (!(error instanceof EngineMutationError)) {
      throw error;
    }
    const after = engine.exportSnapshot();
    deepStrictEqual(after, before);
    return {
      accepted: false,
      before,
      after,
    };
  }
}

function exportReplaySnapshot(actions: readonly CoreAction[]): WorkbookSnapshot {
  const replay = new SpreadsheetEngine({
    workbookName,
    replicaId: "fuzz-core-replay",
  });
  replay.createSheet(sheetName);
  for (const action of actions) {
    applyCoreAction(replay, action);
  }
  return replay.exportSnapshot();
}

interface EngineHistoryModel {
  applied: CoreAction[];
  undone: CoreAction[];
}

function currentModelSnapshot(model: EngineHistoryModel): WorkbookSnapshot {
  return exportReplaySnapshot(model.applied);
}

function applyAndExpectSnapshot(
  engine: SpreadsheetEngine,
  model: EngineHistoryModel,
  _actionDescription: string,
): void {
  const snapshot = engine.exportSnapshot();
  assertSnapshotInvariants(snapshot);
  expect(normalizeSnapshotForSemanticComparison(snapshot)).toEqual(
    normalizeSnapshotForSemanticComparison(currentModelSnapshot(model)),
  );
}

function applyActionCommandArbitrary(
  actionArbitrary: fc.Arbitrary<CoreAction>,
): fc.Arbitrary<AsyncCommand<EngineHistoryModel, SpreadsheetEngine>> {
  return actionArbitrary.map((action) => ({
    check: () => true,
    run: async (model, real) => {
      const result = applyActionAndCaptureResult(real, action);
      if (result.accepted) {
        model.applied.push(action);
        model.undone = [];
      }
      applyAndExpectSnapshot(real, model, `apply ${action.kind}`);
    },
    toString: () => `apply(${action.kind})`,
  }));
}

const undoCommandArbitrary: fc.Arbitrary<AsyncCommand<EngineHistoryModel, SpreadsheetEngine>> =
  fc.constant({
    check: (model) => model.applied.length > 0,
    run: async (model, real) => {
      if (!real.undo()) {
        throw new Error("Undo command failed despite a non-empty applied history");
      }
      const action = model.applied.pop();
      if (!action) {
        throw new Error("Undo command ran with no applied action");
      }
      model.undone.push(action);
      applyAndExpectSnapshot(real, model, "undo");
    },
    toString: () => "undo()",
  });

const redoCommandArbitrary: fc.Arbitrary<AsyncCommand<EngineHistoryModel, SpreadsheetEngine>> =
  fc.constant({
    check: (model) => model.undone.length > 0,
    run: async (model, real) => {
      if (!real.redo()) {
        throw new Error("Redo command failed despite a non-empty undone history");
      }
      const action = model.undone.pop();
      if (!action) {
        throw new Error("Redo command ran with no undone action");
      }
      model.applied.push(action);
      applyAndExpectSnapshot(real, model, "redo");
    },
    toString: () => "redo()",
  });

const engineHistoryCommandArbitraries: Array<
  fc.Arbitrary<AsyncCommand<EngineHistoryModel, SpreadsheetEngine>>
> = [
  applyActionCommandArbitrary(valuesActionArbitrary),
  applyActionCommandArbitrary(formulaActionArbitrary),
  applyActionCommandArbitrary(styleActionArbitrary),
  applyActionCommandArbitrary(formatActionArbitrary),
  applyActionCommandArbitrary(clearActionArbitrary),
  applyActionCommandArbitrary(fillActionArbitrary),
  applyActionCommandArbitrary(copyActionArbitrary),
  applyActionCommandArbitrary(moveActionArbitrary),
  applyActionCommandArbitrary(insertRowsActionArbitrary),
  applyActionCommandArbitrary(deleteRowsActionArbitrary),
  applyActionCommandArbitrary(insertColumnsActionArbitrary),
  applyActionCommandArbitrary(deleteColumnsActionArbitrary),
  undoCommandArbitrary,
  redoCommandArbitrary,
];

describe("engine fuzz", () => {
  it("preserves replay parity and snapshot roundtrips across rich random command streams", async () => {
    await runProperty({
      suite: "core/rich-command-replay-roundtrip",
      arbitrary: fc.array(coreActionArbitrary, { minLength: 5, maxLength: 16 }),
      predicate: async (actions) => {
        const engine = new SpreadsheetEngine({
          workbookName,
          replicaId: "fuzz-core",
        });
        engine.createSheet(sheetName);

        const applied: CoreAction[] = [];
        for (const action of actions) {
          const result = applyActionAndCaptureResult(engine, action);
          if (result.accepted) {
            applied.push(action);
          }

          const snapshot = result.after;
          assertSnapshotInvariants(snapshot);
          expect(exportReplaySnapshot(applied)).toEqual(snapshot);

          const restored = new SpreadsheetEngine({
            workbookName: snapshot.workbook.name,
            replicaId: "fuzz-core-restore",
          });
          restored.importSnapshot(snapshot);
          expect(restored.exportSnapshot()).toEqual(snapshot);
        }
      },
    });
  });

  it("keeps model-based history semantics aligned with replayed workbook state", async () => {
    const ran = await runModelProperty({
      suite: "core/model-history-undo-redo",
      commands: fc.commands(engineHistoryCommandArbitraries, { maxCommands: 18 }),
      createModel: () => ({
        applied: [],
        undone: [],
      }),
      createReal: async () => {
        const engine = new SpreadsheetEngine({
          workbookName,
          replicaId: "fuzz-core-model",
        });
        engine.createSheet(sheetName);
        applyAndExpectSnapshot(engine, { applied: [], undone: [] }, "initial model state");
        return engine;
      },
    });
    expect(ran).toBe(true);
  });
});
