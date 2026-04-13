import { deepStrictEqual } from "node:assert";
import fc from "fast-check";
import { formatAddress, parseCellAddress } from "@bilig/formula";
import type {
  CellRangeRef,
  WorkbookCommentThreadSnapshot,
  WorkbookConditionalFormatSnapshot,
  WorkbookDataValidationSnapshot,
  WorkbookNoteSnapshot,
  WorkbookRangeProtectionSnapshot,
  WorkbookSnapshot,
  WorkbookSortSnapshot,
} from "@bilig/protocol";
import { SpreadsheetEngine } from "../engine.js";
import { EngineMutationError } from "../engine/errors.js";
import {
  applyCoreAction,
  assertSnapshotInvariants,
  clearFormatActionArbitrary,
  clearStyleActionArbitrary,
  coreReplicaActionArbitrary,
  engineFuzzSheetName,
  normalizeSnapshotForSemanticComparison,
  rangeArbitrary,
  styleActionArbitrary,
  formatActionArbitrary,
  type CoreAction,
  type EngineSeedName,
} from "./engine-fuzz-helpers.js";

export const metadataSeedNames = [
  "cross-sheet-graph",
  "sparse-format",
  "annotation-rich",
  "named-structures",
  "validation-filter-sort",
  "structural-metadata",
  "pivot-analytics",
] as const satisfies readonly EngineSeedName[];

export const metadataSeedNameArbitrary = fc.constantFrom<EngineSeedName>(...metadataSeedNames);

const metadataRangeArbitrary = fc.oneof(
  rangeArbitrary,
  fc.constant({
    sheetName: engineFuzzSheetName,
    startAddress: "A1",
    endAddress: "C4",
  } satisfies CellRangeRef),
  fc.constant({
    sheetName: engineFuzzSheetName,
    startAddress: "B2",
    endAddress: "C5",
  } satisfies CellRangeRef),
);

const addressArbitrary = fc
  .record({
    row: fc.integer({ min: 0, max: 5 }),
    col: fc.integer({ min: 0, max: 5 }),
  })
  .map(({ row, col }) => formatAddress(row, col));

const freezePaneActionArbitrary = fc
  .record({
    rows: fc.integer({ min: 0, max: 2 }),
    cols: fc.integer({ min: 0, max: 2 }),
  })
  .filter(({ rows, cols }) => rows !== 0 || cols !== 0)
  .map(({ rows, cols }) => ({ kind: "setFreezePane" as const, rows, cols }));

const clearFreezePaneActionArbitrary = fc.constant({ kind: "clearFreezePane" as const });

const sheetProtectionActionArbitrary = fc
  .boolean()
  .map((hideFormulas) => ({ kind: "setSheetProtection" as const, hideFormulas }));

const clearSheetProtectionActionArbitrary = fc.constant({
  kind: "clearSheetProtection" as const,
});

const filterActionArbitrary = metadataRangeArbitrary.map((range) => ({
  kind: "setFilter" as const,
  range,
}));

const clearFilterActionArbitrary = metadataRangeArbitrary.map((range) => ({
  kind: "clearFilter" as const,
  range,
}));

const sortActionArbitrary = metadataRangeArbitrary.chain((range) => {
  const start = parseCellAddress(range.startAddress, range.sheetName);
  const end = parseCellAddress(range.endAddress, range.sheetName);
  return fc
    .record({
      keyCol: fc.integer({ min: start.col, max: end.col }),
      direction: fc.constantFrom<"asc" | "desc">("asc", "desc"),
    })
    .map(({ keyCol, direction }) => ({
      kind: "setSort" as const,
      range,
      keys: [{ keyAddress: formatAddress(start.row, keyCol), direction }],
    }));
});

const clearSortActionArbitrary = metadataRangeArbitrary.map((range) => ({
  kind: "clearSort" as const,
  range,
}));

const validationActionArbitrary = metadataRangeArbitrary.map((range) => ({
  kind: "setValidation" as const,
  validation: {
    range,
    rule: {
      kind: "list",
      values: ["Draft", "Final", "Review"],
    },
    allowBlank: false,
    showDropdown: true,
    errorStyle: "stop",
    errorTitle: "Required",
    errorMessage: "Pick one of the allowed statuses.",
  } satisfies WorkbookDataValidationSnapshot,
}));

const clearValidationActionArbitrary = metadataRangeArbitrary.map((range) => ({
  kind: "clearValidation" as const,
  range,
}));

const conditionalFormatActionArbitrary = metadataRangeArbitrary
  .filter((range) => range.startAddress !== range.endAddress)
  .chain((range) =>
    fc
      .record({
        id: fc.integer({ min: 1, max: 8 }),
        color: fc.constantFrom("#dcfce7", "#dbeafe", "#fee2e2"),
        operator: fc.constantFrom<"greaterThan" | "lessThan">("greaterThan", "lessThan"),
        threshold: fc.integer({ min: 0, max: 30 }),
      })
      .map(({ id, color, operator, threshold }) => ({
        kind: "setConditionalFormat" as const,
        format: {
          id: `cf-fuzz-${id}`,
          range,
          rule: {
            kind: "cellIs",
            operator,
            values: [threshold],
          },
          style: {
            fill: { backgroundColor: color },
          },
          priority: id,
        } satisfies WorkbookConditionalFormatSnapshot,
      })),
  );

const deleteConditionalFormatActionArbitrary = fc
  .integer({ min: 1, max: 8 })
  .map((id) => ({ kind: "deleteConditionalFormat" as const, id: `cf-fuzz-${id}` }));

const rangeProtectionActionArbitrary = metadataRangeArbitrary.map((range) => ({
  kind: "setRangeProtection" as const,
  protection: {
    id: `protect-${range.startAddress}-${range.endAddress}`.toLowerCase(),
    range,
    hideFormulas: true,
  } satisfies WorkbookRangeProtectionSnapshot,
}));

const deleteRangeProtectionActionArbitrary = metadataRangeArbitrary.map((range) => ({
  kind: "deleteRangeProtection" as const,
  id: `protect-${range.startAddress}-${range.endAddress}`.toLowerCase(),
}));

const commentThreadActionArbitrary = addressArbitrary.map((address) => ({
  kind: "setCommentThread" as const,
  thread: {
    threadId: `thread-${address.toLowerCase()}`,
    sheetName: engineFuzzSheetName,
    address,
    comments: [{ id: `comment-${address.toLowerCase()}`, body: `Review ${address}` }],
  } satisfies WorkbookCommentThreadSnapshot,
}));

const deleteCommentThreadActionArbitrary = addressArbitrary.map((address) => ({
  kind: "deleteCommentThread" as const,
  address,
}));

const noteActionArbitrary = addressArbitrary.map((address) => ({
  kind: "setNote" as const,
  note: {
    sheetName: engineFuzzSheetName,
    address,
    text: `Check ${address}`,
  } satisfies WorkbookNoteSnapshot,
}));

const deleteNoteActionArbitrary = addressArbitrary.map((address) => ({
  kind: "deleteNote" as const,
  address,
}));

export type MetadataAction =
  | { kind: "setFreezePane"; rows: number; cols: number }
  | { kind: "clearFreezePane" }
  | { kind: "setSheetProtection"; hideFormulas: boolean }
  | { kind: "clearSheetProtection" }
  | { kind: "setFilter"; range: CellRangeRef }
  | { kind: "clearFilter"; range: CellRangeRef }
  | { kind: "setSort"; range: CellRangeRef; keys: WorkbookSortSnapshot["keys"] }
  | { kind: "clearSort"; range: CellRangeRef }
  | { kind: "setValidation"; validation: WorkbookDataValidationSnapshot }
  | { kind: "clearValidation"; range: CellRangeRef }
  | { kind: "setConditionalFormat"; format: WorkbookConditionalFormatSnapshot }
  | { kind: "deleteConditionalFormat"; id: string }
  | { kind: "setRangeProtection"; protection: WorkbookRangeProtectionSnapshot }
  | { kind: "deleteRangeProtection"; id: string }
  | { kind: "setCommentThread"; thread: WorkbookCommentThreadSnapshot }
  | { kind: "deleteCommentThread"; address: string }
  | { kind: "setNote"; note: WorkbookNoteSnapshot }
  | { kind: "deleteNote"; address: string };

export type EngineSemanticAction = CoreAction | MetadataAction;

export const metadataActionArbitrary = fc.oneof<MetadataAction>(
  freezePaneActionArbitrary,
  clearFreezePaneActionArbitrary,
  sheetProtectionActionArbitrary,
  clearSheetProtectionActionArbitrary,
  filterActionArbitrary,
  clearFilterActionArbitrary,
  sortActionArbitrary,
  clearSortActionArbitrary,
  validationActionArbitrary,
  clearValidationActionArbitrary,
  conditionalFormatActionArbitrary,
  deleteConditionalFormatActionArbitrary,
  rangeProtectionActionArbitrary,
  deleteRangeProtectionActionArbitrary,
  commentThreadActionArbitrary,
  deleteCommentThreadActionArbitrary,
  noteActionArbitrary,
  deleteNoteActionArbitrary,
);

export const metadataSemanticActionArbitrary = fc.oneof<EngineSemanticAction>(
  styleActionArbitrary,
  clearStyleActionArbitrary,
  formatActionArbitrary,
  clearFormatActionArbitrary,
  metadataActionArbitrary,
);

export const snapshotSemanticActionArbitrary = fc.oneof<EngineSemanticAction>(
  coreReplicaActionArbitrary,
  metadataActionArbitrary,
);

function isCoreAction(action: EngineSemanticAction): action is CoreAction {
  return (
    action.kind === "values" ||
    action.kind === "formula" ||
    action.kind === "style" ||
    action.kind === "clearStyle" ||
    action.kind === "format" ||
    action.kind === "clearFormat" ||
    action.kind === "clear" ||
    action.kind === "fill" ||
    action.kind === "copy" ||
    action.kind === "move" ||
    action.kind === "insertRows" ||
    action.kind === "deleteRows" ||
    action.kind === "insertColumns" ||
    action.kind === "deleteColumns"
  );
}

export function applyMetadataAction(engine: SpreadsheetEngine, action: MetadataAction): void {
  switch (action.kind) {
    case "setFreezePane":
      engine.setFreezePane(engineFuzzSheetName, action.rows, action.cols);
      return;
    case "clearFreezePane":
      engine.clearFreezePane(engineFuzzSheetName);
      return;
    case "setSheetProtection":
      engine.setSheetProtection({
        sheetName: engineFuzzSheetName,
        hideFormulas: action.hideFormulas,
      });
      return;
    case "clearSheetProtection":
      engine.clearSheetProtection(engineFuzzSheetName);
      return;
    case "setFilter":
      engine.setFilter(engineFuzzSheetName, action.range);
      return;
    case "clearFilter":
      engine.clearFilter(engineFuzzSheetName, action.range);
      return;
    case "setSort":
      engine.setSort(engineFuzzSheetName, action.range, action.keys);
      return;
    case "clearSort":
      engine.clearSort(engineFuzzSheetName, action.range);
      return;
    case "setValidation":
      engine.setDataValidation(action.validation);
      return;
    case "clearValidation":
      engine.clearDataValidation(engineFuzzSheetName, action.range);
      return;
    case "setConditionalFormat":
      engine.setConditionalFormat(action.format);
      return;
    case "deleteConditionalFormat":
      engine.deleteConditionalFormat(action.id);
      return;
    case "setRangeProtection":
      engine.setRangeProtection(action.protection);
      return;
    case "deleteRangeProtection":
      engine.deleteRangeProtection(action.id);
      return;
    case "setCommentThread":
      engine.setCommentThread(action.thread);
      return;
    case "deleteCommentThread":
      engine.deleteCommentThread(engineFuzzSheetName, action.address);
      return;
    case "setNote":
      engine.setNote(action.note);
      return;
    case "deleteNote":
      engine.deleteNote(engineFuzzSheetName, action.address);
      return;
  }
}

export function applyEngineSemanticAction(
  engine: SpreadsheetEngine,
  action: EngineSemanticAction,
): void {
  if (isCoreAction(action)) {
    applyCoreAction(engine, action);
    return;
  }
  applyMetadataAction(engine, action);
}

export function applyEngineSemanticActionAndCaptureResult(
  engine: SpreadsheetEngine,
  action: EngineSemanticAction,
): { accepted: boolean; before: WorkbookSnapshot; after: WorkbookSnapshot } {
  const before = engine.exportSnapshot();
  try {
    applyEngineSemanticAction(engine, action);
    const after = engine.exportSnapshot();
    return {
      accepted: hasSemanticChange(before, after),
      before,
      after,
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

export async function exportEngineSemanticReplaySnapshot(
  initialSnapshot: WorkbookSnapshot,
  actions: readonly EngineSemanticAction[],
): Promise<WorkbookSnapshot> {
  const replay = new SpreadsheetEngine({
    workbookName: initialSnapshot.workbook.name,
    replicaId: `semantic-replay-${initialSnapshot.workbook.name}`,
  });
  await replay.ready();
  replay.importSnapshot(structuredClone(initialSnapshot));
  for (const action of actions) {
    applyEngineSemanticAction(replay, action);
  }
  return replay.exportSnapshot();
}

export function projectMetadataSnapshot(snapshot: WorkbookSnapshot): WorkbookSnapshot {
  const normalized = normalizeSnapshotForSemanticComparison(snapshot);
  return {
    version: normalized.version,
    workbook: structuredClone(normalized.workbook),
    sheets: normalized.sheets.map((sheet) =>
      Object.assign(
        {
          id: sheet.id,
          name: sheet.name,
          order: sheet.order,
          cells: [] as WorkbookSnapshot["sheets"][number]["cells"],
        },
        sheet.metadata ? { metadata: structuredClone(sheet.metadata) } : {},
      ),
    ),
  };
}

export function assertNoSemanticEmptyCells(snapshot: WorkbookSnapshot): void {
  assertSnapshotInvariants(snapshot);
  snapshot.sheets.forEach((sheet) => {
    sheet.cells.forEach((cell) => {
      if (cell.formula === undefined && cell.format === undefined && cell.value === null) {
        throw new Error(
          `Semantically empty tracked cell leaked into snapshot at ${sheet.name}!${cell.address}`,
        );
      }
    });
  });
}

function hasSemanticChange(before: WorkbookSnapshot, after: WorkbookSnapshot): boolean {
  try {
    deepStrictEqual(after, before);
    return false;
  } catch {
    return true;
  }
}
