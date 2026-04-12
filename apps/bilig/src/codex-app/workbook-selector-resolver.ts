import { formatAddress, parseCellAddress } from "@bilig/formula";
import type {
  CellRangeRef,
  WorkbookDefinedNameSnapshot,
  WorkbookTableSnapshot,
} from "@bilig/protocol";
import type { WorkbookAgentUiContext } from "@bilig/contracts";
import { z } from "zod";
import type { WorkbookRuntime } from "../workbook-runtime/runtime-manager.js";

const selectorRevisionShape = {
  revision: z.number().int().positive().optional(),
} as const;

const a1RangeSelectorSchema = z.object({
  kind: z.literal("a1Range"),
  sheet: z.string().trim().min(1),
  start: z.string().trim().min(1),
  end: z.string().trim().min(1),
  ...selectorRevisionShape,
});

const namedRangeSelectorSchema = z.object({
  kind: z.literal("namedRange"),
  name: z.string().trim().min(1),
  ...selectorRevisionShape,
});

const tableSelectorSchema = z.object({
  kind: z.literal("table"),
  table: z.string().trim().min(1),
  sheet: z.string().trim().min(1).optional(),
  ...selectorRevisionShape,
});

const tableColumnSelectorSchema = z.object({
  kind: z.literal("tableColumn"),
  table: z.string().trim().min(1),
  column: z.string().trim().min(1),
  sheet: z.string().trim().min(1).optional(),
  ...selectorRevisionShape,
});

const currentSelectionSelectorSchema = z.object({
  kind: z.literal("currentSelection"),
  ...selectorRevisionShape,
});

const currentRegionSelectorSchema = z.object({
  kind: z.literal("currentRegion"),
  anchor: z
    .object({
      sheet: z.string().trim().min(1),
      address: z.string().trim().min(1),
    })
    .optional(),
  ...selectorRevisionShape,
});

const visibleRowsSelectorSchema = z.object({
  kind: z.literal("visibleRows"),
  sheet: z.string().trim().min(1).optional(),
  ...selectorRevisionShape,
});

export const workbookSemanticSelectorSchema = z.discriminatedUnion("kind", [
  a1RangeSelectorSchema,
  namedRangeSelectorSchema,
  tableSelectorSchema,
  tableColumnSelectorSchema,
  currentSelectionSelectorSchema,
  currentRegionSelectorSchema,
  visibleRowsSelectorSchema,
]);

export type WorkbookSemanticSelector = z.infer<typeof workbookSemanticSelectorSchema>;

export type WorkbookSelectorResolutionErrorCode =
  | "selector_not_found"
  | "selector_ambiguous"
  | "selector_type_mismatch"
  | "selector_revision_stale"
  | "selector_hidden_by_filter"
  | "selector_blocked_by_protection";

export class WorkbookSelectorResolutionError extends Error {
  readonly code: WorkbookSelectorResolutionErrorCode;

  constructor(code: WorkbookSelectorResolutionErrorCode, message: string) {
    super(message);
    this.code = code;
    this.name = "WorkbookSelectorResolutionError";
  }
}

export interface ResolvedWorkbookSelector {
  readonly selector: WorkbookSemanticSelector;
  readonly resolvedRevision: number;
  readonly objectType: "range" | "table" | "tableColumn";
  readonly derivedA1Ranges: readonly CellRangeRef[];
  readonly displayLabel: string;
  readonly table: WorkbookTableSnapshot | null;
  readonly namedRange: WorkbookDefinedNameSnapshot | null;
}

export interface ResolveWorkbookSelectorInput {
  readonly runtime: WorkbookRuntime;
  readonly selector: WorkbookSemanticSelector;
  readonly uiContext: WorkbookAgentUiContext | null;
}

function normalizeRange(range: CellRangeRef): CellRangeRef {
  const start = parseCellAddress(range.startAddress, range.sheetName);
  const end = parseCellAddress(range.endAddress, range.sheetName);
  const startRow = Math.min(start.row, end.row);
  const endRow = Math.max(start.row, end.row);
  const startCol = Math.min(start.col, end.col);
  const endCol = Math.max(start.col, end.col);
  return {
    sheetName: range.sheetName,
    startAddress: formatAddress(startRow, startCol),
    endAddress: formatAddress(endRow, endCol),
  };
}

function createRangeRef(
  sheetName: string,
  startRow: number,
  startCol: number,
  endRow: number,
  endCol: number,
): CellRangeRef {
  return normalizeRange({
    sheetName,
    startAddress: formatAddress(startRow, startCol),
    endAddress: formatAddress(endRow, endCol),
  });
}

function throwResolutionError(code: WorkbookSelectorResolutionErrorCode, message: string): never {
  throw new WorkbookSelectorResolutionError(code, message);
}

function assertSelectorRevision(
  runtime: WorkbookRuntime,
  selector: WorkbookSemanticSelector,
): void {
  if (selector.revision === undefined || selector.revision === runtime.headRevision) {
    return;
  }
  throwResolutionError(
    "selector_revision_stale",
    `Selector revision ${String(selector.revision)} does not match workbook revision ${String(runtime.headRevision)}`,
  );
}

function resolveUiSelection(uiContext: WorkbookAgentUiContext | null): CellRangeRef {
  if (!uiContext) {
    throwResolutionError(
      "selector_not_found",
      "No browser workbook context is attached to this chat session",
    );
  }
  return normalizeRange({
    sheetName: uiContext.selection.sheetName,
    startAddress: uiContext.selection.range?.startAddress ?? uiContext.selection.address,
    endAddress: uiContext.selection.range?.endAddress ?? uiContext.selection.address,
  });
}

function resolveUiVisibleRows(
  runtime: WorkbookRuntime,
  uiContext: WorkbookAgentUiContext | null,
  sheetName: string | undefined,
): CellRangeRef {
  if (!uiContext) {
    throwResolutionError(
      "selector_not_found",
      "Visible-row selectors require attached browser workbook context",
    );
  }
  const contextSheetName = uiContext.selection.sheetName;
  if (sheetName && sheetName !== contextSheetName) {
    throwResolutionError(
      "selector_not_found",
      `Visible rows are only available for the current sheet ${contextSheetName}`,
    );
  }
  const usedRange = getSheetUsedRange(runtime, contextSheetName);
  const colStart = usedRange?.startCol ?? uiContext.viewport.colStart;
  const colEnd = usedRange?.endCol ?? uiContext.viewport.colEnd;
  return createRangeRef(
    contextSheetName,
    uiContext.viewport.rowStart,
    colStart,
    uiContext.viewport.rowEnd,
    colEnd,
  );
}

function getSheetUsedRange(
  runtime: WorkbookRuntime,
  sheetName: string,
): {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
} | null {
  const sheet = runtime.engine.exportSnapshot().sheets.find((entry) => entry.name === sheetName);
  if (!sheet || sheet.cells.length === 0) {
    return null;
  }
  let startRow = Number.POSITIVE_INFINITY;
  let endRow = Number.NEGATIVE_INFINITY;
  let startCol = Number.POSITIVE_INFINITY;
  let endCol = Number.NEGATIVE_INFINITY;
  for (const cell of sheet.cells) {
    const parsed = parseCellAddress(cell.address, sheetName);
    startRow = Math.min(startRow, parsed.row);
    endRow = Math.max(endRow, parsed.row);
    startCol = Math.min(startCol, parsed.col);
    endCol = Math.max(endCol, parsed.col);
  }
  return { startRow, endRow, startCol, endCol };
}

function createPopulatedCellKey(row: number, col: number): string {
  return `${String(row)}:${String(col)}`;
}

function resolveCurrentRegion(
  runtime: WorkbookRuntime,
  anchor: {
    sheet: string;
    address: string;
  },
): CellRangeRef {
  const sheet = runtime.engine.exportSnapshot().sheets.find((entry) => entry.name === anchor.sheet);
  if (!sheet) {
    throwResolutionError("selector_not_found", `Sheet ${anchor.sheet} does not exist`);
  }
  const parsedAnchor = parseCellAddress(anchor.address, anchor.sheet);
  const populated = new Set(
    sheet.cells.map((cell) => {
      const parsed = parseCellAddress(cell.address, sheet.name);
      return createPopulatedCellKey(parsed.row, parsed.col);
    }),
  );
  let startRow = parsedAnchor.row;
  let endRow = parsedAnchor.row;
  let startCol = parsedAnchor.col;
  let endCol = parsedAnchor.col;

  const rowHasData = (row: number, left: number, right: number): boolean => {
    for (let col = left; col <= right; col += 1) {
      if (populated.has(createPopulatedCellKey(row, col))) {
        return true;
      }
    }
    return false;
  };

  const colHasData = (col: number, top: number, bottom: number): boolean => {
    for (let row = top; row <= bottom; row += 1) {
      if (populated.has(createPopulatedCellKey(row, col))) {
        return true;
      }
    }
    return false;
  };

  while (rowHasData(startRow - 1, startCol, endCol)) {
    startRow -= 1;
  }
  while (rowHasData(endRow + 1, startCol, endCol)) {
    endRow += 1;
  }
  while (colHasData(startCol - 1, startRow, endRow)) {
    startCol -= 1;
  }
  while (colHasData(endCol + 1, startRow, endRow)) {
    endCol += 1;
  }

  return createRangeRef(anchor.sheet, startRow, startCol, endRow, endCol);
}

function findTable(
  runtime: WorkbookRuntime,
  tableName: string,
  sheetName?: string,
): WorkbookTableSnapshot {
  const matches = runtime.engine
    .getTables()
    .filter(
      (table) =>
        table.name.trim().toUpperCase() === tableName.trim().toUpperCase() &&
        (sheetName === undefined ||
          table.sheetName.trim().toUpperCase() === sheetName.trim().toUpperCase()),
    );
  if (matches.length === 0) {
    throwResolutionError(
      "selector_not_found",
      sheetName
        ? `Table ${tableName} does not exist on sheet ${sheetName}`
        : `Table ${tableName} does not exist`,
    );
  }
  if (matches.length > 1) {
    throwResolutionError(
      "selector_ambiguous",
      `Table ${tableName} resolves to multiple sheet matches`,
    );
  }
  return matches[0]!;
}

function resolveTableColumnRange(table: WorkbookTableSnapshot, columnName: string): CellRangeRef {
  const columnIndex = table.columnNames.findIndex(
    (value) => value.trim().toUpperCase() === columnName.trim().toUpperCase(),
  );
  if (columnIndex === -1) {
    throwResolutionError(
      "selector_not_found",
      `Column ${columnName} does not exist on table ${table.name}`,
    );
  }
  const start = parseCellAddress(table.startAddress, table.sheetName);
  const end = parseCellAddress(table.endAddress, table.sheetName);
  const column = start.col + columnIndex;
  const dataStartRow = start.row + (table.headerRow ? 1 : 0);
  const dataEndRow = end.row - (table.totalsRow ? 1 : 0);
  if (dataStartRow <= dataEndRow) {
    return createRangeRef(table.sheetName, dataStartRow, column, dataEndRow, column);
  }
  return createRangeRef(table.sheetName, start.row, column, start.row, column);
}

function resolveDefinedNameRange(
  runtime: WorkbookRuntime,
  definedName: WorkbookDefinedNameSnapshot,
): {
  readonly range: CellRangeRef;
  readonly objectType: "range" | "tableColumn";
  readonly table: WorkbookTableSnapshot | null;
} {
  const value = definedName.value;
  if (typeof value !== "object" || value === null || !("kind" in value)) {
    throwResolutionError(
      "selector_type_mismatch",
      `Named range ${definedName.name} is a scalar and does not resolve to workbook cells`,
    );
  }
  switch (value.kind) {
    case "cell-ref":
      return {
        range: createRangeRef(
          value.sheetName,
          parseCellAddress(value.address, value.sheetName).row,
          parseCellAddress(value.address, value.sheetName).col,
          parseCellAddress(value.address, value.sheetName).row,
          parseCellAddress(value.address, value.sheetName).col,
        ),
        objectType: "range",
        table: null,
      };
    case "range-ref":
      return {
        range: normalizeRange({
          sheetName: value.sheetName,
          startAddress: value.startAddress,
          endAddress: value.endAddress,
        }),
        objectType: "range",
        table: null,
      };
    case "structured-ref": {
      const table = findTable(runtime, value.tableName);
      return {
        range: resolveTableColumnRange(table, value.columnName),
        objectType: "tableColumn",
        table,
      };
    }
    case "scalar":
    case "formula":
      throwResolutionError(
        "selector_type_mismatch",
        `Named range ${definedName.name} is ${value.kind} metadata, not a cell range`,
      );
  }
}

function resolveNamedRange(
  runtime: WorkbookRuntime,
  selector: z.infer<typeof namedRangeSelectorSchema>,
): ResolvedWorkbookSelector {
  const namedRange = runtime.engine
    .getDefinedNames()
    .find((entry) => entry.name.trim().toUpperCase() === selector.name.trim().toUpperCase());
  if (!namedRange) {
    throwResolutionError("selector_not_found", `Named range ${selector.name} does not exist`);
  }
  const resolved = resolveDefinedNameRange(runtime, namedRange);
  return {
    selector,
    resolvedRevision: runtime.headRevision,
    objectType: resolved.objectType,
    derivedA1Ranges: [resolved.range],
    displayLabel: namedRange.name,
    table: resolved.table,
    namedRange,
  };
}

export function listWorkbookNamedRanges(runtime: WorkbookRuntime): readonly {
  readonly name: string;
  readonly valueKind: string;
  readonly displayLabel: string;
  readonly range: CellRangeRef | null;
  readonly tableName: string | null;
  readonly columnName: string | null;
}[] {
  return runtime.engine.getDefinedNames().map((namedRange) => {
    const value = namedRange.value;
    if (typeof value !== "object" || value === null || !("kind" in value)) {
      return {
        name: namedRange.name,
        valueKind: "scalar",
        displayLabel: namedRange.name,
        range: null,
        tableName: null,
        columnName: null,
      };
    }
    switch (value.kind) {
      case "cell-ref":
        return {
          name: namedRange.name,
          valueKind: value.kind,
          displayLabel: namedRange.name,
          range: createRangeRef(
            value.sheetName,
            parseCellAddress(value.address, value.sheetName).row,
            parseCellAddress(value.address, value.sheetName).col,
            parseCellAddress(value.address, value.sheetName).row,
            parseCellAddress(value.address, value.sheetName).col,
          ),
          tableName: null,
          columnName: null,
        };
      case "range-ref":
        return {
          name: namedRange.name,
          valueKind: value.kind,
          displayLabel: namedRange.name,
          range: normalizeRange({
            sheetName: value.sheetName,
            startAddress: value.startAddress,
            endAddress: value.endAddress,
          }),
          tableName: null,
          columnName: null,
        };
      case "structured-ref":
        return {
          name: namedRange.name,
          valueKind: value.kind,
          displayLabel: `${value.tableName}[${value.columnName}]`,
          range: null,
          tableName: value.tableName,
          columnName: value.columnName,
        };
      case "scalar":
      case "formula":
        return {
          name: namedRange.name,
          valueKind: value.kind,
          displayLabel: namedRange.name,
          range: null,
          tableName: null,
          columnName: null,
        };
    }
  });
}

export function listWorkbookTables(runtime: WorkbookRuntime): readonly WorkbookTableSnapshot[] {
  return runtime.engine.getTables().map((table) => ({
    name: table.name,
    sheetName: table.sheetName,
    startAddress: table.startAddress,
    endAddress: table.endAddress,
    columnNames: [...table.columnNames],
    headerRow: table.headerRow,
    totalsRow: table.totalsRow,
  }));
}

export function resolveWorkbookSelector(
  input: ResolveWorkbookSelectorInput,
): ResolvedWorkbookSelector {
  assertSelectorRevision(input.runtime, input.selector);
  switch (input.selector.kind) {
    case "a1Range": {
      return {
        selector: input.selector,
        resolvedRevision: input.runtime.headRevision,
        objectType: "range",
        derivedA1Ranges: [
          normalizeRange({
            sheetName: input.selector.sheet,
            startAddress: input.selector.start,
            endAddress: input.selector.end,
          }),
        ],
        displayLabel: `${input.selector.sheet}!${input.selector.start}:${input.selector.end}`,
        table: null,
        namedRange: null,
      };
    }
    case "namedRange":
      return resolveNamedRange(input.runtime, input.selector);
    case "table": {
      const table = findTable(input.runtime, input.selector.table, input.selector.sheet);
      return {
        selector: input.selector,
        resolvedRevision: input.runtime.headRevision,
        objectType: "table",
        derivedA1Ranges: [
          normalizeRange({
            sheetName: table.sheetName,
            startAddress: table.startAddress,
            endAddress: table.endAddress,
          }),
        ],
        displayLabel: table.name,
        table,
        namedRange: null,
      };
    }
    case "tableColumn": {
      const table = findTable(input.runtime, input.selector.table, input.selector.sheet);
      return {
        selector: input.selector,
        resolvedRevision: input.runtime.headRevision,
        objectType: "tableColumn",
        derivedA1Ranges: [resolveTableColumnRange(table, input.selector.column)],
        displayLabel: `${table.name}[${input.selector.column}]`,
        table,
        namedRange: null,
      };
    }
    case "currentSelection":
      return {
        selector: input.selector,
        resolvedRevision: input.runtime.headRevision,
        objectType: "range",
        derivedA1Ranges: [resolveUiSelection(input.uiContext)],
        displayLabel: "Current selection",
        table: null,
        namedRange: null,
      };
    case "currentRegion": {
      const anchor =
        input.selector.anchor ??
        (input.uiContext
          ? {
              sheet: input.uiContext.selection.sheetName,
              address: input.uiContext.selection.address,
            }
          : null);
      if (!anchor) {
        throwResolutionError(
          "selector_not_found",
          "Current-region selectors require either an explicit anchor or browser workbook context",
        );
      }
      return {
        selector: input.selector,
        resolvedRevision: input.runtime.headRevision,
        objectType: "range",
        derivedA1Ranges: [resolveCurrentRegion(input.runtime, anchor)],
        displayLabel: `Current region around ${anchor.sheet}!${anchor.address}`,
        table: null,
        namedRange: null,
      };
    }
    case "visibleRows":
      return {
        selector: input.selector,
        resolvedRevision: input.runtime.headRevision,
        objectType: "range",
        derivedA1Ranges: [
          resolveUiVisibleRows(input.runtime, input.uiContext, input.selector.sheet),
        ],
        displayLabel: `Visible rows on ${input.selector.sheet ?? input.uiContext?.selection.sheetName ?? "current sheet"}`,
        table: null,
        namedRange: null,
      };
  }
}

export function resolveWorkbookSelectorToSingleRange(input: ResolveWorkbookSelectorInput): {
  readonly resolution: ResolvedWorkbookSelector;
  readonly range: CellRangeRef;
} {
  const resolution = resolveWorkbookSelector(input);
  if (resolution.derivedA1Ranges.length !== 1) {
    throwResolutionError(
      "selector_type_mismatch",
      `Selector ${resolution.displayLabel} resolved to ${String(resolution.derivedA1Ranges.length)} ranges`,
    );
  }
  return {
    resolution,
    range: resolution.derivedA1Ranges[0]!,
  };
}
