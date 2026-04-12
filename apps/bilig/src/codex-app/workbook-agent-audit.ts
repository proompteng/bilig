import { SpreadsheetEngine } from "@bilig/core";
import { formatAddress, parseCellAddress, translateFormulaReferences } from "@bilig/formula";
import type { CellRangeRef, SheetMetadataSnapshot, WorkbookSnapshot } from "@bilig/protocol";
import { isDeepStrictEqual } from "node:util";
import type { WorkbookRuntime } from "../workbook-runtime/runtime-manager.js";
import {
  findWorkbookFormulaIssues,
  summarizeWorkbookStructure,
  type WorkbookFormulaIssue,
} from "./workbook-agent-comprehension.js";

const DEFAULT_AUDIT_LIMIT = 50;
const MAX_AUDIT_LIMIT = 200;
const DEFAULT_HIDDEN_PRECEDENT_DEPTH = 4;
const MAX_HIDDEN_PRECEDENT_DEPTH = 6;
const MAX_HIDDEN_PRECEDENT_NODES = 160;
const MIN_INCONSISTENT_GROUP_SIZE = 3;

interface CellBounds {
  minRow: number;
  maxRow: number;
  minCol: number;
  maxCol: number;
}

interface FormulaCellRef {
  sheetName: string;
  address: string;
  formula: string;
  row: number;
  col: number;
}

interface FormulaRunGroup {
  axis: "row" | "column";
  sheetName: string;
  cells: readonly FormulaCellRef[];
}

interface MetadataBoundsDriver {
  source:
    | "styleRange"
    | "formatRange"
    | "filter"
    | "sort"
    | "validation"
    | "conditionalFormat"
    | "protectedRange"
    | "commentThread"
    | "note"
    | "table"
    | "pivot"
    | "chart"
    | "image"
    | "shape"
    | "spill"
    | "definedName"
    | "rowMetadata"
    | "columnMetadata";
  startAddress: string;
  endAddress: string;
}

interface InvariantProblem {
  code: string;
  message: string;
  sheetName?: string;
  address?: string;
}

export interface WorkbookBrokenReferenceReport {
  summary: {
    scannedFormulaCells: number;
    brokenReferenceCount: number;
    truncated: boolean;
  };
  issues: WorkbookFormulaIssue[];
}

export interface WorkbookHiddenRowDependencyHit {
  sheetName: string;
  address: string;
  rowNumber: number;
  depth: number;
}

export interface WorkbookHiddenRowDependencyIssue {
  sheetName: string;
  address: string;
  formula: string;
  hiddenPrecedentCount: number;
  hiddenPrecedents: WorkbookHiddenRowDependencyHit[];
}

export interface WorkbookHiddenRowDependencyReport {
  summary: {
    scannedFormulaCells: number;
    affectedFormulaCount: number;
    hiddenPrecedentCount: number;
    truncated: boolean;
  };
  issues: WorkbookHiddenRowDependencyIssue[];
}

export interface WorkbookInconsistentFormulaOutlier {
  address: string;
  actualFormula: string;
  expectedFormula: string;
}

export interface WorkbookInconsistentFormulaGroupReport {
  axis: "row" | "column";
  sheetName: string;
  groupRange: CellRangeRef;
  formulaCellCount: number;
  dominantFormula: string;
  dominantCount: number;
  outliers: WorkbookInconsistentFormulaOutlier[];
}

export interface WorkbookInconsistentFormulaReport {
  summary: {
    scannedFormulaCells: number;
    inconsistentGroupCount: number;
    outlierCount: number;
    truncated: boolean;
  };
  groups: WorkbookInconsistentFormulaGroupReport[];
}

export interface WorkbookUsedRangeBloatSheetReport {
  sheetName: string;
  populatedCellCount: number;
  populatedRange: CellRangeRef | null;
  compositeRange: CellRangeRef;
  extraRows: number;
  extraColumns: number;
  populatedArea: number;
  compositeArea: number;
  bloatArea: number;
  drivers: MetadataBoundsDriver[];
}

export interface WorkbookUsedRangeBloatReport {
  summary: {
    scannedSheetCount: number;
    bloatedSheetCount: number;
    totalBloatArea: number;
    truncated: boolean;
  };
  sheets: WorkbookUsedRangeBloatSheetReport[];
}

export interface WorkbookPerformanceHotspot {
  sheetName: string;
  cellCount: number;
  formulaCellCount: number;
  jsOnlyFormulaCount: number;
  issueCount: number;
  pivotCount: number;
  spillCount: number;
  usedRange: {
    startAddress: string;
    endAddress: string;
  } | null;
  reasons: string[];
}

export interface WorkbookPerformanceHotspotReport {
  summary: {
    scannedSheetCount: number;
    hotspotCount: number;
    truncated: boolean;
    recalcMetrics: ReturnType<WorkbookRuntime["engine"]["getLastMetrics"]>;
  };
  hotspots: WorkbookPerformanceHotspot[];
}

export interface WorkbookInvariantVerificationReport {
  summary: {
    ok: boolean;
    problemCount: number;
    roundTripChecked: boolean;
    roundTripStable: boolean;
  };
  problems: Array<{
    code: string;
    message: string;
    sheetName?: string;
    address?: string;
  }>;
}

function clampLimit(limit: number | undefined): number {
  if (!Number.isFinite(limit) || typeof limit !== "number") {
    return DEFAULT_AUDIT_LIMIT;
  }
  return Math.max(1, Math.min(MAX_AUDIT_LIMIT, Math.trunc(limit)));
}

function clampHiddenDepth(depth: number | undefined): number {
  if (!Number.isFinite(depth) || typeof depth !== "number") {
    return DEFAULT_HIDDEN_PRECEDENT_DEPTH;
  }
  return Math.max(1, Math.min(MAX_HIDDEN_PRECEDENT_DEPTH, Math.trunc(depth)));
}

function boundsFromAddress(sheetName: string, address: string): CellBounds {
  const parsed = parseCellAddress(address, sheetName);
  return {
    minRow: parsed.row,
    maxRow: parsed.row,
    minCol: parsed.col,
    maxCol: parsed.col,
  };
}

function boundsFromRange(range: CellRangeRef): CellBounds {
  const start = parseCellAddress(range.startAddress, range.sheetName);
  const end = parseCellAddress(range.endAddress, range.sheetName);
  return {
    minRow: Math.min(start.row, end.row),
    maxRow: Math.max(start.row, end.row),
    minCol: Math.min(start.col, end.col),
    maxCol: Math.max(start.col, end.col),
  };
}

function mergeBounds(base: CellBounds | null, next: CellBounds | null): CellBounds | null {
  if (!next) {
    return base;
  }
  if (!base) {
    return next;
  }
  return {
    minRow: Math.min(base.minRow, next.minRow),
    maxRow: Math.max(base.maxRow, next.maxRow),
    minCol: Math.min(base.minCol, next.minCol),
    maxCol: Math.max(base.maxCol, next.maxCol),
  };
}

function boundsToRange(sheetName: string, bounds: CellBounds): CellRangeRef {
  return {
    sheetName,
    startAddress: formatAddress(bounds.minRow, bounds.minCol),
    endAddress: formatAddress(bounds.maxRow, bounds.maxCol),
  };
}

function boundsArea(bounds: CellBounds | null): number {
  if (!bounds) {
    return 0;
  }
  return (bounds.maxRow - bounds.minRow + 1) * (bounds.maxCol - bounds.minCol + 1);
}

function collectFormulaCells(
  snapshot: WorkbookSnapshot,
  sheetName?: string,
): readonly FormulaCellRef[] {
  const cells: FormulaCellRef[] = [];
  for (const sheet of snapshot.sheets) {
    if (sheetName !== undefined && sheet.name !== sheetName) {
      continue;
    }
    for (const cell of sheet.cells) {
      if (!cell.formula) {
        continue;
      }
      const parsed = parseCellAddress(cell.address, sheet.name);
      cells.push({
        sheetName: sheet.name,
        address: cell.address,
        formula: cell.formula,
        row: parsed.row,
        col: parsed.col,
      });
    }
  }
  return cells;
}

function splitQualifiedAddress(qualifiedAddress: string): {
  sheetName: string;
  address: string;
} {
  const separator = qualifiedAddress.lastIndexOf("!");
  if (separator <= 0 || separator >= qualifiedAddress.length - 1) {
    throw new Error(`Invalid qualified workbook address: ${qualifiedAddress}`);
  }
  return {
    sheetName: qualifiedAddress.slice(0, separator),
    address: qualifiedAddress.slice(separator + 1),
  };
}

function buildHiddenRowIntervals(runtime: WorkbookRuntime): Map<string, Array<[number, number]>> {
  const hiddenRows = new Map<string, Array<[number, number]>>();
  const snapshot = runtime.engine.exportSnapshot();
  for (const sheet of snapshot.sheets) {
    const intervals = runtime.engine
      .getRowMetadata(sheet.name)
      .filter((entry) => entry.hidden === true)
      .map((entry) => [entry.start, entry.start + entry.count - 1] as [number, number]);
    if (intervals.length > 0) {
      hiddenRows.set(sheet.name, intervals);
    }
  }
  return hiddenRows;
}

function isRowHidden(hiddenRows: readonly [number, number][], row: number): boolean {
  return hiddenRows.some(([start, end]) => row >= start && row <= end);
}

function collectHiddenPrecedents(
  runtime: WorkbookRuntime,
  hiddenRowsBySheet: ReadonlyMap<string, readonly [number, number][]>,
  rootSheetName: string,
  rootAddress: string,
  maxDepth: number,
): WorkbookHiddenRowDependencyHit[] {
  const hiddenHits: WorkbookHiddenRowDependencyHit[] = [];
  const queue: Array<{ qualifiedAddress: string; depth: number }> = [
    { qualifiedAddress: `${rootSheetName}!${rootAddress}`, depth: 0 },
  ];
  const visited = new Set<string>(queue.map((entry) => entry.qualifiedAddress));
  let visitedNodes = 0;
  while (queue.length > 0) {
    const current = queue.shift();
    if (!current || current.depth >= maxDepth) {
      continue;
    }
    const { sheetName, address } = splitQualifiedAddress(current.qualifiedAddress);
    const dependencies = runtime.engine.getDependencies(sheetName, address);
    for (const precedent of dependencies.directPrecedents) {
      if (visited.has(precedent)) {
        continue;
      }
      visited.add(precedent);
      visitedNodes += 1;
      if (visitedNodes > MAX_HIDDEN_PRECEDENT_NODES) {
        return hiddenHits;
      }
      const parsed = splitQualifiedAddress(precedent);
      const location = parseCellAddress(parsed.address, parsed.sheetName);
      const hiddenRows = hiddenRowsBySheet.get(parsed.sheetName);
      if (hiddenRows && isRowHidden(hiddenRows, location.row)) {
        hiddenHits.push({
          sheetName: parsed.sheetName,
          address: parsed.address,
          rowNumber: location.row + 1,
          depth: current.depth + 1,
        });
      }
      queue.push({
        qualifiedAddress: precedent,
        depth: current.depth + 1,
      });
    }
  }
  return hiddenHits;
}

function buildFormulaRunGroups(
  snapshot: WorkbookSnapshot,
  axis: "row" | "column",
  sheetName?: string,
): readonly FormulaRunGroup[] {
  const groups: FormulaRunGroup[] = [];
  for (const sheet of snapshot.sheets) {
    if (sheetName !== undefined && sheet.name !== sheetName) {
      continue;
    }
    const formulaCells = collectFormulaCells(
      {
        ...snapshot,
        sheets: [sheet],
      },
      sheet.name,
    );
    const cellsByAxis = new Map<number, FormulaCellRef[]>();
    for (const cell of formulaCells) {
      const key = axis === "column" ? cell.col : cell.row;
      const entries = cellsByAxis.get(key) ?? [];
      entries.push(cell);
      cellsByAxis.set(key, entries);
    }
    for (const entries of cellsByAxis.values()) {
      const sorted = [...entries].toSorted((left, right) =>
        axis === "column" ? left.row - right.row : left.col - right.col,
      );
      let current: FormulaCellRef[] = [];
      let lastIndex = Number.NaN;
      for (const cell of sorted) {
        const index = axis === "column" ? cell.row : cell.col;
        if (current.length === 0 || index === lastIndex + 1) {
          current.push(cell);
          lastIndex = index;
          continue;
        }
        if (current.length >= MIN_INCONSISTENT_GROUP_SIZE) {
          groups.push({
            axis,
            sheetName: sheet.name,
            cells: [...current],
          });
        }
        current = [cell];
        lastIndex = index;
      }
      if (current.length >= MIN_INCONSISTENT_GROUP_SIZE) {
        groups.push({
          axis,
          sheetName: sheet.name,
          cells: [...current],
        });
      }
    }
  }
  return groups;
}

function normalizeFormulaForAnchor(formula: string, rowDelta: number, colDelta: number): string {
  try {
    return translateFormulaReferences(formula, rowDelta, colDelta);
  } catch {
    return `__raw__:${formula}`;
  }
}

function translateNormalizedFormula(
  normalizedFormula: string,
  rowDelta: number,
  colDelta: number,
  fallbackFormula: string,
): string {
  if (normalizedFormula.startsWith("__raw__:")) {
    return `=${fallbackFormula}`;
  }
  try {
    return `=${translateFormulaReferences(normalizedFormula, rowDelta, colDelta)}`;
  } catch {
    return `=${fallbackFormula}`;
  }
}

function summarizeGroupRange(group: FormulaRunGroup): CellRangeRef {
  const first = group.cells[0];
  const last = group.cells[group.cells.length - 1];
  if (!first || !last) {
    throw new Error("Formula run group must contain at least one cell");
  }
  return {
    sheetName: group.sheetName,
    startAddress: first.address,
    endAddress: last.address,
  };
}

function collectSheetBloatDrivers(
  snapshot: WorkbookSnapshot,
  sheetName: string,
  metadata: SheetMetadataSnapshot | undefined,
): {
  bounds: CellBounds | null;
  rowExtent: CellBounds | null;
  columnExtent: CellBounds | null;
  drivers: MetadataBoundsDriver[];
} {
  let bounds: CellBounds | null = null;
  let rowExtent: CellBounds | null = null;
  let columnExtent: CellBounds | null = null;
  const drivers: MetadataBoundsDriver[] = [];

  const addRangeDriver = (source: MetadataBoundsDriver["source"], range: CellRangeRef) => {
    bounds = mergeBounds(bounds, boundsFromRange(range));
    drivers.push({
      source,
      startAddress: range.startAddress,
      endAddress: range.endAddress,
    });
  };

  const addAddressDriver = (
    source: MetadataBoundsDriver["source"],
    address: string,
    endAddress = address,
  ) => {
    const range = {
      sheetName,
      startAddress: address,
      endAddress,
    } satisfies CellRangeRef;
    addRangeDriver(source, range);
  };

  for (const styleRange of metadata?.styleRanges ?? []) {
    addRangeDriver("styleRange", styleRange.range);
  }
  for (const formatRange of metadata?.formatRanges ?? []) {
    addRangeDriver("formatRange", formatRange.range);
  }
  for (const filter of metadata?.filters ?? []) {
    addRangeDriver("filter", filter);
  }
  for (const sort of metadata?.sorts ?? []) {
    addRangeDriver("sort", sort.range);
  }
  for (const validation of metadata?.validations ?? []) {
    addRangeDriver("validation", validation.range);
  }
  for (const conditionalFormat of metadata?.conditionalFormats ?? []) {
    addRangeDriver("conditionalFormat", conditionalFormat.range);
  }
  for (const protectedRange of metadata?.protectedRanges ?? []) {
    addRangeDriver("protectedRange", protectedRange.range);
  }
  for (const commentThread of metadata?.commentThreads ?? []) {
    addAddressDriver("commentThread", commentThread.address);
  }
  for (const note of metadata?.notes ?? []) {
    addAddressDriver("note", note.address);
  }

  for (const table of snapshot.workbook.metadata?.tables ?? []) {
    if (table.sheetName === sheetName) {
      addRangeDriver("table", {
        sheetName,
        startAddress: table.startAddress,
        endAddress: table.endAddress,
      });
    }
  }
  for (const pivot of snapshot.workbook.metadata?.pivots ?? []) {
    if (pivot.sheetName !== sheetName) {
      continue;
    }
    const start = parseCellAddress(pivot.address, sheetName);
    addAddressDriver(
      "pivot",
      pivot.address,
      formatAddress(
        start.row + Math.max(0, pivot.rows - 1),
        start.col + Math.max(0, pivot.cols - 1),
      ),
    );
  }
  for (const chart of snapshot.workbook.metadata?.charts ?? []) {
    if (chart.sheetName === sheetName) {
      const start = parseCellAddress(chart.address, sheetName);
      addAddressDriver(
        "chart",
        chart.address,
        formatAddress(
          start.row + Math.max(0, chart.rows - 1),
          start.col + Math.max(0, chart.cols - 1),
        ),
      );
    }
    if (chart.source.sheetName === sheetName) {
      addRangeDriver("chart", chart.source);
    }
  }
  for (const image of snapshot.workbook.metadata?.images ?? []) {
    if (image.sheetName !== sheetName) {
      continue;
    }
    const start = parseCellAddress(image.address, sheetName);
    addAddressDriver(
      "image",
      image.address,
      formatAddress(
        start.row + Math.max(0, image.rows - 1),
        start.col + Math.max(0, image.cols - 1),
      ),
    );
  }
  for (const shape of snapshot.workbook.metadata?.shapes ?? []) {
    if (shape.sheetName !== sheetName) {
      continue;
    }
    const start = parseCellAddress(shape.address, sheetName);
    addAddressDriver(
      "shape",
      shape.address,
      formatAddress(
        start.row + Math.max(0, shape.rows - 1),
        start.col + Math.max(0, shape.cols - 1),
      ),
    );
  }
  for (const spill of snapshot.workbook.metadata?.spills ?? []) {
    if (spill.sheetName !== sheetName) {
      continue;
    }
    const start = parseCellAddress(spill.address, sheetName);
    addAddressDriver(
      "spill",
      spill.address,
      formatAddress(
        start.row + Math.max(0, spill.rows - 1),
        start.col + Math.max(0, spill.cols - 1),
      ),
    );
  }
  for (const definedName of snapshot.workbook.metadata?.definedNames ?? []) {
    const value = definedName.value;
    if (typeof value !== "object" || value === null) {
      continue;
    }
    if (value.kind === "cell-ref" && value.sheetName === sheetName) {
      addAddressDriver("definedName", value.address);
    }
    if (value.kind === "range-ref" && value.sheetName === sheetName) {
      addRangeDriver("definedName", {
        sheetName,
        startAddress: value.startAddress,
        endAddress: value.endAddress,
      });
    }
  }

  const rowRecords = [...(metadata?.rowMetadata ?? []), ...(metadata?.rows ?? [])];
  for (const record of rowRecords) {
    const start = "index" in record ? record.index : record.start;
    const count = "count" in record ? record.count : 1;
    const next = {
      minRow: start,
      maxRow: start + count - 1,
      minCol: 0,
      maxCol: 0,
    } satisfies CellBounds;
    rowExtent = mergeBounds(rowExtent, next);
  }
  if (rowExtent) {
    drivers.push({
      source: "rowMetadata",
      startAddress: formatAddress(rowExtent.minRow, 0),
      endAddress: formatAddress(rowExtent.maxRow, 0),
    });
  }

  const columnRecords = [...(metadata?.columnMetadata ?? []), ...(metadata?.columns ?? [])];
  for (const record of columnRecords) {
    const start = "index" in record ? record.index : record.start;
    const count = "count" in record ? record.count : 1;
    const next = {
      minRow: 0,
      maxRow: 0,
      minCol: start,
      maxCol: start + count - 1,
    } satisfies CellBounds;
    columnExtent = mergeBounds(columnExtent, next);
  }
  if (columnExtent) {
    drivers.push({
      source: "columnMetadata",
      startAddress: formatAddress(0, columnExtent.minCol),
      endAddress: formatAddress(0, columnExtent.maxCol),
    });
  }

  return {
    bounds,
    rowExtent,
    columnExtent,
    drivers,
  };
}

function buildContentBounds(sheet: WorkbookSnapshot["sheets"][number]): CellBounds | null {
  let bounds: CellBounds | null = null;
  for (const cell of sheet.cells) {
    bounds = mergeBounds(bounds, boundsFromAddress(sheet.name, cell.address));
  }
  return bounds;
}

function addProblem(
  problems: InvariantProblem[],
  code: string,
  message: string,
  options: Pick<InvariantProblem, "sheetName" | "address"> = {},
): void {
  const problem: InvariantProblem = {
    code,
    message,
  };
  if (options.sheetName !== undefined) {
    problem.sheetName = options.sheetName;
  }
  if (options.address !== undefined) {
    problem.address = options.address;
  }
  problems.push(problem);
}

function ensureSheetExists(
  sheetNames: ReadonlySet<string>,
  problems: InvariantProblem[],
  sheetName: string,
  context: string,
): boolean {
  if (sheetNames.has(sheetName)) {
    return true;
  }
  addProblem(problems, "missingSheet", `${context} references missing sheet ${sheetName}`);
  return false;
}

function validateRangeRef(
  problems: InvariantProblem[],
  expectedSheetName: string | undefined,
  range: CellRangeRef,
  context: string,
): void {
  if (expectedSheetName !== undefined && range.sheetName !== expectedSheetName) {
    addProblem(
      problems,
      "rangeSheetMismatch",
      `${context} range points at ${range.sheetName} instead of ${expectedSheetName}`,
      { sheetName: expectedSheetName },
    );
  }
  try {
    parseCellAddress(range.startAddress, range.sheetName);
    parseCellAddress(range.endAddress, range.sheetName);
  } catch (error) {
    addProblem(
      problems,
      "invalidRangeAddress",
      `${context} range is invalid: ${error instanceof Error ? error.message : String(error)}`,
      { sheetName: expectedSheetName ?? range.sheetName },
    );
  }
}

function validateAddressRef(
  problems: InvariantProblem[],
  sheetName: string,
  address: string,
  context: string,
): void {
  try {
    parseCellAddress(address, sheetName);
  } catch (error) {
    addProblem(
      problems,
      "invalidCellAddress",
      `${context} address is invalid: ${error instanceof Error ? error.message : String(error)}`,
      { sheetName, address },
    );
  }
}

export function scanWorkbookBrokenReferences(
  runtime: WorkbookRuntime,
  input: {
    sheetName?: string | undefined;
    limit?: number | undefined;
  } = {},
): WorkbookBrokenReferenceReport {
  const limit = clampLimit(input.limit);
  const report = findWorkbookFormulaIssues(runtime, {
    ...(input.sheetName !== undefined ? { sheetName: input.sheetName } : {}),
    limit: MAX_AUDIT_LIMIT,
  });
  const issues = report.issues.filter((issue) => issue.errorText === "#REF!");
  return {
    summary: {
      scannedFormulaCells: report.summary.scannedFormulaCells,
      brokenReferenceCount: issues.length,
      truncated: issues.length > limit,
    },
    issues: issues.slice(0, limit),
  };
}

export function scanWorkbookHiddenRowsAffectingResults(
  runtime: WorkbookRuntime,
  input: {
    sheetName?: string | undefined;
    limit?: number | undefined;
    depth?: number | undefined;
  } = {},
): WorkbookHiddenRowDependencyReport {
  const snapshot = runtime.engine.exportSnapshot();
  const hiddenRowsBySheet = buildHiddenRowIntervals(runtime);
  const depth = clampHiddenDepth(input.depth);
  const limit = clampLimit(input.limit);
  const issues: WorkbookHiddenRowDependencyIssue[] = [];
  let scannedFormulaCells = 0;
  let hiddenPrecedentCount = 0;

  for (const cell of collectFormulaCells(snapshot, input.sheetName)) {
    scannedFormulaCells += 1;
    const hits = collectHiddenPrecedents(
      runtime,
      hiddenRowsBySheet,
      cell.sheetName,
      cell.address,
      depth,
    );
    if (hits.length === 0) {
      continue;
    }
    hiddenPrecedentCount += hits.length;
    issues.push({
      sheetName: cell.sheetName,
      address: cell.address,
      formula: `=${cell.formula}`,
      hiddenPrecedentCount: hits.length,
      hiddenPrecedents: hits.slice(0, 10),
    });
  }

  const sortedIssues = issues.toSorted((left, right) => {
    if (left.hiddenPrecedentCount !== right.hiddenPrecedentCount) {
      return right.hiddenPrecedentCount - left.hiddenPrecedentCount;
    }
    if (left.sheetName !== right.sheetName) {
      return left.sheetName.localeCompare(right.sheetName);
    }
    return left.address.localeCompare(right.address, undefined, { numeric: true });
  });

  return {
    summary: {
      scannedFormulaCells,
      affectedFormulaCount: sortedIssues.length,
      hiddenPrecedentCount,
      truncated: sortedIssues.length > limit,
    },
    issues: sortedIssues.slice(0, limit),
  };
}

export function scanWorkbookInconsistentFormulas(
  runtime: WorkbookRuntime,
  input: {
    sheetName?: string | undefined;
    limit?: number | undefined;
  } = {},
): WorkbookInconsistentFormulaReport {
  const snapshot = runtime.engine.exportSnapshot();
  const limit = clampLimit(input.limit);
  const groups: WorkbookInconsistentFormulaGroupReport[] = [];
  const candidateGroups = [
    ...buildFormulaRunGroups(snapshot, "column", input.sheetName),
    ...buildFormulaRunGroups(snapshot, "row", input.sheetName),
  ];

  for (const group of candidateGroups) {
    const anchor = group.cells[0];
    if (!anchor) {
      continue;
    }
    const signatureMap = new Map<
      string,
      {
        count: number;
        representative: FormulaCellRef;
      }
    >();
    const signatures = group.cells.map((cell) => {
      const signature = normalizeFormulaForAnchor(
        cell.formula,
        anchor.row - cell.row,
        anchor.col - cell.col,
      );
      const existing = signatureMap.get(signature);
      if (existing) {
        existing.count += 1;
      } else {
        signatureMap.set(signature, {
          count: 1,
          representative: cell,
        });
      }
      return {
        cell,
        signature,
      };
    });

    const dominantEntry = [...signatureMap.entries()].toSorted((left, right) => {
      if (left[1].count !== right[1].count) {
        return right[1].count - left[1].count;
      }
      return left[1].representative.address.localeCompare(
        right[1].representative.address,
        undefined,
        {
          numeric: true,
        },
      );
    })[0];
    if (!dominantEntry || dominantEntry[1].count < 2) {
      continue;
    }
    const [dominantSignature, dominantData] = dominantEntry;
    const outliers = signatures
      .filter((entry) => entry.signature !== dominantSignature)
      .map((entry) => ({
        address: entry.cell.address,
        actualFormula: `=${entry.cell.formula}`,
        expectedFormula: translateNormalizedFormula(
          dominantSignature,
          entry.cell.row - anchor.row,
          entry.cell.col - anchor.col,
          dominantData.representative.formula,
        ),
      }));
    if (outliers.length === 0) {
      continue;
    }
    groups.push({
      axis: group.axis,
      sheetName: group.sheetName,
      groupRange: summarizeGroupRange(group),
      formulaCellCount: group.cells.length,
      dominantFormula: `=${dominantData.representative.formula}`,
      dominantCount: dominantData.count,
      outliers,
    });
  }

  const sortedGroups = groups.toSorted((left, right) => {
    if (left.outliers.length !== right.outliers.length) {
      return right.outliers.length - left.outliers.length;
    }
    if (left.formulaCellCount !== right.formulaCellCount) {
      return right.formulaCellCount - left.formulaCellCount;
    }
    if (left.sheetName !== right.sheetName) {
      return left.sheetName.localeCompare(right.sheetName);
    }
    return left.groupRange.startAddress.localeCompare(right.groupRange.startAddress, undefined, {
      numeric: true,
    });
  });

  return {
    summary: {
      scannedFormulaCells: collectFormulaCells(snapshot, input.sheetName).length,
      inconsistentGroupCount: sortedGroups.length,
      outlierCount: sortedGroups.reduce((sum, group) => sum + group.outliers.length, 0),
      truncated: sortedGroups.length > limit,
    },
    groups: sortedGroups.slice(0, limit),
  };
}

export function scanWorkbookUsedRangeBloat(
  runtime: WorkbookRuntime,
  input: {
    sheetName?: string | undefined;
    limit?: number | undefined;
  } = {},
): WorkbookUsedRangeBloatReport {
  const snapshot = runtime.engine.exportSnapshot();
  const limit = clampLimit(input.limit);
  const reports: WorkbookUsedRangeBloatSheetReport[] = [];

  for (const sheet of snapshot.sheets) {
    if (input.sheetName !== undefined && sheet.name !== input.sheetName) {
      continue;
    }
    const contentBounds = buildContentBounds(sheet);
    const drivers = collectSheetBloatDrivers(snapshot, sheet.name, sheet.metadata);
    let compositeBounds = mergeBounds(contentBounds, drivers.bounds);
    if (compositeBounds && drivers.rowExtent) {
      compositeBounds = {
        ...compositeBounds,
        minRow: Math.min(compositeBounds.minRow, drivers.rowExtent.minRow),
        maxRow: Math.max(compositeBounds.maxRow, drivers.rowExtent.maxRow),
      };
    } else if (!compositeBounds && drivers.rowExtent && drivers.columnExtent) {
      compositeBounds = {
        minRow: drivers.rowExtent.minRow,
        maxRow: drivers.rowExtent.maxRow,
        minCol: drivers.columnExtent.minCol,
        maxCol: drivers.columnExtent.maxCol,
      };
    }
    if (compositeBounds && drivers.columnExtent) {
      compositeBounds = {
        ...compositeBounds,
        minCol: Math.min(compositeBounds.minCol, drivers.columnExtent.minCol),
        maxCol: Math.max(compositeBounds.maxCol, drivers.columnExtent.maxCol),
      };
    }
    if (!compositeBounds) {
      continue;
    }
    const populatedArea = boundsArea(contentBounds);
    const compositeArea = boundsArea(compositeBounds);
    if (
      contentBounds &&
      contentBounds.minRow === compositeBounds.minRow &&
      contentBounds.maxRow === compositeBounds.maxRow &&
      contentBounds.minCol === compositeBounds.minCol &&
      contentBounds.maxCol === compositeBounds.maxCol
    ) {
      continue;
    }
    const extraRows =
      contentBounds === null
        ? compositeBounds.maxRow - compositeBounds.minRow + 1
        : Math.max(0, contentBounds.minRow - compositeBounds.minRow) +
          Math.max(0, compositeBounds.maxRow - contentBounds.maxRow);
    const extraColumns =
      contentBounds === null
        ? compositeBounds.maxCol - compositeBounds.minCol + 1
        : Math.max(0, contentBounds.minCol - compositeBounds.minCol) +
          Math.max(0, compositeBounds.maxCol - contentBounds.maxCol);
    reports.push({
      sheetName: sheet.name,
      populatedCellCount: sheet.cells.length,
      populatedRange: contentBounds ? boundsToRange(sheet.name, contentBounds) : null,
      compositeRange: boundsToRange(sheet.name, compositeBounds),
      extraRows,
      extraColumns,
      populatedArea,
      compositeArea,
      bloatArea: Math.max(0, compositeArea - populatedArea),
      drivers: drivers.drivers.slice(0, 12),
    });
  }

  const sortedReports = reports.toSorted((left, right) => {
    if (left.bloatArea !== right.bloatArea) {
      return right.bloatArea - left.bloatArea;
    }
    if (left.extraRows !== right.extraRows) {
      return right.extraRows - left.extraRows;
    }
    if (left.extraColumns !== right.extraColumns) {
      return right.extraColumns - left.extraColumns;
    }
    return left.sheetName.localeCompare(right.sheetName);
  });

  return {
    summary: {
      scannedSheetCount:
        input.sheetName === undefined
          ? snapshot.sheets.length
          : snapshot.sheets.filter((sheet) => sheet.name === input.sheetName).length,
      bloatedSheetCount: sortedReports.length,
      totalBloatArea: sortedReports.reduce((sum, report) => sum + report.bloatArea, 0),
      truncated: sortedReports.length > limit,
    },
    sheets: sortedReports.slice(0, limit),
  };
}

export function scanWorkbookPerformanceHotspots(
  runtime: WorkbookRuntime,
  input: {
    sheetName?: string | undefined;
    limit?: number | undefined;
  } = {},
): WorkbookPerformanceHotspotReport {
  const limit = clampLimit(input.limit);
  const structure = summarizeWorkbookStructure(runtime);
  const formulaIssues = findWorkbookFormulaIssues(runtime, {
    ...(input.sheetName !== undefined ? { sheetName: input.sheetName } : {}),
    limit: MAX_AUDIT_LIMIT,
  });
  const jsOnlyBySheet = new Map<string, number>();
  const issueCountBySheet = new Map<string, number>();
  for (const issue of formulaIssues.issues) {
    issueCountBySheet.set(issue.sheetName, (issueCountBySheet.get(issue.sheetName) ?? 0) + 1);
    if (issue.issueKinds.includes("unsupported")) {
      jsOnlyBySheet.set(issue.sheetName, (jsOnlyBySheet.get(issue.sheetName) ?? 0) + 1);
    }
  }

  const hotspots = structure.sheets
    .filter((sheet) => input.sheetName === undefined || sheet.name === input.sheetName)
    .map((sheet) => {
      const jsOnlyFormulaCount = jsOnlyBySheet.get(sheet.name) ?? 0;
      const issueCount = issueCountBySheet.get(sheet.name) ?? 0;
      const reasons: string[] = [];
      if (jsOnlyFormulaCount > 0) {
        reasons.push(
          `${String(jsOnlyFormulaCount)} JS-only formula${jsOnlyFormulaCount === 1 ? "" : "s"}`,
        );
      }
      if (sheet.pivotCount > 0) {
        reasons.push(
          `${String(sheet.pivotCount)} pivot output${sheet.pivotCount === 1 ? "" : "s"}`,
        );
      }
      if (sheet.spillCount > 0) {
        reasons.push(`${String(sheet.spillCount)} spill range${sheet.spillCount === 1 ? "" : "s"}`);
      }
      if (sheet.formulaCellCount > 0) {
        reasons.push(
          `${String(sheet.formulaCellCount)} formula cell${sheet.formulaCellCount === 1 ? "" : "s"}`,
        );
      }
      if (issueCount > 0) {
        reasons.push(`${String(issueCount)} formula issue${issueCount === 1 ? "" : "s"}`);
      }
      return {
        sheetName: sheet.name,
        cellCount: sheet.cellCount,
        formulaCellCount: sheet.formulaCellCount,
        jsOnlyFormulaCount,
        issueCount,
        pivotCount: sheet.pivotCount,
        spillCount: sheet.spillCount,
        usedRange: sheet.usedRange,
        reasons,
      } satisfies WorkbookPerformanceHotspot;
    })
    .filter((sheet) => sheet.reasons.length > 0)
    .toSorted((left, right) => {
      if (left.jsOnlyFormulaCount !== right.jsOnlyFormulaCount) {
        return right.jsOnlyFormulaCount - left.jsOnlyFormulaCount;
      }
      if (left.pivotCount !== right.pivotCount) {
        return right.pivotCount - left.pivotCount;
      }
      if (left.formulaCellCount !== right.formulaCellCount) {
        return right.formulaCellCount - left.formulaCellCount;
      }
      if (left.spillCount !== right.spillCount) {
        return right.spillCount - left.spillCount;
      }
      if (left.issueCount !== right.issueCount) {
        return right.issueCount - left.issueCount;
      }
      if (left.cellCount !== right.cellCount) {
        return right.cellCount - left.cellCount;
      }
      return left.sheetName.localeCompare(right.sheetName);
    });

  return {
    summary: {
      scannedSheetCount:
        input.sheetName === undefined
          ? structure.summary.sheetCount
          : structure.sheets.filter((sheet) => sheet.name === input.sheetName).length,
      hotspotCount: hotspots.length,
      truncated: hotspots.length > limit,
      recalcMetrics: runtime.engine.getLastMetrics(),
    },
    hotspots: hotspots.slice(0, limit),
  };
}

function collectInvariantProblems(snapshot: WorkbookSnapshot): InvariantProblem[] {
  const problems: InvariantProblem[] = [];
  const sheetNames = new Set<string>();
  const sheetOrders = new Set<number>();

  for (const sheet of snapshot.sheets) {
    if (sheetNames.has(sheet.name)) {
      addProblem(problems, "duplicateSheetName", `Duplicate sheet name ${sheet.name}`, {
        sheetName: sheet.name,
      });
    } else {
      sheetNames.add(sheet.name);
    }
    if (sheetOrders.has(sheet.order)) {
      addProblem(problems, "duplicateSheetOrder", `Duplicate sheet order ${String(sheet.order)}`, {
        sheetName: sheet.name,
      });
    } else {
      sheetOrders.add(sheet.order);
    }
    const addresses = new Set<string>();
    for (const cell of sheet.cells) {
      validateAddressRef(problems, sheet.name, cell.address, "Cell");
      if (addresses.has(cell.address)) {
        addProblem(
          problems,
          "duplicateCellAddress",
          `Duplicate cell ${sheet.name}!${cell.address}`,
          {
            sheetName: sheet.name,
            address: cell.address,
          },
        );
      } else {
        addresses.add(cell.address);
      }
    }

    for (const styleRange of sheet.metadata?.styleRanges ?? []) {
      validateRangeRef(problems, sheet.name, styleRange.range, "Style range");
    }
    for (const formatRange of sheet.metadata?.formatRanges ?? []) {
      validateRangeRef(problems, sheet.name, formatRange.range, "Number format range");
    }
    for (const filter of sheet.metadata?.filters ?? []) {
      validateRangeRef(problems, sheet.name, filter, "Filter");
    }
    for (const sort of sheet.metadata?.sorts ?? []) {
      validateRangeRef(problems, sheet.name, sort.range, "Sort");
      for (const key of sort.keys) {
        validateAddressRef(problems, sheet.name, key.keyAddress, "Sort key");
      }
    }
    for (const validation of sheet.metadata?.validations ?? []) {
      validateRangeRef(problems, sheet.name, validation.range, "Validation");
    }
    for (const conditionalFormat of sheet.metadata?.conditionalFormats ?? []) {
      validateRangeRef(problems, sheet.name, conditionalFormat.range, "Conditional format");
    }
    for (const protectedRange of sheet.metadata?.protectedRanges ?? []) {
      validateRangeRef(problems, sheet.name, protectedRange.range, "Protected range");
    }
    for (const commentThread of sheet.metadata?.commentThreads ?? []) {
      if (commentThread.sheetName !== sheet.name) {
        addProblem(
          problems,
          "commentThreadSheetMismatch",
          `Comment thread ${commentThread.threadId} points at ${commentThread.sheetName} instead of ${sheet.name}`,
          { sheetName: sheet.name, address: commentThread.address },
        );
      }
      validateAddressRef(problems, sheet.name, commentThread.address, "Comment thread");
    }
    for (const note of sheet.metadata?.notes ?? []) {
      validateAddressRef(problems, sheet.name, note.address, "Note");
    }
    for (const record of sheet.metadata?.rowMetadata ?? []) {
      if (record.count <= 0 || record.start < 0) {
        addProblem(
          problems,
          "invalidRowMetadata",
          `Invalid row metadata region start=${String(record.start)} count=${String(record.count)}`,
          { sheetName: sheet.name },
        );
      }
    }
    for (const record of sheet.metadata?.columnMetadata ?? []) {
      if (record.count <= 0 || record.start < 0) {
        addProblem(
          problems,
          "invalidColumnMetadata",
          `Invalid column metadata region start=${String(record.start)} count=${String(record.count)}`,
          { sheetName: sheet.name },
        );
      }
    }
  }

  for (const table of snapshot.workbook.metadata?.tables ?? []) {
    if (!ensureSheetExists(sheetNames, problems, table.sheetName, `Table ${table.name}`)) {
      continue;
    }
    validateRangeRef(
      problems,
      table.sheetName,
      {
        sheetName: table.sheetName,
        startAddress: table.startAddress,
        endAddress: table.endAddress,
      },
      `Table ${table.name}`,
    );
  }
  for (const pivot of snapshot.workbook.metadata?.pivots ?? []) {
    const hasPivotSheet = ensureSheetExists(
      sheetNames,
      problems,
      pivot.sheetName,
      `Pivot ${pivot.name}`,
    );
    const hasSourceSheet = ensureSheetExists(
      sheetNames,
      problems,
      pivot.source.sheetName,
      `Pivot ${pivot.name} source`,
    );
    if (hasPivotSheet) {
      validateAddressRef(problems, pivot.sheetName, pivot.address, `Pivot ${pivot.name}`);
    }
    if (hasSourceSheet) {
      validateRangeRef(
        problems,
        pivot.source.sheetName,
        pivot.source,
        `Pivot ${pivot.name} source`,
      );
    }
    if (pivot.rows <= 0 || pivot.cols <= 0) {
      addProblem(
        problems,
        "invalidPivotExtent",
        `Pivot ${pivot.name} has invalid output extent ${String(pivot.rows)}x${String(pivot.cols)}`,
      );
    }
  }
  for (const chart of snapshot.workbook.metadata?.charts ?? []) {
    const hasChartSheet = ensureSheetExists(
      sheetNames,
      problems,
      chart.sheetName,
      `Chart ${chart.id}`,
    );
    const hasSourceSheet = ensureSheetExists(
      sheetNames,
      problems,
      chart.source.sheetName,
      `Chart ${chart.id} source`,
    );
    if (hasChartSheet) {
      validateAddressRef(problems, chart.sheetName, chart.address, `Chart ${chart.id}`);
    }
    if (hasSourceSheet) {
      validateRangeRef(problems, chart.source.sheetName, chart.source, `Chart ${chart.id} source`);
    }
    if (chart.rows <= 0 || chart.cols <= 0) {
      addProblem(
        problems,
        "invalidChartExtent",
        `Chart ${chart.id} has invalid footprint ${String(chart.rows)}x${String(chart.cols)}`,
      );
    }
  }
  for (const image of snapshot.workbook.metadata?.images ?? []) {
    if (ensureSheetExists(sheetNames, problems, image.sheetName, `Image ${image.id}`)) {
      validateAddressRef(problems, image.sheetName, image.address, `Image ${image.id}`);
    }
    if (image.rows <= 0 || image.cols <= 0) {
      addProblem(
        problems,
        "invalidImageExtent",
        `Image ${image.id} has invalid footprint ${String(image.rows)}x${String(image.cols)}`,
        { sheetName: image.sheetName, address: image.address },
      );
    }
  }
  for (const shape of snapshot.workbook.metadata?.shapes ?? []) {
    if (ensureSheetExists(sheetNames, problems, shape.sheetName, `Shape ${shape.id}`)) {
      validateAddressRef(problems, shape.sheetName, shape.address, `Shape ${shape.id}`);
    }
    if (shape.rows <= 0 || shape.cols <= 0) {
      addProblem(
        problems,
        "invalidShapeExtent",
        `Shape ${shape.id} has invalid footprint ${String(shape.rows)}x${String(shape.cols)}`,
        { sheetName: shape.sheetName, address: shape.address },
      );
    }
  }
  for (const spill of snapshot.workbook.metadata?.spills ?? []) {
    if (ensureSheetExists(sheetNames, problems, spill.sheetName, "Spill range")) {
      validateAddressRef(problems, spill.sheetName, spill.address, "Spill range");
    }
    if (spill.rows <= 0 || spill.cols <= 0) {
      addProblem(
        problems,
        "invalidSpillExtent",
        `Spill ${spill.sheetName}!${spill.address} has invalid extent ${String(spill.rows)}x${String(spill.cols)}`,
        { sheetName: spill.sheetName, address: spill.address },
      );
    }
  }
  for (const definedName of snapshot.workbook.metadata?.definedNames ?? []) {
    const value = definedName.value;
    if (typeof value !== "object" || value === null) {
      continue;
    }
    if (value.kind === "cell-ref") {
      if (
        ensureSheetExists(sheetNames, problems, value.sheetName, `Defined name ${definedName.name}`)
      ) {
        validateAddressRef(
          problems,
          value.sheetName,
          value.address,
          `Defined name ${definedName.name}`,
        );
      }
    }
    if (value.kind === "range-ref") {
      if (
        ensureSheetExists(sheetNames, problems, value.sheetName, `Defined name ${definedName.name}`)
      ) {
        validateRangeRef(
          problems,
          value.sheetName,
          {
            sheetName: value.sheetName,
            startAddress: value.startAddress,
            endAddress: value.endAddress,
          },
          `Defined name ${definedName.name}`,
        );
      }
    }
  }

  return problems;
}

export async function verifyWorkbookInvariants(
  runtime: WorkbookRuntime,
  input: {
    roundTrip?: boolean | undefined;
  } = {},
): Promise<WorkbookInvariantVerificationReport> {
  const snapshot = runtime.engine.exportSnapshot();
  const problems = collectInvariantProblems(snapshot);
  const roundTripChecked = input.roundTrip !== false;
  let roundTripStable = true;
  if (roundTripChecked) {
    const restored = new SpreadsheetEngine({
      workbookName: `${runtime.documentId}:audit-roundtrip`,
      replicaId: `audit:${runtime.documentId}`,
    });
    await restored.ready();
    restored.importSnapshot(snapshot);
    const exported = restored.exportSnapshot();
    if (!isDeepStrictEqual(exported, snapshot)) {
      roundTripStable = false;
      addProblem(
        problems,
        "roundTripMismatch",
        "Snapshot export/import round-trip changed workbook state",
      );
    }
  }
  return {
    summary: {
      ok: problems.length === 0,
      problemCount: problems.length,
      roundTripChecked,
      roundTripStable,
    },
    problems,
  };
}
