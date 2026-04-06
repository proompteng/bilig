import { FormulaMode, ValueTag, formatErrorCode } from "@bilig/protocol";
import type { WorkbookRuntime } from "../workbook-runtime/runtime-manager.js";

const DEFAULT_FORMULA_ISSUE_LIMIT = 50;
const MAX_FORMULA_ISSUE_LIMIT = 200;
const DEFAULT_SEARCH_LIMIT = 20;
const MAX_SEARCH_LIMIT = 50;
const DEFAULT_TRACE_DEPTH = 2;
const MAX_TRACE_DEPTH = 4;
const MAX_TRACE_NODES = 96;

export interface WorkbookFormulaIssue {
  readonly sheetName: string;
  readonly address: string;
  readonly formula: string;
  readonly valueText: string;
  readonly issueKinds: readonly ("error" | "cycle" | "unsupported")[];
  readonly errorText: string | null;
  readonly mode: "literal" | "wasm" | "js";
  readonly directPrecedentCount: number;
  readonly directDependentCount: number;
}

export interface WorkbookFormulaIssueReport {
  readonly summary: {
    readonly scannedFormulaCells: number;
    readonly issueCount: number;
    readonly errorCount: number;
    readonly cycleCount: number;
    readonly unsupportedCount: number;
    readonly truncated: boolean;
  };
  readonly issues: readonly WorkbookFormulaIssue[];
}

export interface WorkbookSearchMatch {
  readonly kind: "sheet" | "cell";
  readonly score: number;
  readonly reasons: readonly string[];
  readonly sheetName: string;
  readonly address: string | null;
  readonly formula: string | null;
  readonly inputText: string | null;
  readonly valueText: string | null;
  readonly snippet: string;
}

export interface WorkbookSearchReport {
  readonly query: string;
  readonly summary: {
    readonly matchCount: number;
    readonly truncated: boolean;
  };
  readonly matches: readonly WorkbookSearchMatch[];
}

export interface WorkbookDependencyTraceNode {
  readonly sheetName: string;
  readonly address: string;
  readonly formula: string | null;
  readonly valueText: string;
  readonly mode: "literal" | "wasm" | "js";
  readonly inCycle: boolean;
}

export interface WorkbookDependencyTraceLayer {
  readonly depth: number;
  readonly precedents: readonly WorkbookDependencyTraceNode[];
  readonly dependents: readonly WorkbookDependencyTraceNode[];
}

export interface WorkbookDependencyTraceReport {
  readonly root: WorkbookDependencyTraceNode;
  readonly direction: "precedents" | "dependents" | "both";
  readonly depth: number;
  readonly summary: {
    readonly precedentCount: number;
    readonly dependentCount: number;
    readonly truncated: boolean;
  };
  readonly layers: readonly WorkbookDependencyTraceLayer[];
}

interface FormulaIssueCache {
  readonly headRevision: number;
  readonly calculatedRevision: number;
  readonly report: WorkbookFormulaIssueReport;
}

interface SearchIndexEntry {
  readonly kind: "sheet" | "cell";
  readonly sheetName: string;
  readonly address: string | null;
  readonly formula: string | null;
  readonly inputText: string | null;
  readonly valueText: string | null;
  readonly snippet: string;
  readonly searchText: string;
}

interface SearchIndexCache {
  readonly headRevision: number;
  readonly calculatedRevision: number;
  readonly entries: readonly SearchIndexEntry[];
}

const formulaIssueCache = new WeakMap<object, FormulaIssueCache>();
const searchIndexCache = new WeakMap<object, SearchIndexCache>();

function stringifyLiteralValue(value: unknown): string | null {
  if (value === null || value === undefined) {
    return null;
  }
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  return JSON.stringify(value);
}

function valueToText(value: {
  readonly tag: ValueTag;
  readonly value?: string | number | boolean;
  readonly code?: number;
}): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return "";
    case ValueTag.Number:
      return value.value === undefined ? "" : String(value.value);
    case ValueTag.Boolean:
      return value.value ? "TRUE" : "FALSE";
    case ValueTag.String:
      return typeof value.value === "string" ? value.value : "";
    case ValueTag.Error:
      return typeof value.code === "number" ? formatErrorCode(value.code) : "#ERROR!";
    default:
      return "";
  }
}

function modeToLabel(mode: FormulaMode | undefined): "literal" | "wasm" | "js" {
  if (mode === FormulaMode.WasmFastPath) {
    return "wasm";
  }
  if (mode === FormulaMode.JsOnly) {
    return "js";
  }
  return "literal";
}

function clampLimit(limit: number | undefined, fallback: number, max: number): number {
  if (!Number.isFinite(limit) || typeof limit !== "number") {
    return fallback;
  }
  return Math.max(1, Math.min(max, Math.trunc(limit)));
}

function normalizeQuery(query: string): {
  readonly queryText: string;
  readonly queryLower: string;
  readonly tokens: readonly string[];
} {
  const queryText = query.trim();
  if (queryText.length === 0) {
    throw new Error("query must not be empty");
  }
  const queryLower = queryText.toLowerCase();
  const tokens = queryLower.split(/[^a-z0-9]+/).filter((entry) => entry.length > 0);
  return {
    queryText,
    queryLower,
    tokens: tokens.length > 0 ? tokens : [queryLower],
  };
}

function splitQualifiedAddress(qualifiedAddress: string): {
  readonly sheetName: string;
  readonly address: string;
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

function describeTraceNode(
  runtime: WorkbookRuntime,
  sheetName: string,
  address: string,
): WorkbookDependencyTraceNode {
  const explanation = runtime.engine.explainCell(sheetName, address);
  return {
    sheetName,
    address,
    formula: explanation.formula ? `=${explanation.formula}` : null,
    valueText: valueToText(explanation.value),
    mode: modeToLabel(explanation.mode),
    inCycle: explanation.inCycle,
  };
}

function getFormulaIssueReport(runtime: WorkbookRuntime): WorkbookFormulaIssueReport {
  const cached = formulaIssueCache.get(runtime.engine);
  if (
    cached &&
    cached.headRevision === runtime.headRevision &&
    cached.calculatedRevision === runtime.calculatedRevision
  ) {
    return cached.report;
  }

  const snapshot = runtime.engine.exportSnapshot();
  const issues: WorkbookFormulaIssue[] = [];
  let scannedFormulaCells = 0;
  let errorCount = 0;
  let cycleCount = 0;
  let unsupportedCount = 0;

  for (const sheet of snapshot.sheets) {
    for (const cell of sheet.cells) {
      if (!cell.formula) {
        continue;
      }
      scannedFormulaCells += 1;
      const explanation = runtime.engine.explainCell(sheet.name, cell.address);
      const issueKinds: Array<"error" | "cycle" | "unsupported"> = [];
      if (explanation.value.tag === ValueTag.Error) {
        issueKinds.push("error");
        errorCount += 1;
      }
      if (explanation.inCycle) {
        issueKinds.push("cycle");
        cycleCount += 1;
      }
      if (explanation.mode === FormulaMode.JsOnly) {
        issueKinds.push("unsupported");
        unsupportedCount += 1;
      }
      if (issueKinds.length === 0) {
        continue;
      }
      issues.push({
        sheetName: sheet.name,
        address: cell.address,
        formula: `=${cell.formula}`,
        valueText: valueToText(explanation.value),
        issueKinds,
        errorText:
          explanation.value.tag === ValueTag.Error && typeof explanation.value.code === "number"
            ? formatErrorCode(explanation.value.code)
            : null,
        mode: modeToLabel(explanation.mode),
        directPrecedentCount: explanation.directPrecedents.length,
        directDependentCount: explanation.directDependents.length,
      });
    }
  }

  issues.sort((left, right) => {
    const leftSeverity =
      (left.issueKinds.includes("cycle") ? 4 : 0) +
      (left.issueKinds.includes("error") ? 2 : 0) +
      (left.issueKinds.includes("unsupported") ? 1 : 0);
    const rightSeverity =
      (right.issueKinds.includes("cycle") ? 4 : 0) +
      (right.issueKinds.includes("error") ? 2 : 0) +
      (right.issueKinds.includes("unsupported") ? 1 : 0);
    if (leftSeverity !== rightSeverity) {
      return rightSeverity - leftSeverity;
    }
    if (left.sheetName !== right.sheetName) {
      return left.sheetName.localeCompare(right.sheetName);
    }
    return left.address.localeCompare(right.address, undefined, { numeric: true });
  });

  const report: WorkbookFormulaIssueReport = {
    summary: {
      scannedFormulaCells,
      issueCount: issues.length,
      errorCount,
      cycleCount,
      unsupportedCount,
      truncated: false,
    },
    issues,
  };
  formulaIssueCache.set(runtime.engine, {
    headRevision: runtime.headRevision,
    calculatedRevision: runtime.calculatedRevision,
    report,
  });
  return report;
}

function getSearchIndex(runtime: WorkbookRuntime): readonly SearchIndexEntry[] {
  const cached = searchIndexCache.get(runtime.engine);
  if (
    cached &&
    cached.headRevision === runtime.headRevision &&
    cached.calculatedRevision === runtime.calculatedRevision
  ) {
    return cached.entries;
  }

  const snapshot = runtime.engine.exportSnapshot();
  const entries: SearchIndexEntry[] = [];
  for (const sheet of snapshot.sheets) {
    entries.push({
      kind: "sheet",
      sheetName: sheet.name,
      address: null,
      formula: null,
      inputText: null,
      valueText: null,
      snippet: `Sheet ${sheet.name}`,
      searchText: `sheet ${sheet.name}`.toLowerCase(),
    });
    for (const cell of sheet.cells) {
      const liveCell = runtime.engine.getCell(sheet.name, cell.address);
      const formula = cell.formula ? `=${cell.formula}` : null;
      const inputText = stringifyLiteralValue(cell.value);
      const valueText = valueToText(liveCell.value);
      const snippet = (formula ?? inputText ?? valueText) || `${sheet.name}!${cell.address}`;
      entries.push({
        kind: "cell",
        sheetName: sheet.name,
        address: cell.address,
        formula,
        inputText,
        valueText: valueText.length > 0 ? valueText : null,
        snippet,
        searchText: [
          sheet.name,
          cell.address,
          `${sheet.name}!${cell.address}`,
          formula ?? "",
          inputText ?? "",
          valueText,
        ]
          .join("\n")
          .toLowerCase(),
      });
    }
  }

  searchIndexCache.set(runtime.engine, {
    headRevision: runtime.headRevision,
    calculatedRevision: runtime.calculatedRevision,
    entries,
  });
  return entries;
}

function scoreSearchEntry(
  entry: SearchIndexEntry,
  queryLower: string,
  tokens: readonly string[],
): { readonly score: number; readonly reasons: readonly string[] } | null {
  const reasons: string[] = [];
  let score = 0;
  const qualifiedAddress = entry.address ? `${entry.sheetName}!${entry.address}` : null;
  const exactAddressMatch =
    qualifiedAddress !== null &&
    (qualifiedAddress.toLowerCase() === queryLower || entry.address?.toLowerCase() === queryLower);
  const exactSheetMatch = entry.sheetName.toLowerCase() === queryLower;
  if (exactAddressMatch) {
    score += 120;
    reasons.push("address");
  } else if (exactSheetMatch) {
    score += 100;
    reasons.push("sheet");
  }

  const matchedAllTokens = tokens.every((token) => entry.searchText.includes(token));
  if (!matchedAllTokens && score === 0) {
    return null;
  }
  if (matchedAllTokens) {
    score += tokens.length * 12;
  }
  if (entry.formula?.toLowerCase().includes(queryLower)) {
    score += 50;
    reasons.push("formula");
  }
  if (entry.inputText?.toLowerCase().includes(queryLower)) {
    score += 35;
    reasons.push("input");
  }
  if (entry.valueText?.toLowerCase().includes(queryLower)) {
    score += 30;
    reasons.push("value");
  }
  if (entry.kind === "sheet" && entry.sheetName.toLowerCase().includes(queryLower)) {
    score += 40;
    if (!reasons.includes("sheet")) {
      reasons.push("sheet");
    }
  }
  if (entry.kind === "cell" && qualifiedAddress?.toLowerCase().includes(queryLower)) {
    score += 25;
    if (!reasons.includes("address")) {
      reasons.push("address");
    }
  }
  if (score === 0) {
    return null;
  }
  return {
    score,
    reasons,
  };
}

export function findWorkbookFormulaIssues(
  runtime: WorkbookRuntime,
  input: {
    readonly sheetName?: string;
    readonly limit?: number;
  } = {},
): WorkbookFormulaIssueReport {
  const limit = clampLimit(input.limit, DEFAULT_FORMULA_ISSUE_LIMIT, MAX_FORMULA_ISSUE_LIMIT);
  const cached = getFormulaIssueReport(runtime);
  const filteredIssues =
    typeof input.sheetName === "string"
      ? cached.issues.filter((issue) => issue.sheetName === input.sheetName)
      : cached.issues;
  return {
    summary: {
      ...cached.summary,
      issueCount: filteredIssues.length,
      errorCount: filteredIssues.filter((issue) => issue.issueKinds.includes("error")).length,
      cycleCount: filteredIssues.filter((issue) => issue.issueKinds.includes("cycle")).length,
      unsupportedCount: filteredIssues.filter((issue) => issue.issueKinds.includes("unsupported"))
        .length,
      truncated: filteredIssues.length > limit,
    },
    issues: filteredIssues.slice(0, limit),
  };
}

export function searchWorkbook(
  runtime: WorkbookRuntime,
  input: {
    readonly query: string;
    readonly sheetName?: string;
    readonly limit?: number;
  },
): WorkbookSearchReport {
  const { queryText, queryLower, tokens } = normalizeQuery(input.query);
  const limit = clampLimit(input.limit, DEFAULT_SEARCH_LIMIT, MAX_SEARCH_LIMIT);
  const matches = getSearchIndex(runtime)
    .filter((entry) => input.sheetName === undefined || entry.sheetName === input.sheetName)
    .flatMap((entry) => {
      const match = scoreSearchEntry(entry, queryLower, tokens);
      if (!match) {
        return [];
      }
      return [
        {
          kind: entry.kind,
          score: match.score,
          reasons: match.reasons,
          sheetName: entry.sheetName,
          address: entry.address,
          formula: entry.formula,
          inputText: entry.inputText,
          valueText: entry.valueText,
          snippet: entry.snippet,
        } satisfies WorkbookSearchMatch,
      ];
    })
    .toSorted((left, right) => {
      if (left.score !== right.score) {
        return right.score - left.score;
      }
      if (left.kind !== right.kind) {
        return left.kind === "sheet" ? -1 : 1;
      }
      if (left.sheetName !== right.sheetName) {
        return left.sheetName.localeCompare(right.sheetName);
      }
      return (left.address ?? "").localeCompare(right.address ?? "", undefined, { numeric: true });
    });
  return {
    query: queryText,
    summary: {
      matchCount: matches.length,
      truncated: matches.length > limit,
    },
    matches: matches.slice(0, limit),
  };
}

export function traceWorkbookDependencies(
  runtime: WorkbookRuntime,
  input: {
    readonly sheetName: string;
    readonly address: string;
    readonly direction?: "precedents" | "dependents" | "both";
    readonly depth?: number;
  },
): WorkbookDependencyTraceReport {
  const direction = input.direction ?? "both";
  const depth = clampLimit(input.depth, DEFAULT_TRACE_DEPTH, MAX_TRACE_DEPTH);
  const rootId = `${input.sheetName}!${input.address}`;
  const visited = new Set<string>([rootId]);
  let remainingBudget = MAX_TRACE_NODES - 1;
  let frontierPrecedents = direction === "dependents" ? [] : [rootId];
  let frontierDependents = direction === "precedents" ? [] : [rootId];
  let precedentCount = 0;
  let dependentCount = 0;
  let truncated = false;
  const layers: WorkbookDependencyTraceLayer[] = [];

  for (let layerDepth = 1; layerDepth <= depth; layerDepth += 1) {
    const layerPrecedents: WorkbookDependencyTraceNode[] = [];
    const layerDependents: WorkbookDependencyTraceNode[] = [];
    const nextPrecedents: string[] = [];
    const nextDependents: string[] = [];

    for (const current of frontierPrecedents) {
      const { sheetName, address } = splitQualifiedAddress(current);
      const dependencies = runtime.engine.getDependencies(sheetName, address);
      for (const candidate of dependencies.directPrecedents) {
        if (visited.has(candidate)) {
          continue;
        }
        if (remainingBudget <= 0) {
          truncated = true;
          continue;
        }
        visited.add(candidate);
        remainingBudget -= 1;
        precedentCount += 1;
        const parsed = splitQualifiedAddress(candidate);
        layerPrecedents.push(describeTraceNode(runtime, parsed.sheetName, parsed.address));
        nextPrecedents.push(candidate);
      }
    }

    for (const current of frontierDependents) {
      const { sheetName, address } = splitQualifiedAddress(current);
      const dependencies = runtime.engine.getDependencies(sheetName, address);
      for (const candidate of dependencies.directDependents) {
        if (visited.has(candidate)) {
          continue;
        }
        if (remainingBudget <= 0) {
          truncated = true;
          continue;
        }
        visited.add(candidate);
        remainingBudget -= 1;
        dependentCount += 1;
        const parsed = splitQualifiedAddress(candidate);
        layerDependents.push(describeTraceNode(runtime, parsed.sheetName, parsed.address));
        nextDependents.push(candidate);
      }
    }

    if (layerPrecedents.length === 0 && layerDependents.length === 0) {
      break;
    }

    layers.push({
      depth: layerDepth,
      precedents: layerPrecedents,
      dependents: layerDependents,
    });
    frontierPrecedents = nextPrecedents;
    frontierDependents = nextDependents;
  }

  return {
    root: describeTraceNode(runtime, input.sheetName, input.address),
    direction,
    depth,
    summary: {
      precedentCount,
      dependentCount,
      truncated,
    },
    layers,
  };
}
