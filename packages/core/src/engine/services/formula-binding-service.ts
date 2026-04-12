import { Effect } from "effect";
import {
  compileFormula,
  compileFormulaAst,
  type FormulaNode,
  parseCellAddress,
  parseRangeAddress,
  renameFormulaSheetReferences,
} from "@bilig/formula";
import { FormulaMode, ErrorCode, Opcode } from "@bilig/protocol";
import { CellFlags } from "../../cell-store.js";
import type { EdgeArena, EdgeSlice } from "../../edge-arena.js";
import { resolveRuntimeDirectLookupBinding } from "../direct-vector-lookup.js";
import {
  entityPayload,
  isExactLookupColumnEntity,
  isRangeEntity,
  isSortedLookupColumnEntity,
  makeCellEntity,
  makeExactLookupColumnEntity,
  makeRangeEntity,
  makeSortedLookupColumnEntity,
} from "../../entity-ids.js";
import { growUint32 } from "../../engine-buffer-utils.js";
import {
  resolveMetadataReferencesInAst,
  spillDependencyKeyFromRef,
  tableDependencyKey,
} from "../../engine-metadata-utils.js";
import { errorValue } from "../../engine-value-utils.js";
import { normalizeDefinedName } from "../../workbook-store.js";
import {
  type EngineRuntimeState,
  type MaterializedDependencies,
  type RuntimeDirectLookupDescriptor,
  type RuntimeFormula,
  UNRESOLVED_WASM_OPERAND,
  type U32,
} from "../runtime-state.js";
import { EngineFormulaBindingError } from "../errors.js";
import type { Uint32Arena, Float64Arena } from "@bilig/formula";
import type { EngineCompiledPlanService } from "./compiled-plan-service.js";
import type { ExactColumnIndexService } from "./exact-column-index-service.js";
import type { SortedColumnSearchService } from "./sorted-column-search-service.js";

export interface EngineFormulaBindingService {
  readonly bindFormula: (
    cellIndex: number,
    ownerSheetName: string,
    source: string,
  ) => Effect.Effect<boolean, EngineFormulaBindingError>;
  readonly clearFormula: (cellIndex: number) => Effect.Effect<boolean, EngineFormulaBindingError>;
  readonly invalidateFormula: (cellIndex: number) => Effect.Effect<void, EngineFormulaBindingError>;
  readonly rewriteCellFormulasForSheetRename: (
    oldSheetName: string,
    newSheetName: string,
    formulaChangedCount: number,
  ) => Effect.Effect<number, EngineFormulaBindingError>;
  readonly rebuildAllFormulaBindings: () => Effect.Effect<number[], EngineFormulaBindingError>;
  readonly rebindFormulaCells: (
    candidates: readonly number[],
    formulaChangedCount: number,
  ) => Effect.Effect<number, EngineFormulaBindingError>;
  readonly rebindDefinedNameDependents: (
    names: readonly string[],
    formulaChangedCount: number,
  ) => Effect.Effect<number, EngineFormulaBindingError>;
  readonly rebindTableDependents: (
    tableNames: readonly string[],
    formulaChangedCount: number,
  ) => Effect.Effect<number, EngineFormulaBindingError>;
  readonly rebindFormulasForSheet: (
    sheetName: string,
    formulaChangedCount: number,
    candidates?: readonly number[] | U32,
  ) => Effect.Effect<number, EngineFormulaBindingError>;
  readonly bindFormulaNow: (cellIndex: number, ownerSheetName: string, source: string) => boolean;
  readonly clearFormulaNow: (cellIndex: number) => boolean;
  readonly invalidateFormulaNow: (cellIndex: number) => void;
  readonly rebindFormulaCellsNow: (
    candidates: readonly number[],
    formulaChangedCount: number,
  ) => number;
  readonly rebindDefinedNameDependentsNow: (
    names: readonly string[],
    formulaChangedCount: number,
  ) => number;
  readonly rebindTableDependentsNow: (
    tableNames: readonly string[],
    formulaChangedCount: number,
  ) => number;
  readonly rebindFormulasForSheetNow: (
    sheetName: string,
    formulaChangedCount: number,
    candidates?: readonly number[] | U32,
  ) => number;
}

function formulaBindingErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message;
}

function appendTrackedReverseEdge(
  registry: Map<string, Set<number>>,
  key: string,
  dependentCellIndex: number,
): void {
  const existing = registry.get(key);
  if (existing) {
    existing.add(dependentCellIndex);
    return;
  }
  registry.set(key, new Set([dependentCellIndex]));
}

function removeTrackedReverseEdge(
  registry: Map<string, Set<number>>,
  key: string,
  dependentCellIndex: number,
): void {
  const existing = registry.get(key);
  if (!existing) {
    return;
  }
  existing.delete(dependentCellIndex);
  if (existing.size === 0) {
    registry.delete(key);
  }
}

function collectTrackedDependents(
  registry: Map<string, Set<number>>,
  keys: readonly string[],
): number[] {
  const candidates = new Set<number>();
  keys.forEach((key) => {
    registry.get(key)?.forEach((cellIndex) => {
      candidates.add(cellIndex);
    });
  });
  return [...candidates];
}

function directLookupColumnInfo(directLookup: RuntimeDirectLookupDescriptor): {
  sheetName: string;
  col: number;
  isExact: boolean;
} {
  switch (directLookup.kind) {
    case "exact":
      return {
        sheetName: directLookup.prepared.sheetName,
        col: directLookup.prepared.col,
        isExact: true,
      };
    case "exact-uniform-numeric":
      return {
        sheetName: directLookup.sheetName,
        col: directLookup.col,
        isExact: true,
      };
    case "approximate":
      return {
        sheetName: directLookup.prepared.sheetName,
        col: directLookup.prepared.col,
        isExact: false,
      };
    case "approximate-uniform-numeric":
      return {
        sheetName: directLookup.sheetName,
        col: directLookup.col,
        isExact: false,
      };
  }
}

function staticIntegerValue(node: FormulaNode | undefined): number | undefined {
  if (!node) {
    return undefined;
  }
  if (node.kind === "NumberLiteral") {
    return Number.isInteger(node.value) ? node.value : undefined;
  }
  if (
    node.kind === "UnaryExpr" &&
    node.operator === "-" &&
    node.argument.kind === "NumberLiteral" &&
    Number.isInteger(node.argument.value)
  ) {
    return -node.argument.value;
  }
  return undefined;
}

function hasIndexedExactLookupCandidate(node: FormulaNode): boolean {
  return collectIndexedExactLookupCandidates(node).length > 0;
}

function hasDirectApproximateLookupCandidate(node: FormulaNode): boolean {
  return collectDirectApproximateLookupCandidates(node).length > 0;
}

interface IndexedExactLookupCandidate {
  sheetName?: string;
  start: string;
  end: string;
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
}

function collectIndexedExactLookupCandidates(node: FormulaNode): IndexedExactLookupCandidate[] {
  switch (node.kind) {
    case "CallExpr": {
      const callee = node.callee.trim().toUpperCase();
      const lookupRange = node.args[1];
      if (
        lookupRange?.kind === "RangeRef" &&
        lookupRange.refKind === "cells" &&
        lookupRange.start !== lookupRange.end
      ) {
        const isIndexedLookupCall =
          (callee === "MATCH" &&
            node.args.length === 3 &&
            staticIntegerValue(node.args[2]) === 0) ||
          (callee === "XMATCH" &&
            node.args.length >= 2 &&
            node.args.length <= 4 &&
            (node.args.length === 2 || staticIntegerValue(node.args[2]) === 0) &&
            (node.args.length < 4 ||
              staticIntegerValue(node.args[3]) === 1 ||
              staticIntegerValue(node.args[3]) === -1));
        if (isIndexedLookupCall) {
          const parsedRange = parseRangeAddress(
            `${lookupRange.start}:${lookupRange.end}`,
            lookupRange.sheetName,
          );
          if (parsedRange.kind === "cells") {
            return [
              {
                ...(lookupRange.sheetName === undefined
                  ? {}
                  : { sheetName: lookupRange.sheetName }),
                start: lookupRange.start,
                end: lookupRange.end,
                startRow: parsedRange.start.row,
                endRow: parsedRange.end.row,
                startCol: parsedRange.start.col,
                endCol: parsedRange.end.col,
              },
              ...node.args.flatMap(collectIndexedExactLookupCandidates),
            ];
          }
        }
      }
      return node.args.flatMap(collectIndexedExactLookupCandidates);
    }
    case "UnaryExpr":
      return collectIndexedExactLookupCandidates(node.argument);
    case "BinaryExpr":
      return [
        ...collectIndexedExactLookupCandidates(node.left),
        ...collectIndexedExactLookupCandidates(node.right),
      ];
    case "InvokeExpr":
      return [
        ...collectIndexedExactLookupCandidates(node.callee),
        ...node.args.flatMap(collectIndexedExactLookupCandidates),
      ];
    case "BooleanLiteral":
    case "CellRef":
    case "ColumnRef":
    case "ErrorLiteral":
    case "NameRef":
    case "NumberLiteral":
    case "RangeRef":
    case "RowRef":
    case "SpillRef":
    case "StringLiteral":
    case "StructuredRef":
      return [];
  }
}

interface DirectApproximateLookupCandidate {
  sheetName?: string;
  start: string;
  end: string;
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
}

type ParsedCompiledFormula = ReturnType<typeof compileFormula> & {
  parsedDeps?: Array<{ address: string; kind: "cell"; sheetName?: string }>;
  parsedSymbolicRefs?: Array<{ address: string; sheetName?: string }>;
};

function uint32ArrayEqual(
  left: Uint32Array | readonly number[],
  right: Uint32Array | readonly number[],
): boolean {
  if (left.length !== right.length) {
    return false;
  }
  for (let index = 0; index < left.length; index += 1) {
    if (left[index] !== right[index]) {
      return false;
    }
  }
  return true;
}

function stringArrayEqual(left: readonly string[], right: readonly string[]): boolean {
  if (left.length !== right.length) {
    return false;
  }
  for (let index = 0; index < left.length; index += 1) {
    if (left[index] !== right[index]) {
      return false;
    }
  }
  return true;
}

function collectDirectApproximateLookupCandidates(
  node: FormulaNode,
): DirectApproximateLookupCandidate[] {
  switch (node.kind) {
    case "CallExpr": {
      const callee = node.callee.trim().toUpperCase();
      const lookupRange = node.args[1];
      if (
        lookupRange?.kind === "RangeRef" &&
        lookupRange.refKind === "cells" &&
        lookupRange.start !== lookupRange.end
      ) {
        const matchMode = staticIntegerValue(node.args[2]);
        const searchMode = node.args.length >= 4 ? staticIntegerValue(node.args[3]) : 1;
        const isDirectApproximateLookupCall =
          (callee === "MATCH" && node.args.length === 3 && (matchMode === 1 || matchMode === -1)) ||
          (callee === "XMATCH" &&
            node.args.length >= 3 &&
            node.args.length <= 4 &&
            (matchMode === 1 || matchMode === -1) &&
            searchMode === 1);
        if (isDirectApproximateLookupCall) {
          const parsedRange = parseRangeAddress(
            `${lookupRange.start}:${lookupRange.end}`,
            lookupRange.sheetName,
          );
          if (parsedRange.kind === "cells") {
            return [
              {
                ...(lookupRange.sheetName === undefined
                  ? {}
                  : { sheetName: lookupRange.sheetName }),
                start: lookupRange.start,
                end: lookupRange.end,
                startRow: parsedRange.start.row,
                endRow: parsedRange.end.row,
                startCol: parsedRange.start.col,
                endCol: parsedRange.end.col,
              },
              ...node.args.flatMap(collectDirectApproximateLookupCandidates),
            ];
          }
        }
      }
      return node.args.flatMap(collectDirectApproximateLookupCandidates);
    }
    case "UnaryExpr":
      return collectDirectApproximateLookupCandidates(node.argument);
    case "BinaryExpr":
      return [
        ...collectDirectApproximateLookupCandidates(node.left),
        ...collectDirectApproximateLookupCandidates(node.right),
      ];
    case "InvokeExpr":
      return [
        ...collectDirectApproximateLookupCandidates(node.callee),
        ...node.args.flatMap(collectDirectApproximateLookupCandidates),
      ];
    case "BooleanLiteral":
    case "CellRef":
    case "ColumnRef":
    case "ErrorLiteral":
    case "NameRef":
    case "NumberLiteral":
    case "RangeRef":
    case "RowRef":
    case "SpillRef":
    case "StringLiteral":
    case "StructuredRef":
      return [];
  }
}

function buildDirectLookupDescriptor(args: {
  readonly compiled: ParsedCompiledFormula;
  readonly ownerSheetName: string;
  readonly workbook: Pick<EngineRuntimeState, "workbook">["workbook"];
  readonly ensureCellTracked: (sheetName: string, address: string) => number;
  readonly exactLookup: Pick<ExactColumnIndexService, "prepareVectorLookup">;
  readonly sortedLookup: Pick<SortedColumnSearchService, "prepareVectorLookup">;
}): RuntimeDirectLookupDescriptor | undefined {
  const binding = resolveRuntimeDirectLookupBinding(args.compiled.jsPlan, args.ownerSheetName);
  if (!binding) {
    return undefined;
  }
  if (
    !args.workbook.getSheet(binding.lookupSheetName) ||
    !args.workbook.getSheet(binding.operandSheetName)
  ) {
    return undefined;
  }
  const operandCellIndex = args.ensureCellTracked(binding.operandSheetName, binding.operandAddress);
  if (binding.kind === "exact") {
    const prepared = args.exactLookup.prepareVectorLookup({
      sheetName: binding.lookupSheetName,
      rowStart: binding.rowStart,
      rowEnd: binding.rowEnd,
      col: binding.col,
    });
    if (
      prepared.comparableKind === "numeric" &&
      prepared.uniformStart !== undefined &&
      prepared.uniformStep !== undefined
    ) {
      return {
        kind: "exact-uniform-numeric",
        operandCellIndex,
        sheetName: binding.lookupSheetName,
        rowStart: binding.rowStart,
        rowEnd: binding.rowEnd,
        col: binding.col,
        length: prepared.length,
        columnVersion: prepared.columnVersion,
        sheetColumnVersions: prepared.sheetColumnVersions,
        start: prepared.uniformStart,
        step: prepared.uniformStep,
        searchMode: binding.searchMode,
      };
    }
    return {
      kind: "exact",
      operandCellIndex,
      prepared,
      searchMode: binding.searchMode,
    };
  }
  const prepared = args.sortedLookup.prepareVectorLookup({
    sheetName: binding.lookupSheetName,
    rowStart: binding.rowStart,
    rowEnd: binding.rowEnd,
    col: binding.col,
  });
  if (
    prepared.comparableKind === "numeric" &&
    prepared.uniformStart !== undefined &&
    prepared.uniformStep !== undefined
  ) {
    return {
      kind: "approximate-uniform-numeric",
      operandCellIndex,
      sheetName: binding.lookupSheetName,
      rowStart: binding.rowStart,
      rowEnd: binding.rowEnd,
      col: binding.col,
      length: prepared.length,
      columnVersion: prepared.columnVersion,
      sheetColumnVersions: prepared.sheetColumnVersions,
      start: prepared.uniformStart,
      step: prepared.uniformStep,
      matchMode: binding.matchMode,
    };
  }
  return {
    kind: "approximate",
    operandCellIndex,
    prepared,
    matchMode: binding.matchMode,
  };
}

const PUSH_CELL_OPCODE = Number(Opcode.PushCell);
const PUSH_RANGE_OPCODE = Number(Opcode.PushRange);
const PUSH_STRING_OPCODE = Number(Opcode.PushString);

export function createEngineFormulaBindingService(args: {
  readonly state: Pick<
    EngineRuntimeState,
    "workbook" | "strings" | "formulas" | "ranges" | "useColumnIndex"
  >;
  readonly compiledPlans: EngineCompiledPlanService;
  readonly exactLookup: Pick<ExactColumnIndexService, "primeColumnIndex" | "prepareVectorLookup">;
  readonly sortedLookup: Pick<
    SortedColumnSearchService,
    "primeColumnIndex" | "prepareVectorLookup"
  >;
  readonly edgeArena: EdgeArena;
  readonly programArena: Uint32Arena;
  readonly constantArena: Float64Arena;
  readonly rangeListArena: Uint32Arena;
  readonly reverseState: {
    reverseCellEdges: Array<EdgeSlice | undefined>;
    reverseRangeEdges: Array<EdgeSlice | undefined>;
    reverseDefinedNameEdges: Map<string, Set<number>>;
    reverseTableEdges: Map<string, Set<number>>;
    reverseSpillEdges: Map<string, Set<number>>;
    reverseExactLookupColumnEdges: Map<number, EdgeSlice>;
    reverseSortedLookupColumnEdges: Map<number, EdgeSlice>;
  };
  readonly ensureCellTracked: (sheetName: string, address: string) => number;
  readonly ensureCellTrackedByCoords: (sheetId: number, row: number, col: number) => number;
  readonly forEachSheetCell: (
    sheetId: number,
    fn: (cellIndex: number, row: number, col: number) => void,
  ) => void;
  readonly scheduleWasmProgramSync: () => void;
  readonly markFormulaChanged: (cellIndex: number, count: number) => number;
  readonly resolveStructuredReference: (
    tableName: string,
    columnName: string,
  ) => FormulaNode | undefined;
  readonly resolveSpillReference: (
    currentSheetName: string,
    sheetName: string | undefined,
    address: string,
  ) => FormulaNode | undefined;
  readonly getDependencyBuildEpoch: () => number;
  readonly setDependencyBuildEpoch: (next: number) => void;
  readonly getDependencyBuildSeen: () => U32;
  readonly setDependencyBuildSeen: (next: U32) => void;
  readonly getDependencyBuildCells: () => U32;
  readonly setDependencyBuildCells: (next: U32) => void;
  readonly getDependencyBuildEntities: () => U32;
  readonly setDependencyBuildEntities: (next: U32) => void;
  readonly getDependencyBuildRanges: () => U32;
  readonly setDependencyBuildRanges: (next: U32) => void;
  readonly getDependencyBuildNewRanges: () => U32;
  readonly setDependencyBuildNewRanges: (next: U32) => void;
  readonly getSymbolicRefBindings: () => U32;
  readonly setSymbolicRefBindings: (next: U32) => void;
  readonly getSymbolicRangeBindings: () => U32;
  readonly setSymbolicRangeBindings: (next: U32) => void;
}): EngineFormulaBindingService {
  const ensureDependencyBuildCapacity = (
    cellCapacity: number,
    dependencyCapacity: number,
    symbolicRefCapacity = 0,
    symbolicRangeCapacity = 0,
  ): void => {
    if (cellCapacity > args.getDependencyBuildSeen().length) {
      args.setDependencyBuildSeen(growUint32(args.getDependencyBuildSeen(), cellCapacity));
    }
    if (cellCapacity > args.getDependencyBuildCells().length) {
      args.setDependencyBuildCells(growUint32(args.getDependencyBuildCells(), cellCapacity));
    }
    if (dependencyCapacity > args.getDependencyBuildEntities().length) {
      args.setDependencyBuildEntities(
        growUint32(args.getDependencyBuildEntities(), dependencyCapacity),
      );
    }
    if (dependencyCapacity > args.getDependencyBuildRanges().length) {
      args.setDependencyBuildRanges(
        growUint32(args.getDependencyBuildRanges(), dependencyCapacity),
      );
    }
    if (dependencyCapacity > args.getDependencyBuildNewRanges().length) {
      args.setDependencyBuildNewRanges(
        growUint32(args.getDependencyBuildNewRanges(), dependencyCapacity),
      );
    }
    if (symbolicRefCapacity > args.getSymbolicRefBindings().length) {
      args.setSymbolicRefBindings(growUint32(args.getSymbolicRefBindings(), symbolicRefCapacity));
    }
    if (symbolicRangeCapacity > args.getSymbolicRangeBindings().length) {
      args.setSymbolicRangeBindings(
        growUint32(args.getSymbolicRangeBindings(), symbolicRangeCapacity),
      );
    }
  };

  const setReverseEdgeSlice = (entityId: number, slice: EdgeSlice): void => {
    const empty = slice.ptr < 0 || slice.len === 0;
    if (isRangeEntity(entityId)) {
      args.reverseState.reverseRangeEdges[entityPayload(entityId)] = empty ? undefined : slice;
      return;
    }
    if (isExactLookupColumnEntity(entityId)) {
      if (empty) {
        args.reverseState.reverseExactLookupColumnEdges.delete(entityPayload(entityId));
      } else {
        args.reverseState.reverseExactLookupColumnEdges.set(entityPayload(entityId), slice);
      }
      return;
    }
    if (isSortedLookupColumnEntity(entityId)) {
      if (empty) {
        args.reverseState.reverseSortedLookupColumnEdges.delete(entityPayload(entityId));
      } else {
        args.reverseState.reverseSortedLookupColumnEdges.set(entityPayload(entityId), slice);
      }
      return;
    }
    args.reverseState.reverseCellEdges[entityPayload(entityId)] = empty ? undefined : slice;
  };

  const getReverseEdgeSlice = (entityId: number): EdgeSlice | undefined => {
    if (isRangeEntity(entityId)) {
      return args.reverseState.reverseRangeEdges[entityPayload(entityId)];
    }
    if (isExactLookupColumnEntity(entityId)) {
      return args.reverseState.reverseExactLookupColumnEdges.get(entityPayload(entityId));
    }
    if (isSortedLookupColumnEntity(entityId)) {
      return args.reverseState.reverseSortedLookupColumnEdges.get(entityPayload(entityId));
    }
    return args.reverseState.reverseCellEdges[entityPayload(entityId)];
  };

  const appendReverseEdge = (entityId: number, dependentEntityId: number): void => {
    const slice = getReverseEdgeSlice(entityId) ?? args.edgeArena.empty();
    setReverseEdgeSlice(entityId, args.edgeArena.appendUnique(slice, dependentEntityId));
  };

  const removeReverseEdge = (entityId: number, dependentEntityId: number): void => {
    const slice = getReverseEdgeSlice(entityId);
    if (!slice) {
      return;
    }
    setReverseEdgeSlice(entityId, args.edgeArena.removeValue(slice, dependentEntityId));
  };

  const appendDefinedNameReverseEdge = (name: string, dependentCellIndex: number): void => {
    appendTrackedReverseEdge(
      args.reverseState.reverseDefinedNameEdges,
      normalizeDefinedName(name),
      dependentCellIndex,
    );
  };

  const removeDefinedNameReverseEdge = (name: string, dependentCellIndex: number): void => {
    removeTrackedReverseEdge(
      args.reverseState.reverseDefinedNameEdges,
      normalizeDefinedName(name),
      dependentCellIndex,
    );
  };

  const pruneTrackedDependencyCell = (cellIndex: number, ownerCellIndex: number): void => {
    if (cellIndex === ownerCellIndex) {
      return;
    }
    if (getReverseEdgeSlice(makeCellEntity(cellIndex))) {
      return;
    }
    args.state.workbook.pruneCellIfEmpty(cellIndex);
  };

  const pruneOrphanedDependencyCells = (cellIndices: readonly number[]): void => {
    cellIndices.forEach((cellIndex) => {
      if (getReverseEdgeSlice(makeCellEntity(cellIndex))) {
        return;
      }
      args.state.workbook.pruneCellIfEmpty(cellIndex);
    });
  };

  const isCellIndexMappedNow = (cellIndex: number): boolean => {
    const sheetId = args.state.workbook.cellStore.sheetIds[cellIndex];
    const row = args.state.workbook.cellStore.rows[cellIndex];
    const col = args.state.workbook.cellStore.cols[cellIndex];
    if (sheetId === undefined || row === undefined || col === undefined) {
      return false;
    }
    const sheet = args.state.workbook.getSheetById(sheetId);
    return sheet?.grid.get(row, col) === cellIndex;
  };

  const compileFormulaForSheet = (
    currentSheetName: string,
    source: string,
  ): ReturnType<typeof compileFormula> => {
    const compiled = compileFormula(source);
    if (compiled.mode === FormulaMode.WasmFastPath) {
      if (
        hasIndexedExactLookupCandidate(compiled.optimizedAst) ||
        hasDirectApproximateLookupCandidate(compiled.optimizedAst)
      ) {
        compiled.mode = FormulaMode.JsOnly;
      }
    }
    if (
      compiled.symbolicNames.length === 0 &&
      compiled.symbolicTables.length === 0 &&
      compiled.symbolicSpills.length === 0
    ) {
      return compiled;
    }

    const resolved = resolveMetadataReferencesInAst(compiled.ast, {
      resolveName: (name) => args.state.workbook.getDefinedName(name)?.value,
      resolveStructuredReference: (tableName, columnName) =>
        args.resolveStructuredReference(tableName, columnName),
      resolveSpillReference: (sheetName, address) =>
        args.resolveSpillReference(currentSheetName, sheetName, address),
    });
    if (!resolved.substituted || !resolved.fullyResolved) {
      return compiled;
    }

    const resolvedCompiled = compileFormulaAst(source, resolved.node, {
      originalAst: compiled.ast,
      symbolicNames: compiled.symbolicNames,
      symbolicTables: compiled.symbolicTables,
      symbolicSpills: compiled.symbolicSpills,
    });
    if (resolvedCompiled.mode === FormulaMode.WasmFastPath) {
      if (
        hasIndexedExactLookupCandidate(resolvedCompiled.optimizedAst) ||
        hasDirectApproximateLookupCandidate(resolvedCompiled.optimizedAst)
      ) {
        resolvedCompiled.mode = FormulaMode.JsOnly;
      }
    }
    return resolvedCompiled;
  };

  const materializeDependencies = (
    currentSheetName: string,
    compiled: ParsedCompiledFormula,
    directLookupBinding: ReturnType<typeof resolveRuntimeDirectLookupBinding> | undefined,
  ): MaterializedDependencies => {
    const deps = compiled.deps;
    const parsedCellDeps = compiled.parsedDeps;
    if (
      compiled.symbolicRanges.length === 0 &&
      parsedCellDeps !== undefined &&
      parsedCellDeps.length === deps.length &&
      parsedCellDeps.length > 0 &&
      parsedCellDeps.length <= 2
    ) {
      ensureDependencyBuildCapacity(
        args.state.workbook.cellStore.size + 1,
        parsedCellDeps.length + 1,
        compiled.symbolicRefs.length + 1,
        1,
      );
      let dependencyIndexCount = 0;
      let dependencyEntityCount = 0;
      for (let depIndex = 0; depIndex < parsedCellDeps.length; depIndex += 1) {
        const parsedDep = parsedCellDeps[depIndex]!;
        const sheetName = parsedDep.sheetName ?? currentSheetName;
        const cellIndex = args.ensureCellTracked(sheetName, parsedDep.address);
        let seen = false;
        for (let existingIndex = 0; existingIndex < dependencyIndexCount; existingIndex += 1) {
          if (args.getDependencyBuildCells()[existingIndex] === cellIndex) {
            seen = true;
            break;
          }
        }
        if (!seen) {
          args.getDependencyBuildCells()[dependencyIndexCount] = cellIndex;
          dependencyIndexCount += 1;
        }
        args.getDependencyBuildEntities()[dependencyEntityCount] = makeCellEntity(cellIndex);
        dependencyEntityCount += 1;
      }
      return {
        dependencyIndices: args.getDependencyBuildCells().slice(0, dependencyIndexCount),
        dependencyEntities: args.getDependencyBuildEntities().slice(0, dependencyEntityCount),
        rangeDependencies: args.getDependencyBuildRanges().slice(0, 0),
        symbolicRangeIndices: args.getSymbolicRangeBindings(),
        symbolicRangeCount: 0,
        newRangeIndices: args.getDependencyBuildNewRanges(),
        newRangeCount: 0,
      };
    }

    ensureDependencyBuildCapacity(
      args.state.workbook.cellStore.size + 1,
      deps.length + 1,
      compiled.symbolicRefs.length + 1,
      compiled.symbolicRanges.length + 1,
    );
    let epoch = args.getDependencyBuildEpoch() + 1;
    if (epoch === 0xffff_ffff) {
      epoch = 1;
      args.getDependencyBuildSeen().fill(0);
    }
    args.setDependencyBuildEpoch(epoch);

    let dependencyIndexCount = 0;
    let dependencyEntityCount = 0;
    let rangeDependencyCount = 0;
    let newRangeCount = 0;
    args
      .getSymbolicRangeBindings()
      .fill(UNRESOLVED_WASM_OPERAND, 0, compiled.symbolicRanges.length);

    for (let depIndex = 0; depIndex < deps.length; depIndex += 1) {
      const dep = deps[depIndex]!;
      const parsedDep = compiled.parsedDeps?.[depIndex];
      if (parsedDep?.kind === "cell") {
        const sheetName = parsedDep.sheetName ?? currentSheetName;
        const cellIndex = args.ensureCellTracked(sheetName, parsedDep.address);
        if (args.getDependencyBuildSeen()[cellIndex] !== epoch) {
          args.getDependencyBuildSeen()[cellIndex] = epoch;
          args.getDependencyBuildCells()[dependencyIndexCount] = cellIndex;
          dependencyIndexCount += 1;
        }
        args.getDependencyBuildEntities()[dependencyEntityCount] = makeCellEntity(cellIndex);
        dependencyEntityCount += 1;
        continue;
      }
      if (dep.includes(":")) {
        const range = parseRangeAddress(dep, currentSheetName);
        const sheetName = range.sheetName ?? currentSheetName;
        const isDirectLookupColumn =
          directLookupBinding !== undefined &&
          range.kind === "cells" &&
          range.start.col === range.end.col &&
          sheetName === directLookupBinding.lookupSheetName &&
          range.start.col === directLookupBinding.col &&
          range.start.row === directLookupBinding.rowStart &&
          range.end.row === directLookupBinding.rowEnd;
        if (isDirectLookupColumn) {
          const sheet = args.state.workbook.getSheet(sheetName);
          if (sheet) {
            for (let row = range.start.row; row <= range.end.row; row += 1) {
              const cellIndex = sheet.grid.get(row, range.start.col);
              if (cellIndex === -1) {
                continue;
              }
              if ((args.state.workbook.cellStore.formulaIds[cellIndex] ?? 0) === 0) {
                continue;
              }
              if (args.getDependencyBuildSeen()[cellIndex] !== epoch) {
                args.getDependencyBuildSeen()[cellIndex] = epoch;
                args.getDependencyBuildCells()[dependencyIndexCount] = cellIndex;
                dependencyIndexCount += 1;
              }
              args.getDependencyBuildEntities()[dependencyEntityCount] = makeCellEntity(cellIndex);
              dependencyEntityCount += 1;
            }
          }
          continue;
        }
        const symbolicRangeIndex = compiled.symbolicRanges.indexOf(dep);
        if (range.sheetName && !args.state.workbook.getSheet(sheetName)) {
          continue;
        }
        const sheet = args.state.workbook.getSheet(sheetName);
        if (!sheet) {
          continue;
        }
        const registered = args.state.ranges.intern(sheet.id, range, {
          ensureCell: (sheetId, row, col) => args.ensureCellTrackedByCoords(sheetId, row, col),
          forEachSheetCell: (sheetId, fn) => args.forEachSheetCell(sheetId, fn),
        });
        if (symbolicRangeIndex !== -1) {
          args.getSymbolicRangeBindings()[symbolicRangeIndex] = registered.rangeIndex;
        }
        const rangeEntity = makeRangeEntity(registered.rangeIndex);
        args.getDependencyBuildEntities()[dependencyEntityCount] = rangeEntity;
        dependencyEntityCount += 1;
        args.getDependencyBuildRanges()[rangeDependencyCount] = registered.rangeIndex;
        rangeDependencyCount += 1;
        const memberIndices = args.state.ranges.expandToCells(registered.rangeIndex);
        for (let memberIndex = 0; memberIndex < memberIndices.length; memberIndex += 1) {
          const cellIndex = memberIndices[memberIndex]!;
          if (args.getDependencyBuildSeen()[cellIndex] === epoch) {
            continue;
          }
          args.getDependencyBuildSeen()[cellIndex] = epoch;
          args.getDependencyBuildCells()[dependencyIndexCount] = cellIndex;
          dependencyIndexCount += 1;
        }
        if (registered.materialized) {
          args.getDependencyBuildNewRanges()[newRangeCount] = registered.rangeIndex;
          newRangeCount += 1;
        }
        continue;
      }
      const parsed = parseCellAddress(dep, currentSheetName);
      const sheetName = parsed.sheetName ?? currentSheetName;
      if (parsed.sheetName && !args.state.workbook.getSheet(sheetName)) {
        continue;
      }
      const cellIndex = args.ensureCellTracked(sheetName, parsed.text);
      if (args.getDependencyBuildSeen()[cellIndex] !== epoch) {
        args.getDependencyBuildSeen()[cellIndex] = epoch;
        args.getDependencyBuildCells()[dependencyIndexCount] = cellIndex;
        dependencyIndexCount += 1;
      }
      args.getDependencyBuildEntities()[dependencyEntityCount] = makeCellEntity(cellIndex);
      dependencyEntityCount += 1;
    }
    return {
      dependencyIndices: args.getDependencyBuildCells().slice(0, dependencyIndexCount),
      dependencyEntities: args.getDependencyBuildEntities().slice(0, dependencyEntityCount),
      rangeDependencies: args.getDependencyBuildRanges().slice(0, rangeDependencyCount),
      symbolicRangeIndices: args.getSymbolicRangeBindings(),
      symbolicRangeCount: compiled.symbolicRanges.length,
      newRangeIndices: args.getDependencyBuildNewRanges(),
      newRangeCount,
    };
  };

  const clearFormulaNow = (cellIndex: number): boolean => {
    const existing = args.state.formulas.get(cellIndex);
    if (existing) {
      const dependencyEntities = args.edgeArena.readView(existing.dependencyEntities);
      const formulaEntity = makeCellEntity(cellIndex);
      for (let index = 0; index < dependencyEntities.length; index += 1) {
        const dependencyEntity = dependencyEntities[index]!;
        removeReverseEdge(dependencyEntity, formulaEntity);
        if (!isRangeEntity(dependencyEntity)) {
          pruneTrackedDependencyCell(entityPayload(dependencyEntity), cellIndex);
        }
      }
      existing.compiled.symbolicNames.forEach((name) => {
        removeDefinedNameReverseEdge(name, cellIndex);
      });
      existing.compiled.symbolicTables.forEach((name) => {
        removeTrackedReverseEdge(
          args.reverseState.reverseTableEdges,
          tableDependencyKey(name),
          cellIndex,
        );
      });
      const ownerSheetName = args.state.workbook.getSheetNameById(
        args.state.workbook.cellStore.sheetIds[cellIndex]!,
      );
      existing.compiled.symbolicSpills.forEach((key) => {
        removeTrackedReverseEdge(
          args.reverseState.reverseSpillEdges,
          spillDependencyKeyFromRef(key, ownerSheetName),
          cellIndex,
        );
      });
      const existingDirectLookup = existing.directLookup;
      if (existingDirectLookup) {
        const lookupInfo = directLookupColumnInfo(existingDirectLookup);
        const lookupSheet = args.state.workbook.getSheet(lookupInfo.sheetName);
        if (lookupSheet) {
          const lookupEntity = lookupInfo.isExact
            ? makeExactLookupColumnEntity(lookupSheet.id, lookupInfo.col)
            : makeSortedLookupColumnEntity(lookupSheet.id, lookupInfo.col);
          removeReverseEdge(lookupEntity, formulaEntity);
          if (
            existingDirectLookup.kind === "exact" ||
            existingDirectLookup.kind === "approximate" ||
            existingDirectLookup.kind === "exact-uniform-numeric" ||
            existingDirectLookup.kind === "approximate-uniform-numeric"
          ) {
            const rowStart =
              existingDirectLookup.kind === "exact" || existingDirectLookup.kind === "approximate"
                ? existingDirectLookup.prepared.rowStart
                : existingDirectLookup.rowStart;
            const rowEnd =
              existingDirectLookup.kind === "exact" || existingDirectLookup.kind === "approximate"
                ? existingDirectLookup.prepared.rowEnd
                : existingDirectLookup.rowEnd;
            for (let row = rowStart; row <= rowEnd; row += 1) {
              const memberCellIndex = args.ensureCellTrackedByCoords(
                lookupSheet.id,
                row,
                lookupInfo.col,
              );
              removeReverseEdge(makeCellEntity(memberCellIndex), lookupEntity);
            }
          }
        }
      }
      for (let index = 0; index < existing.rangeDependencies.length; index += 1) {
        const rangeIndex = existing.rangeDependencies[index]!;
        const released = args.state.ranges.release(rangeIndex);
        if (!released.removed) {
          continue;
        }
        const rangeEntity = makeRangeEntity(rangeIndex);
        for (let memberIndex = 0; memberIndex < released.members.length; memberIndex += 1) {
          const memberCellIndex = released.members[memberIndex]!;
          removeReverseEdge(makeCellEntity(memberCellIndex), rangeEntity);
          pruneTrackedDependencyCell(memberCellIndex, cellIndex);
        }
        setReverseEdgeSlice(rangeEntity, args.edgeArena.empty());
      }
      args.edgeArena.free(existing.dependencyEntities);
      args.compiledPlans.release(existing.planId);
    }
    args.state.formulas.delete(cellIndex);
    args.state.workbook.cellStore.flags[cellIndex] =
      (args.state.workbook.cellStore.flags[cellIndex] ?? 0) &
      ~(
        CellFlags.HasFormula |
        CellFlags.JsOnly |
        CellFlags.InCycle |
        CellFlags.SpillChild |
        CellFlags.PivotOutput
      );
    args.scheduleWasmProgramSync();
    return existing !== undefined;
  };

  const bindFormulaNow = (cellIndex: number, ownerSheetName: string, source: string): boolean => {
    const compiled = compileFormulaForSheet(ownerSheetName, source) as ParsedCompiledFormula;
    const directLookupBinding = resolveRuntimeDirectLookupBinding(compiled.jsPlan, ownerSheetName);
    const indexedExactLookupCandidates = args.state.useColumnIndex
      ? collectIndexedExactLookupCandidates(compiled.optimizedAst)
      : [];
    const directApproximateLookupCandidates = collectDirectApproximateLookupCandidates(
      compiled.optimizedAst,
    );
    const dependencies = materializeDependencies(ownerSheetName, compiled, directLookupBinding);
    const existing = args.state.formulas.get(cellIndex);
    const topologyChanged =
      existing === undefined ||
      !uint32ArrayEqual(
        args.edgeArena.readView(existing.dependencyEntities),
        dependencies.dependencyEntities,
      ) ||
      !uint32ArrayEqual(existing.rangeDependencies, dependencies.rangeDependencies) ||
      !stringArrayEqual(existing.compiled.symbolicNames, compiled.symbolicNames) ||
      !stringArrayEqual(existing.compiled.symbolicTables, compiled.symbolicTables) ||
      !stringArrayEqual(existing.compiled.symbolicSpills, compiled.symbolicSpills);
    clearFormulaNow(cellIndex);

    ensureDependencyBuildCapacity(
      args.state.workbook.cellStore.size + 1,
      compiled.deps.length + 1,
      compiled.symbolicRefs.length + 1,
      compiled.symbolicRanges.length + 1,
    );
    for (let index = 0; index < compiled.symbolicRefs.length; index += 1) {
      const parsedRef = compiled.parsedSymbolicRefs?.[index];
      if (parsedRef && parsedRef.sheetName === undefined) {
        args.getSymbolicRefBindings()[index] = args.ensureCellTracked(
          ownerSheetName,
          parsedRef.address,
        );
        continue;
      }
      const ref = compiled.symbolicRefs[index]!;
      const [qualifiedSheetName, qualifiedAddress] = ref.includes("!")
        ? ref.split("!")
        : [undefined, ref];
      const fallbackAddress = parseCellAddress(qualifiedAddress, qualifiedSheetName).text;
      const sheetName =
        parsedRef?.sheetName ??
        qualifiedSheetName ??
        args.state.workbook.getSheetNameById(args.state.workbook.cellStore.sheetIds[cellIndex]!);
      if (
        (parsedRef?.sheetName ?? qualifiedSheetName) &&
        !args.state.workbook.getSheet(sheetName)
      ) {
        args.getSymbolicRefBindings()[index] = UNRESOLVED_WASM_OPERAND;
        continue;
      }
      args.getSymbolicRefBindings()[index] = args.ensureCellTracked(
        sheetName,
        parsedRef?.address ?? fallbackAddress,
      );
    }

    const literalStringIds = compiled.symbolicStrings.map((value) =>
      args.state.strings.intern(value),
    );
    const runtimeProgram = new Uint32Array(compiled.program.length);
    runtimeProgram.set(compiled.program);
    compiled.program.forEach((instruction, index) => {
      const opcode = instruction >>> 24;
      const operand = instruction & 0x00ff_ffff;
      if (opcode === PUSH_CELL_OPCODE) {
        const targetIndex =
          operand < compiled.symbolicRefs.length
            ? (args.getSymbolicRefBindings()[operand] ?? 0)
            : 0;
        runtimeProgram[index] = (PUSH_CELL_OPCODE << 24) | (targetIndex & 0x00ff_ffff);
        return;
      }
      if (opcode === PUSH_RANGE_OPCODE) {
        const targetIndex =
          operand < dependencies.symbolicRangeCount
            ? (dependencies.symbolicRangeIndices[operand] ?? 0)
            : 0;
        runtimeProgram[index] = (PUSH_RANGE_OPCODE << 24) | (targetIndex & 0x00ff_ffff);
        return;
      }
      if (opcode === PUSH_STRING_OPCODE) {
        const stringId = operand < literalStringIds.length ? (literalStringIds[operand] ?? 0) : 0;
        runtimeProgram[index] = (PUSH_STRING_OPCODE << 24) | (stringId & 0x00ff_ffff);
      }
    });

    const dependencyEntities = args.edgeArena.replace(
      args.edgeArena.empty(),
      dependencies.dependencyEntities,
    );
    const directLookup = buildDirectLookupDescriptor({
      compiled,
      ownerSheetName,
      workbook: args.state.workbook,
      ensureCellTracked: args.ensureCellTracked,
      exactLookup: args.exactLookup,
      sortedLookup: args.sortedLookup,
    });
    const plan = args.compiledPlans.intern(source, compiled);
    const runtimeFormula: RuntimeFormula = {
      cellIndex,
      formulaSlotId: 0,
      planId: plan.id,
      source,
      compiled: plan.compiled,
      plan,
      dependencyIndices: dependencies.dependencyIndices,
      dependencyEntities,
      rangeDependencies: dependencies.rangeDependencies,
      runtimeProgram,
      constants: compiled.constants,
      programOffset: 0,
      programLength: runtimeProgram.length,
      constNumberOffset: 0,
      constNumberLength: compiled.constants.length,
      rangeListOffset: 0,
      rangeListLength: dependencies.rangeDependencies.length,
      directLookup,
    };
    const formulaSlotId = args.state.formulas.set(cellIndex, runtimeFormula);
    runtimeFormula.formulaSlotId = formulaSlotId;
    args.state.workbook.cellStore.flags[cellIndex] =
      ((args.state.workbook.cellStore.flags[cellIndex] ?? 0) &
        ~(CellFlags.SpillChild | CellFlags.PivotOutput)) |
      CellFlags.HasFormula;
    if (runtimeFormula.compiled.mode === FormulaMode.JsOnly) {
      args.state.workbook.cellStore.flags[cellIndex] =
        (args.state.workbook.cellStore.flags[cellIndex] ?? 0) | CellFlags.JsOnly;
    } else {
      args.state.workbook.cellStore.flags[cellIndex] =
        (args.state.workbook.cellStore.flags[cellIndex] ?? 0) & ~CellFlags.JsOnly;
    }

    for (let rangeCursor = 0; rangeCursor < dependencies.newRangeCount; rangeCursor += 1) {
      const rangeIndex = dependencies.newRangeIndices[rangeCursor]!;
      const memberIndices = args.state.ranges.expandToCells(rangeIndex);
      const rangeEntity = makeRangeEntity(rangeIndex);
      for (let index = 0; index < memberIndices.length; index += 1) {
        appendReverseEdge(makeCellEntity(memberIndices[index]!), rangeEntity);
      }
    }
    const formulaEntity = makeCellEntity(cellIndex);
    for (let index = 0; index < dependencies.dependencyEntities.length; index += 1) {
      appendReverseEdge(dependencies.dependencyEntities[index]!, formulaEntity);
    }
    runtimeFormula.compiled.symbolicNames.forEach((name) => {
      appendDefinedNameReverseEdge(name, cellIndex);
    });
    runtimeFormula.compiled.symbolicTables.forEach((name) => {
      appendTrackedReverseEdge(
        args.reverseState.reverseTableEdges,
        tableDependencyKey(name),
        cellIndex,
      );
    });
    runtimeFormula.compiled.symbolicSpills.forEach((key) => {
      appendTrackedReverseEdge(
        args.reverseState.reverseSpillEdges,
        spillDependencyKeyFromRef(
          key,
          args.state.workbook.getSheetNameById(args.state.workbook.cellStore.sheetIds[cellIndex]!),
        ),
        cellIndex,
      );
    });
    if (directLookup) {
      const lookupInfo = directLookupColumnInfo(directLookup);
      const lookupSheet = args.state.workbook.getSheet(lookupInfo.sheetName);
      if (lookupSheet) {
        const lookupEntity = lookupInfo.isExact
          ? makeExactLookupColumnEntity(lookupSheet.id, lookupInfo.col)
          : makeSortedLookupColumnEntity(lookupSheet.id, lookupInfo.col);
        const rowStart =
          directLookup.kind === "exact" || directLookup.kind === "approximate"
            ? directLookup.prepared.rowStart
            : directLookup.rowStart;
        const rowEnd =
          directLookup.kind === "exact" || directLookup.kind === "approximate"
            ? directLookup.prepared.rowEnd
            : directLookup.rowEnd;
        for (let row = rowStart; row <= rowEnd; row += 1) {
          const memberCellIndex = args.ensureCellTrackedByCoords(
            lookupSheet.id,
            row,
            lookupInfo.col,
          );
          appendReverseEdge(makeCellEntity(memberCellIndex), lookupEntity);
        }
        appendReverseEdge(lookupEntity, formulaEntity);
      }
    }
    args.scheduleWasmProgramSync();

    indexedExactLookupCandidates.forEach((candidate) => {
      if (candidate.startCol !== candidate.endCol) {
        return;
      }
      args.exactLookup.primeColumnIndex({
        sheetName: candidate.sheetName ?? ownerSheetName,
        rowStart: candidate.startRow,
        rowEnd: candidate.endRow,
        col: candidate.startCol,
      });
    });
    directApproximateLookupCandidates.forEach((candidate) => {
      if (candidate.startCol !== candidate.endCol) {
        return;
      }
      args.sortedLookup.primeColumnIndex({
        sheetName: candidate.sheetName ?? ownerSheetName,
        rowStart: candidate.startRow,
        rowEnd: candidate.endRow,
        col: candidate.startCol,
      });
    });
    return topologyChanged;
  };

  const invalidateFormulaNow = (cellIndex: number): void => {
    clearFormulaNow(cellIndex);
    args.state.workbook.cellStore.setValue(cellIndex, errorValue(ErrorCode.Value));
    args.state.workbook.cellStore.flags[cellIndex] =
      (args.state.workbook.cellStore.flags[cellIndex] ?? 0) &
      ~(
        CellFlags.HasFormula |
        CellFlags.JsOnly |
        CellFlags.InCycle |
        CellFlags.SpillChild |
        CellFlags.PivotOutput
      );
  };

  const rebindFormulaCellsNow = (
    candidates: readonly number[],
    formulaChangedCount: number,
  ): number => {
    candidates.forEach((cellIndex) => {
      const formula = args.state.formulas.get(cellIndex);
      const ownerSheetName = args.state.workbook.getSheetNameById(
        args.state.workbook.cellStore.sheetIds[cellIndex]!,
      );
      if (formula && ownerSheetName) {
        bindFormulaNow(cellIndex, ownerSheetName, formula.source);
      }
      formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount);
    });
    return formulaChangedCount;
  };

  const rebindTrackedDependentsNow = (
    registry: Map<string, Set<number>>,
    keys: readonly string[],
    formulaChangedCount: number,
  ): number => rebindFormulaCellsNow(collectTrackedDependents(registry, keys), formulaChangedCount);

  const rebindFormulasForSheetNow = (
    sheetName: string,
    formulaChangedCount: number,
    candidates?: readonly number[] | U32,
  ): number => {
    if (candidates) {
      for (let index = 0; index < candidates.length; index += 1) {
        const cellIndex = candidates[index]!;
        const formula = args.state.formulas.get(cellIndex);
        if (!formula) {
          continue;
        }
        const ownerSheetName = args.state.workbook.getSheetNameById(
          args.state.workbook.cellStore.sheetIds[cellIndex]!,
        );
        if (!ownerSheetName) {
          continue;
        }
        const touchesSheet = formula.compiled.deps.some((dep) => {
          if (!dep.includes("!")) {
            return false;
          }
          const [qualifiedSheet] = dep.split("!");
          return qualifiedSheet?.replace(/^'(.*)'$/, "$1") === sheetName;
        });
        if (!touchesSheet) {
          continue;
        }
        bindFormulaNow(cellIndex, ownerSheetName, formula.source);
        formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount);
      }
      return formulaChangedCount;
    }

    args.state.formulas.forEach((formula, cellIndex) => {
      if (!formula) {
        return;
      }
      const ownerSheetName = args.state.workbook.getSheetNameById(
        args.state.workbook.cellStore.sheetIds[cellIndex]!,
      );
      if (!ownerSheetName) {
        return;
      }
      const touchesSheet = formula.compiled.deps.some((dep) => {
        if (!dep.includes("!")) {
          return false;
        }
        const [qualifiedSheet] = dep.split("!");
        return qualifiedSheet?.replace(/^'(.*)'$/, "$1") === sheetName;
      });
      if (!touchesSheet) {
        return;
      }
      bindFormulaNow(cellIndex, ownerSheetName, formula.source);
      formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount);
    });

    return formulaChangedCount;
  };

  return {
    bindFormula(cellIndex, ownerSheetName, source) {
      return Effect.try({
        try: () => {
          return bindFormulaNow(cellIndex, ownerSheetName, source);
        },
        catch: (cause) =>
          new EngineFormulaBindingError({
            message: formulaBindingErrorMessage("Failed to bind formula", cause),
            cause,
          }),
      });
    },
    clearFormula(cellIndex) {
      return Effect.try({
        try: () => clearFormulaNow(cellIndex),
        catch: (cause) =>
          new EngineFormulaBindingError({
            message: formulaBindingErrorMessage("Failed to clear formula", cause),
            cause,
          }),
      });
    },
    invalidateFormula(cellIndex) {
      return Effect.try({
        try: () => {
          invalidateFormulaNow(cellIndex);
        },
        catch: (cause) =>
          new EngineFormulaBindingError({
            message: formulaBindingErrorMessage("Failed to invalidate formula", cause),
            cause,
          }),
      });
    },
    rewriteCellFormulasForSheetRename(oldSheetName, newSheetName, formulaChangedCount) {
      return Effect.try({
        try: () => {
          args.state.formulas.forEach((formula, cellIndex) => {
            if (!formula) {
              return;
            }
            const ownerSheetName = args.state.workbook.getSheetNameById(
              args.state.workbook.cellStore.sheetIds[cellIndex]!,
            );
            if (!ownerSheetName) {
              return;
            }
            const nextSource = renameFormulaSheetReferences(
              formula.source,
              oldSheetName,
              newSheetName,
            );
            if (nextSource === formula.source && ownerSheetName !== newSheetName) {
              return;
            }
            bindFormulaNow(cellIndex, ownerSheetName, nextSource);
            formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount);
          });
          return formulaChangedCount;
        },
        catch: (cause) =>
          new EngineFormulaBindingError({
            message: formulaBindingErrorMessage(
              "Failed to rewrite formulas for sheet rename",
              cause,
            ),
            cause,
          }),
      });
    },
    rebuildAllFormulaBindings() {
      return Effect.try({
        try: () => {
          const pending = [...args.state.formulas.entries()].map(([cellIndex, formula]) => ({
            cellIndex,
            source: formula.source,
            dependencyIndices: [...formula.dependencyIndices],
            planId: formula.planId,
          }));
          pending.forEach(({ planId }) => {
            args.compiledPlans.release(planId);
          });
          args.state.formulas.clear();
          args.state.ranges.reset();
          args.edgeArena.reset();
          args.programArena.reset();
          args.constantArena.reset();
          args.rangeListArena.reset();
          args.reverseState.reverseCellEdges.length = 0;
          args.reverseState.reverseRangeEdges.length = 0;
          args.reverseState.reverseDefinedNameEdges.clear();
          args.reverseState.reverseTableEdges.clear();
          args.reverseState.reverseSpillEdges.clear();
          args.reverseState.reverseExactLookupColumnEdges.clear();
          args.reverseState.reverseSortedLookupColumnEdges.clear();

          const activeCellIndices: number[] = [];
          pending.forEach(({ cellIndex, source }) => {
            if (!isCellIndexMappedNow(cellIndex)) {
              args.state.workbook.pruneCellIfEmpty(cellIndex);
              return;
            }
            const ownerSheetName = args.state.workbook.getSheetNameById(
              args.state.workbook.cellStore.sheetIds[cellIndex]!,
            );
            if (!ownerSheetName || !args.state.workbook.getSheet(ownerSheetName)) {
              return;
            }
            try {
              bindFormulaNow(cellIndex, ownerSheetName, source);
            } catch {
              invalidateFormulaNow(cellIndex);
            }
            activeCellIndices.push(cellIndex);
          });
          pruneOrphanedDependencyCells(
            pending.flatMap(({ dependencyIndices }) => dependencyIndices),
          );
          return activeCellIndices;
        },
        catch: (cause) =>
          new EngineFormulaBindingError({
            message: formulaBindingErrorMessage("Failed to rebuild formula bindings", cause),
            cause,
          }),
      });
    },
    rebindFormulaCells(candidates, formulaChangedCount) {
      return Effect.try({
        try: () => rebindFormulaCellsNow(candidates, formulaChangedCount),
        catch: (cause) =>
          new EngineFormulaBindingError({
            message: formulaBindingErrorMessage("Failed to rebind formula cells", cause),
            cause,
          }),
      });
    },
    rebindDefinedNameDependents(names, formulaChangedCount) {
      return Effect.try({
        try: () =>
          rebindTrackedDependentsNow(
            args.reverseState.reverseDefinedNameEdges,
            names.map((name) => normalizeDefinedName(name)),
            formulaChangedCount,
          ),
        catch: (cause) =>
          new EngineFormulaBindingError({
            message: formulaBindingErrorMessage("Failed to rebind defined name dependents", cause),
            cause,
          }),
      });
    },
    rebindTableDependents(tableNames, formulaChangedCount) {
      return Effect.try({
        try: () =>
          rebindTrackedDependentsNow(
            args.reverseState.reverseTableEdges,
            tableNames,
            formulaChangedCount,
          ),
        catch: (cause) =>
          new EngineFormulaBindingError({
            message: formulaBindingErrorMessage("Failed to rebind table dependents", cause),
            cause,
          }),
      });
    },
    rebindFormulasForSheet(sheetName, formulaChangedCount, candidates) {
      return Effect.try({
        try: () => rebindFormulasForSheetNow(sheetName, formulaChangedCount, candidates),
        catch: (cause) =>
          new EngineFormulaBindingError({
            message: formulaBindingErrorMessage("Failed to rebind formulas for sheet", cause),
            cause,
          }),
      });
    },
    bindFormulaNow,
    clearFormulaNow,
    invalidateFormulaNow,
    rebindFormulaCellsNow,
    rebindDefinedNameDependentsNow(names, formulaChangedCount) {
      return rebindFormulaCellsNow(
        collectTrackedDependents(args.reverseState.reverseDefinedNameEdges, names),
        formulaChangedCount,
      );
    },
    rebindTableDependentsNow(tableNames, formulaChangedCount) {
      const normalized = tableNames.map((name) => tableDependencyKey(name));
      return rebindFormulaCellsNow(
        collectTrackedDependents(args.reverseState.reverseTableEdges, normalized),
        formulaChangedCount,
      );
    },
    rebindFormulasForSheetNow,
  };
}
