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
import { isRangeEntity, makeCellEntity, makeRangeEntity, entityPayload } from "../../entity-ids.js";
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
  type RuntimeFormula,
  UNRESOLVED_WASM_OPERAND,
  type U32,
} from "../runtime-state.js";
import { EngineFormulaBindingError } from "../errors.js";
import type { Uint32Arena, Float64Arena } from "@bilig/formula/program-arena";

export interface EngineFormulaBindingService {
  readonly bindFormula: (
    cellIndex: number,
    ownerSheetName: string,
    source: string,
  ) => Effect.Effect<void, EngineFormulaBindingError>;
  readonly clearFormula: (
    cellIndex: number,
  ) => Effect.Effect<boolean, EngineFormulaBindingError>;
  readonly invalidateFormula: (
    cellIndex: number,
  ) => Effect.Effect<void, EngineFormulaBindingError>;
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

const PUSH_CELL_OPCODE = Number(Opcode.PushCell);
const PUSH_RANGE_OPCODE = Number(Opcode.PushRange);
const PUSH_STRING_OPCODE = Number(Opcode.PushString);

export function createEngineFormulaBindingService(args: {
  readonly state: Pick<EngineRuntimeState, "workbook" | "strings" | "formulas" | "ranges">;
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
      args.setDependencyBuildRanges(growUint32(args.getDependencyBuildRanges(), dependencyCapacity));
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
      args.setSymbolicRangeBindings(growUint32(args.getSymbolicRangeBindings(), symbolicRangeCapacity));
    }
  };

  const setReverseEdgeSlice = (entityId: number, slice: EdgeSlice): void => {
    const empty = slice.ptr < 0 || slice.len === 0;
    if (isRangeEntity(entityId)) {
      args.reverseState.reverseRangeEdges[entityPayload(entityId)] = empty ? undefined : slice;
      return;
    }
    args.reverseState.reverseCellEdges[entityPayload(entityId)] = empty ? undefined : slice;
  };

  const getReverseEdgeSlice = (entityId: number): EdgeSlice | undefined => {
    if (isRangeEntity(entityId)) {
      return args.reverseState.reverseRangeEdges[entityPayload(entityId)];
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

  const compileFormulaForSheet = (
    currentSheetName: string,
    source: string,
  ): ReturnType<typeof compileFormula> => {
    const compiled = compileFormula(source);
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

    return compileFormulaAst(source, resolved.node, {
      originalAst: compiled.ast,
      symbolicNames: compiled.symbolicNames,
      symbolicTables: compiled.symbolicTables,
      symbolicSpills: compiled.symbolicSpills,
    });
  };

  const materializeDependencies = (
    currentSheetName: string,
    compiled: ReturnType<typeof compileFormula>,
  ): MaterializedDependencies => {
    const deps = compiled.deps;
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

    for (const dep of deps) {
      if (dep.includes(":")) {
        const range = parseRangeAddress(dep, currentSheetName);
        const sheetName = range.sheetName ?? currentSheetName;
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
        removeReverseEdge(dependencyEntities[index]!, formulaEntity);
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
      for (let index = 0; index < existing.rangeDependencies.length; index += 1) {
        const rangeIndex = existing.rangeDependencies[index]!;
        const released = args.state.ranges.release(rangeIndex);
        if (!released.removed) {
          continue;
        }
        const rangeEntity = makeRangeEntity(rangeIndex);
        for (let memberIndex = 0; memberIndex < released.members.length; memberIndex += 1) {
          removeReverseEdge(makeCellEntity(released.members[memberIndex]!), rangeEntity);
        }
        setReverseEdgeSlice(rangeEntity, args.edgeArena.empty());
      }
      args.edgeArena.free(existing.dependencyEntities);
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

  const bindFormulaNow = (cellIndex: number, ownerSheetName: string, source: string): void => {
    const compiled = compileFormulaForSheet(ownerSheetName, source);
    const dependencies = materializeDependencies(ownerSheetName, compiled);
    clearFormulaNow(cellIndex);

    ensureDependencyBuildCapacity(
      args.state.workbook.cellStore.size + 1,
      compiled.deps.length + 1,
      compiled.symbolicRefs.length + 1,
      compiled.symbolicRanges.length + 1,
    );
    for (let index = 0; index < compiled.symbolicRefs.length; index += 1) {
      const ref = compiled.symbolicRefs[index]!;
      const [qualifiedSheetName, qualifiedAddress] = ref.includes("!")
        ? ref.split("!")
        : [undefined, ref];
      const parsed = parseCellAddress(qualifiedAddress, qualifiedSheetName);
      const sheetName =
        qualifiedSheetName ??
        args.state.workbook.getSheetNameById(args.state.workbook.cellStore.sheetIds[cellIndex]!);
      if (qualifiedSheetName && !args.state.workbook.getSheet(sheetName)) {
        args.getSymbolicRefBindings()[index] = UNRESOLVED_WASM_OPERAND;
        continue;
      }
      args.getSymbolicRefBindings()[index] = args.ensureCellTracked(sheetName, parsed.text);
    }

    const literalStringIds = compiled.symbolicStrings.map((value) => args.state.strings.intern(value));
    const runtimeProgram = new Uint32Array(compiled.program.length);
    runtimeProgram.set(compiled.program);
    compiled.program.forEach((instruction, index) => {
      const opcode = instruction >>> 24;
      const operand = instruction & 0x00ff_ffff;
      if (opcode === PUSH_CELL_OPCODE) {
        const targetIndex =
          operand < compiled.symbolicRefs.length ? (args.getSymbolicRefBindings()[operand] ?? 0) : 0;
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
    const runtimeFormula: RuntimeFormula = {
      cellIndex,
      source,
      compiled: {
        ...compiled,
        depsPtr: dependencyEntities.ptr,
        depsLen: dependencyEntities.len,
      },
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
    };
    const formulaId = args.state.formulas.set(cellIndex, runtimeFormula);
    runtimeFormula.compiled.id = formulaId;
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
      appendTrackedReverseEdge(args.reverseState.reverseTableEdges, tableDependencyKey(name), cellIndex);
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
    args.scheduleWasmProgramSync();
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
          bindFormulaNow(cellIndex, ownerSheetName, source);
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
          }));
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

          const activeCellIndices: number[] = [];
          pending.forEach(({ cellIndex, source }) => {
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
  };
}
