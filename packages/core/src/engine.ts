import {
  ErrorCode,
  FormulaMode,
  ValueTag,
  type CellIndex,
  type CellSnapshot,
  type CellValue,
  type DependencySnapshot,
  type EngineEvent,
  type ExplainCellSnapshot,
  type LiteralInput,
  type RecalcMetrics,
  type SelectionState,
  type WorkbookSnapshot
} from "@bilig/protocol";
import {
  compileFormula,
  evaluatePlan,
  formatAddress,
  parseCellAddress,
  parseRangeAddress
} from "@bilig/formula";
import { Float64Arena, Uint32Arena } from "@bilig/formula/program-arena";
import {
  batchOpOrder,
  compareOpOrder,
  createBatch,
  createReplicaState,
  exportReplicaSnapshot as exportReplicaStateSnapshot,
  hydrateReplicaState,
  markBatchApplied,
  shouldApplyBatch,
  type EngineOp,
  type EngineOpBatch,
  type OpOrder,
  type ReplicaSnapshot,
  type ReplicaVersionSnapshot,
  type ReplicaState
} from "@bilig/crdt";
import { CellFlags } from "./cell-store.js";
import { CycleDetector } from "./cycle-detection.js";
import { EdgeArena, type EdgeSlice } from "./edge-arena.js";
import { entityPayload, isRangeEntity, makeCellEntity, makeRangeEntity } from "./entity-ids.js";
import { EngineEventBus } from "./events.js";
import { FormulaTable } from "./formula-table.js";
import { RangeRegistry } from "./range-registry.js";
import { RecalcScheduler } from "./scheduler.js";
import { selectCellSnapshot, selectMetrics, selectSelectionState, selectViewportCells } from "./selectors.js";
import { StringPool } from "./string-pool.js";
import { WasmKernelFacade } from "./wasm-facade.js";
import { WorkbookStore } from "./workbook-store.js";
import { cellToCsvValue, parseCsv, parseCsvCellInput, serializeCsv } from "./csv.js";

export interface CommitOp {
  kind: "upsertWorkbook" | "upsertSheet" | "deleteSheet" | "upsertCell" | "deleteCell";
  name?: string;
  order?: number;
  sheetName?: string;
  addr?: string;
  value?: LiteralInput;
  formula?: string;
  format?: string;
}

export interface SpreadsheetEngineOptions {
  workbookName?: string;
  replicaId?: string;
}

export interface EngineReplicaSnapshot {
  replica: ReplicaSnapshot;
  entityVersions: ReplicaVersionSnapshot[];
  sheetDeleteVersions: Array<{ sheetName: string; order: OpOrder }>;
}

interface RuntimeFormula {
  cellIndex: number;
  source: string;
  compiled: ReturnType<typeof compileFormula>;
  dependencyIndices: Uint32Array;
  dependencyEntities: EdgeSlice;
  rangeDependencies: Uint32Array;
  runtimeProgram: Uint32Array;
  constants: Float64Array;
  programOffset: number;
  programLength: number;
  constNumberOffset: number;
  constNumberLength: number;
  rangeListOffset: number;
  rangeListLength: number;
}

type U32 = Uint32Array<ArrayBufferLike>;

interface MaterializedDependencies {
  dependencyIndices: Uint32Array;
  dependencyEntities: Uint32Array;
  rangeDependencies: Uint32Array;
  symbolicRangeIndices: U32;
  symbolicRangeCount: number;
  newRangeIndices: U32;
  newRangeCount: number;
}

function emptyValue(): CellValue {
  return { tag: ValueTag.Empty };
}

function errorValue(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code };
}

function literalToValue(input: LiteralInput, stringPool: StringPool): CellValue {
  if (input === null) return emptyValue();
  if (typeof input === "number") return { tag: ValueTag.Number, value: input };
  if (typeof input === "boolean") return { tag: ValueTag.Boolean, value: input };
  return { tag: ValueTag.String, value: input, stringId: stringPool.intern(input) };
}

export class SpreadsheetEngine {
  readonly workbook: WorkbookStore;
  readonly strings = new StringPool();
  readonly events = new EngineEventBus();
  readonly replica: ReplicaState;
  readonly ranges = new RangeRegistry();
  readonly scheduler = new RecalcScheduler();
  readonly wasm = new WasmKernelFacade();

  private readonly formulas: FormulaTable<RuntimeFormula>;
  private readonly cycleDetector = new CycleDetector();
  private readonly edgeArena = new EdgeArena();
  private readonly programArena = new Uint32Arena();
  private readonly constantArena = new Float64Arena();
  private readonly rangeListArena = new Uint32Arena();
  private reverseCellEdges: Array<EdgeSlice | undefined> = [];
  private reverseRangeEdges: Array<EdgeSlice | undefined> = [];
  private readonly batchListeners = new Set<(batch: EngineOpBatch) => void>();
  private readonly selectionListeners = new Set<() => void>();
  private readonly entityVersions = new Map<string, OpOrder>();
  private readonly sheetDeleteVersions = new Map<string, OpOrder>();
  private selection: SelectionState = { sheetName: "Sheet1", address: "A1" };
  private pendingKernelSync: U32 = new Uint32Array(128);
  private wasmBatch: U32 = new Uint32Array(128);
  private mutationRoots: U32 = new Uint32Array(128);
  private changedInputEpoch = 1;
  private changedInputSeen: U32 = new Uint32Array(128);
  private changedInputBuffer: U32 = new Uint32Array(128);
  private changedFormulaEpoch = 1;
  private changedFormulaSeen: U32 = new Uint32Array(128);
  private changedFormulaBuffer: U32 = new Uint32Array(128);
  private changedUnionEpoch = 1;
  private changedUnionSeen: U32 = new Uint32Array(128);
  private changedUnion: U32 = new Uint32Array(128);
  private materializedCellCount = 0;
  private materializedCells: U32 = new Uint32Array(128);
  private explicitChangedEpoch = 1;
  private explicitChangedSeen: U32 = new Uint32Array(128);
  private explicitChangedBuffer: U32 = new Uint32Array(128);
  private dependencyBuildEpoch = 1;
  private dependencyBuildSeen: U32 = new Uint32Array(128);
  private dependencyBuildCells: U32 = new Uint32Array(128);
  private dependencyBuildEntities: U32 = new Uint32Array(128);
  private dependencyBuildRanges: U32 = new Uint32Array(128);
  private dependencyBuildNewRanges: U32 = new Uint32Array(128);
  private symbolicRefBindings: U32 = new Uint32Array(128);
  private symbolicRangeBindings: U32 = new Uint32Array(128);
  private impactedFormulaEpoch = 1;
  private impactedFormulaSeen: U32 = new Uint32Array(128);
  private impactedFormulaBuffer: U32 = new Uint32Array(128);
  private wasmProgramTargets: U32 = new Uint32Array(128);
  private wasmProgramOffsets: U32 = new Uint32Array(128);
  private wasmProgramLengths: U32 = new Uint32Array(128);
  private wasmConstantOffsets: U32 = new Uint32Array(128);
  private wasmConstantLengths: U32 = new Uint32Array(128);
  private wasmRangeOffsets: U32 = new Uint32Array(128);
  private wasmRangeLengths: U32 = new Uint32Array(128);
  private topoIndegree: U32 = new Uint32Array(128);
  private topoQueue: U32 = new Uint32Array(128);
  private topoFormulaBuffer: U32 = new Uint32Array(128);
  private topoEntityQueue: U32 = new Uint32Array(128);
  private topoFormulaSeenEpoch = 1;
  private topoRangeSeenEpoch = 1;
  private topoFormulaSeen: U32 = new Uint32Array(128);
  private topoRangeSeen: U32 = new Uint32Array(128);
  private batchMutationDepth = 0;
  private wasmProgramSyncPending = false;
  private lastMetrics: RecalcMetrics = {
    batchId: 0,
    changedInputCount: 0,
    dirtyFormulaCount: 0,
    wasmFormulaCount: 0,
    jsFormulaCount: 0,
    rangeNodeVisits: 0,
    recalcMs: 0,
    compileMs: 0
  };

  constructor(options: SpreadsheetEngineOptions = {}) {
    this.workbook = new WorkbookStore(options.workbookName ?? "Workbook");
    this.formulas = new FormulaTable(this.workbook.cellStore);
    this.replica = createReplicaState(options.replicaId ?? "local");
    void this.wasm.init();
  }

  async ready(): Promise<void> {
    await this.wasm.init();
  }

  subscribe(listener: (event: EngineEvent) => void): () => void {
    return this.events.subscribe(listener);
  }

  subscribeCell(sheetName: string, address: string, listener: () => void): () => void {
    const cellIndex = this.workbook.getCellIndex(sheetName, address);
    if (cellIndex !== undefined) {
      return this.events.subscribeCellIndex(cellIndex, listener);
    }
    return this.events.subscribeCellAddress(`${sheetName}!${address}`, listener);
  }

  subscribeCells(sheetName: string, addresses: readonly string[], listener: () => void): () => void {
    const cellIndices: number[] = [];
    const qualifiedAddresses: string[] = [];
    addresses.forEach((address) => {
      const cellIndex = this.workbook.getCellIndex(sheetName, address);
      if (cellIndex !== undefined) {
        cellIndices.push(cellIndex);
        return;
      }
      qualifiedAddresses.push(`${sheetName}!${address}`);
    });
    return this.events.subscribeCells(cellIndices, qualifiedAddresses, listener);
  }

  subscribeBatches(listener: (batch: EngineOpBatch) => void): () => void {
    this.batchListeners.add(listener);
    return () => {
      this.batchListeners.delete(listener);
    };
  }

  subscribeSelection(listener: () => void): () => void {
    this.selectionListeners.add(listener);
    return () => {
      this.selectionListeners.delete(listener);
    };
  }

  getSelectionState(): SelectionState {
    return this.selection;
  }

  setSelection(sheetName: string, address: string | null): void {
    if (this.selection.sheetName === sheetName && this.selection.address === address) {
      return;
    }
    this.selection = { sheetName, address };
    this.selectionListeners.forEach((listener) => listener());
  }

  getLastMetrics(): RecalcMetrics {
    return this.lastMetrics;
  }

  createSheet(name: string): void {
    this.applyLocalOps([{ kind: "upsertSheet", name, order: this.workbook.sheetsByName.size }]);
  }

  deleteSheet(name: string): void {
    this.applyLocalOps([{ kind: "deleteSheet", name }]);
  }

  setCellValue(sheetName: string, address: string, value: LiteralInput): CellValue {
    this.applyLocalOps([{ kind: "setCellValue", sheetName, address, value }]);
    return this.getCellValue(sheetName, address);
  }

  setCellFormula(sheetName: string, address: string, formula: string): CellValue {
    this.applyLocalOps([{ kind: "setCellFormula", sheetName, address, formula }]);
    return this.getCellValue(sheetName, address);
  }

  setCellFormat(sheetName: string, address: string, format: string | null): void {
    this.applyLocalOps([{ kind: "setCellFormat", sheetName, address, format }]);
  }

  clearCell(sheetName: string, address: string): void {
    this.applyLocalOps([{ kind: "clearCell", sheetName, address }]);
  }

  exportSheetCsv(sheetName: string): string {
    const sheet = this.workbook.getSheet(sheetName);
    if (!sheet) {
      return "";
    }

    let maxRow = -1;
    let maxCol = -1;
    const cells = new Map<string, string>();

    sheet.grid.forEachCell((cellIndex) => {
      const cell = this.getCellByIndex(cellIndex);
      const parsed = parseCellAddress(cell.address, sheetName);
      maxRow = Math.max(maxRow, parsed.row);
      maxCol = Math.max(maxCol, parsed.col);
      cells.set(`${parsed.row}:${parsed.col}`, cellToCsvValue(cell));
    });

    if (maxRow < 0 || maxCol < 0) {
      return "";
    }

    const rows = Array.from({ length: maxRow + 1 }, (_, row) =>
      Array.from({ length: maxCol + 1 }, (_, col) => cells.get(`${row}:${col}`) ?? "")
    );

    return serializeCsv(rows);
  }

  importSheetCsv(sheetName: string, csv: string): void {
    const rows = parseCsv(csv);
    const existingSheet = this.workbook.getSheet(sheetName);
    const order = existingSheet?.order ?? this.workbook.sheetsByName.size;
    const ops: EngineOp[] = [];
    let potentialNewCells = 0;

    if (existingSheet) {
      ops.push({ kind: "deleteSheet", name: sheetName });
    }
    ops.push({ kind: "upsertSheet", name: sheetName, order });

    rows.forEach((row, rowIndex) => {
      row.forEach((raw, colIndex) => {
        const parsed = parseCsvCellInput(raw);
        if (!parsed) {
          return;
        }
        const address = formatAddress(rowIndex, colIndex);
        if (parsed.formula !== undefined) {
          ops.push({ kind: "setCellFormula", sheetName, address, formula: parsed.formula });
          potentialNewCells += 1;
          return;
        }
        ops.push({ kind: "setCellValue", sheetName, address, value: parsed.value ?? null });
        potentialNewCells += 1;
      });
    });

    this.applyLocalOps(ops, potentialNewCells);
  }

  getCellValue(sheetName: string, address: string): CellValue {
    const cellIndex = this.workbook.getCellIndex(sheetName, address);
    if (cellIndex === undefined) {
      return emptyValue();
    }
    return this.workbook.cellStore.getValue(cellIndex, (id) => this.strings.get(id));
  }

  getCell(sheetName: string, address: string): CellSnapshot {
    const cellIndex = this.workbook.getCellIndex(sheetName, address);
    if (cellIndex === undefined) {
      return {
        sheetName,
        address,
        value: emptyValue(),
        flags: 0,
        version: 0
      };
    }
    return this.getCellByIndex(cellIndex);
  }

  getCellByIndex(cellIndex: number): CellSnapshot {
    const address = this.workbook.getAddress(cellIndex);
    const sheetName = this.workbook.getSheetNameById(this.workbook.cellStore.sheetIds[cellIndex]!);
    const snapshot: CellSnapshot = {
      sheetName,
      address,
      value: this.workbook.cellStore.getValue(cellIndex, (id) => this.strings.get(id)),
      flags: this.workbook.cellStore.flags[cellIndex]!,
      version: this.workbook.cellStore.versions[cellIndex] ?? 0
    };
    const format = this.workbook.getCellFormat(cellIndex);
    if (format !== undefined) {
      snapshot.format = format;
    }
    const formula = this.formulas.get(cellIndex)?.source;
    if (formula !== undefined) {
      snapshot.formula = formula;
    }
    return snapshot;
  }

  getDependencies(sheetName: string, address: string): DependencySnapshot {
    const cellIndex = this.workbook.getCellIndex(sheetName, address);
    if (cellIndex === undefined) return { directDependents: [], directPrecedents: [] };
    const directDependents = new Set<number>();
    const directPrecedents: number[] = [];
    this.forEachFormulaDependencyCell(cellIndex, (dependencyCellIndex) => {
      directPrecedents.push(dependencyCellIndex);
    });
    const dependents = this.getEntityDependents(makeCellEntity(cellIndex));
    for (let index = 0; index < dependents.length; index += 1) {
      const dependent = dependents[index]!;
      if (isRangeEntity(dependent)) {
        const rangeDependents = this.getEntityDependents(dependent);
        for (let rangeIndex = 0; rangeIndex < rangeDependents.length; rangeIndex += 1) {
          const formulaEntity = rangeDependents[rangeIndex]!;
          if (!isRangeEntity(formulaEntity)) {
            directDependents.add(entityPayload(formulaEntity));
          }
        }
        continue;
      }
      directDependents.add(entityPayload(dependent));
    }
    return {
      directPrecedents: directPrecedents.map((dependencyCellIndex) =>
        this.workbook.getQualifiedAddress(dependencyCellIndex)
      ),
      directDependents: [...directDependents].map((dependentCellIndex) =>
        this.workbook.getQualifiedAddress(dependentCellIndex)
      )
    };
  }

  getDependents(sheetName: string, address: string): DependencySnapshot {
    return this.getDependencies(sheetName, address);
  }

  explainCell(sheetName: string, address: string): ExplainCellSnapshot {
    const cellIndex = this.workbook.getCellIndex(sheetName, address);
    if (cellIndex === undefined) {
      return {
        sheetName,
        address,
        value: emptyValue(),
        flags: 0,
        version: 0,
        inCycle: false,
        directPrecedents: [],
        directDependents: []
      };
    }

    const snapshot = this.getCellByIndex(cellIndex);
    const formula = this.formulas.get(cellIndex);
    const flags = this.workbook.cellStore.flags[cellIndex] ?? 0;
    const isFormula = ((flags & CellFlags.HasFormula) !== 0) && formula !== undefined;
    const dependencies = this.getDependencies(sheetName, address);

    const explanation: ExplainCellSnapshot = {
      ...snapshot,
      version: this.workbook.cellStore.versions[cellIndex] ?? 0,
      inCycle: (flags & CellFlags.InCycle) !== 0,
      directPrecedents: dependencies.directPrecedents,
      directDependents: dependencies.directDependents
    };

    if (formula?.source !== undefined) {
      explanation.formula = formula.source;
    }
    if (isFormula) {
      explanation.mode = formula.compiled.mode;
      explanation.topoRank = this.workbook.cellStore.topoRanks[cellIndex] ?? 0;
    }

    return explanation;
  }

  exportSnapshot(): WorkbookSnapshot {
    return {
      version: 1,
      workbook: { name: this.workbook.workbookName },
      sheets: [...this.workbook.sheetsByName.values()]
        .sort((left, right) => left.order - right.order)
        .map((sheet) => {
          const cells: WorkbookSnapshot["sheets"][number]["cells"] = [];
          sheet.grid.forEachCell((cellIndex) => {
            const snapshot = this.getCellByIndex(cellIndex);
            const cell: WorkbookSnapshot["sheets"][number]["cells"][number] = {
              address: snapshot.address
            };
            if (snapshot.format !== undefined) {
              cell.format = snapshot.format;
            }
            if (snapshot.formula) {
              cell.formula = snapshot.formula;
            } else if (snapshot.value.tag === ValueTag.Number) {
              cell.value = snapshot.value.value;
            } else if (snapshot.value.tag === ValueTag.Boolean) {
              cell.value = snapshot.value.value;
            } else if (snapshot.value.tag === ValueTag.String) {
              cell.value = snapshot.value.value;
            } else {
              cell.value = null;
            }
            cells.push(cell);
          });
          return { name: sheet.name, order: sheet.order, cells };
        })
    };
  }

  importSnapshot(snapshot: WorkbookSnapshot): void {
    this.resetWorkbook(snapshot.workbook.name);
    const totalCells = snapshot.sheets.reduce((count, sheet) => count + sheet.cells.length, 0);
    this.workbook.cellStore.ensureCapacity(totalCells);

    snapshot.sheets.forEach((sheet) => {
      this.workbook.createSheet(sheet.name, sheet.order);
    });

    this.beginMutationCollection();
    let changedInputCount = 0;
    let formulaChangedCount = 0;
    let compileMs = 0;
    this.resetMaterializedCellScratch(totalCells);

    this.batchMutationDepth += 1;
    try {
      snapshot.sheets.forEach((sheet) => {
        sheet.cells.forEach((cell) => {
          const cellIndex = this.ensureCellTracked(sheet.name, cell.address);
          if (cell.formula !== undefined) {
            const compileStarted = performance.now();
            const compiled = compileFormula(cell.formula);
            compileMs += performance.now() - compileStarted;
            const dependencies = this.materializeDependencies(sheet.name, compiled);
            this.setFormula(cellIndex, cell.formula, compiled, dependencies);
            formulaChangedCount = this.markFormulaChanged(cellIndex, formulaChangedCount);
          } else {
            const value = literalToValue(cell.value ?? null, this.strings);
            this.workbook.cellStore.setValue(
              cellIndex,
              value,
              value.tag === ValueTag.String ? value.stringId : 0
            );
            this.workbook.cellStore.flags[cellIndex] =
              (this.workbook.cellStore.flags[cellIndex] ?? 0) & ~CellFlags.HasFormula;
            changedInputCount = this.markInputChanged(cellIndex, changedInputCount);
          }
          if (cell.format !== undefined) {
            this.workbook.setCellFormat(cellIndex, cell.format);
          }
        });
      });

      formulaChangedCount = this.syncDynamicRanges(formulaChangedCount);
    } finally {
      this.batchMutationDepth -= 1;
      this.flushWasmProgramSync();
    }

    this.lastMetrics.compileMs = compileMs;
    if (formulaChangedCount > 0) {
      this.rebuildTopoRanks();
      this.detectCycles();
    }

    const changedInputArray = this.changedInputBuffer.subarray(0, changedInputCount);
    const changed = this.recalculate(
      this.composeMutationRoots(changedInputCount, formulaChangedCount),
      changedInputArray
    );
    this.lastMetrics.batchId += 1;
    this.lastMetrics.changedInputCount = changedInputCount + formulaChangedCount;

    const event: EngineEvent = {
      kind: "batch",
      changedCellIndices: changed,
      metrics: this.lastMetrics
    };
    if (this.events.hasCellListeners()) {
      this.events.emitAllWatched(event);
      return;
    }
    this.events.emit(event, changed, (cellIndex) => this.workbook.getQualifiedAddress(cellIndex));
  }

  exportReplicaSnapshot(): EngineReplicaSnapshot {
    return {
      replica: exportReplicaStateSnapshot(this.replica),
      entityVersions: [...this.entityVersions.entries()].map(([entityKey, order]) => ({ entityKey, order })),
      sheetDeleteVersions: [...this.sheetDeleteVersions.entries()].map(([sheetName, order]) => ({ sheetName, order }))
    };
  }

  importReplicaSnapshot(snapshot: EngineReplicaSnapshot): void {
    hydrateReplicaState(this.replica, snapshot.replica);
    this.entityVersions.clear();
    snapshot.entityVersions.forEach(({ entityKey, order }) => {
      this.entityVersions.set(entityKey, order);
    });
    this.sheetDeleteVersions.clear();
    snapshot.sheetDeleteVersions.forEach(({ sheetName, order }) => {
      this.sheetDeleteVersions.set(sheetName, order);
    });
  }

  renderCommit(ops: CommitOp[]): void {
    const engineOps: EngineOp[] = [];
    let potentialNewCells = 0;
    ops.forEach((op) => {
      switch (op.kind) {
        case "upsertWorkbook":
          if (op.name) engineOps.push({ kind: "upsertWorkbook", name: op.name });
          break;
        case "upsertSheet":
          if (op.name) engineOps.push({ kind: "upsertSheet", name: op.name, order: op.order ?? 0 });
          break;
        case "deleteSheet":
          if (op.name) engineOps.push({ kind: "deleteSheet", name: op.name });
          break;
        case "upsertCell":
          if (!op.sheetName || !op.addr) break;
          if (op.formula !== undefined) {
            engineOps.push({ kind: "setCellFormula", sheetName: op.sheetName, address: op.addr, formula: op.formula });
          } else {
            engineOps.push({ kind: "setCellValue", sheetName: op.sheetName, address: op.addr, value: op.value ?? null });
          }
          potentialNewCells += 1;
          if (op.format !== undefined) {
            engineOps.push({ kind: "setCellFormat", sheetName: op.sheetName, address: op.addr, format: op.format });
          }
          break;
        case "deleteCell":
          if (op.sheetName && op.addr) {
            engineOps.push({ kind: "clearCell", sheetName: op.sheetName, address: op.addr });
            engineOps.push({ kind: "setCellFormat", sheetName: op.sheetName, address: op.addr, format: null });
          }
          break;
      }
    });
    this.applyLocalOps(engineOps, potentialNewCells);
  }

  applyRemoteBatch(batch: EngineOpBatch): void {
    if (!shouldApplyBatch(this.replica, batch)) return;
    this.applyBatch(batch, "remote");
  }

  private applyLocalOps(ops: EngineOp[], potentialNewCells?: number): void {
    if (ops.length === 0) return;
    const batch = createBatch(this.replica, ops);
    this.applyBatch(batch, "local", potentialNewCells);
  }

  private applyBatch(batch: EngineOpBatch, source: "local" | "remote", potentialNewCells?: number): void {
    this.beginMutationCollection();
    let changedInputCount = 0;
    let formulaChangedCount = 0;
    let explicitChangedCount = 0;
    let topologyChanged = false;
    let sheetDeleted = false;
    let appliedOps = 0;
    const canSkipOrderChecks = source === "local";

    const reservedNewCells = potentialNewCells ?? this.estimatePotentialNewCells(batch.ops);
    this.workbook.cellStore.ensureCapacity(this.workbook.cellStore.size + reservedNewCells);
    this.resetMaterializedCellScratch(reservedNewCells);

    this.batchMutationDepth += 1;
    try {
      batch.ops.forEach((op, opIndex) => {
        const order = batchOpOrder(batch, opIndex);
        if (!canSkipOrderChecks && !this.shouldApplyOp(op, order)) {
          return;
        }

        switch (op.kind) {
          case "upsertWorkbook":
            this.workbook.workbookName = op.name;
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          case "upsertSheet":
            this.workbook.createSheet(op.name, op.order);
            this.entityVersions.set(this.entityKeyForOp(op), order);
            const tombstone = this.sheetDeleteVersions.get(op.name);
            if (!tombstone || compareOpOrder(order, tombstone) > 0) {
              this.sheetDeleteVersions.delete(op.name);
            }
            const reboundCount = formulaChangedCount;
            formulaChangedCount = this.rebindFormulasForSheet(op.name, formulaChangedCount);
            topologyChanged = topologyChanged || formulaChangedCount !== reboundCount;
            break;
          case "deleteSheet":
            const removal = this.removeSheetRuntime(op.name, explicitChangedCount);
            changedInputCount += removal.changedInputCount;
            formulaChangedCount += removal.formulaChangedCount;
            explicitChangedCount = removal.explicitChangedCount;
            this.entityVersions.set(this.entityKeyForOp(op), order);
            this.sheetDeleteVersions.set(op.name, order);
            topologyChanged = true;
            sheetDeleted = true;
            break;
          case "setCellValue": {
            const cellIndex = this.ensureCellTracked(op.sheetName, op.address);
            topologyChanged = this.removeFormula(cellIndex) || topologyChanged;
            const value = literalToValue(op.value, this.strings);
            this.workbook.cellStore.setValue(
              cellIndex,
              value,
              value.tag === ValueTag.String ? value.stringId : 0
            );
            this.workbook.cellStore.flags[cellIndex] =
              (this.workbook.cellStore.flags[cellIndex] ?? 0) & ~CellFlags.HasFormula;
            changedInputCount = this.markInputChanged(cellIndex, changedInputCount);
            explicitChangedCount = this.markExplicitChanged(cellIndex, explicitChangedCount);
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          }
          case "setCellFormula": {
            const cellIndex = this.ensureCellTracked(op.sheetName, op.address);
            const compileStarted = performance.now();
            const compiled = compileFormula(op.formula);
            this.lastMetrics.compileMs = performance.now() - compileStarted;
            const dependencies = this.materializeDependencies(op.sheetName, compiled);
            this.setFormula(cellIndex, op.formula, compiled, dependencies);
            formulaChangedCount = this.markFormulaChanged(cellIndex, formulaChangedCount);
            explicitChangedCount = this.markExplicitChanged(cellIndex, explicitChangedCount);
            this.entityVersions.set(this.entityKeyForOp(op), order);
            topologyChanged = true;
            break;
          }
          case "setCellFormat": {
            const cellIndex = this.ensureCellTracked(op.sheetName, op.address);
            this.workbook.setCellFormat(cellIndex, op.format);
            explicitChangedCount = this.markExplicitChanged(cellIndex, explicitChangedCount);
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          }
          case "clearCell": {
            const cellIndex = this.workbook.getCellIndex(op.sheetName, op.address);
            if (cellIndex === undefined) {
              this.entityVersions.set(this.entityKeyForOp(op), order);
              break;
            }
            topologyChanged = this.removeFormula(cellIndex) || topologyChanged;
            this.workbook.cellStore.setValue(cellIndex, emptyValue());
            changedInputCount = this.markInputChanged(cellIndex, changedInputCount);
            explicitChangedCount = this.markExplicitChanged(cellIndex, explicitChangedCount);
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          }
        }
        appliedOps += 1;
      });

      const reboundCount = formulaChangedCount;
      formulaChangedCount = this.syncDynamicRanges(formulaChangedCount);
      topologyChanged = topologyChanged || formulaChangedCount !== reboundCount;
    } finally {
      this.batchMutationDepth -= 1;
      this.flushWasmProgramSync();
    }

    markBatchApplied(this.replica, batch);
    if (appliedOps === 0) {
      if (source === "local") {
        this.emitBatch(batch);
      }
      return;
    }

    if (topologyChanged) {
      this.rebuildTopoRanks();
      this.detectCycles();
    }
    const changedInputArray = this.changedInputBuffer.subarray(0, changedInputCount);
    const recalculated = this.recalculate(
      this.composeMutationRoots(changedInputCount, formulaChangedCount),
      changedInputArray
    );
    const changed = this.composeEventChanges(recalculated, explicitChangedCount);
    this.lastMetrics.batchId += 1;
    this.lastMetrics.changedInputCount = changedInputCount + formulaChangedCount;
    const event = {
      kind: "batch",
      changedCellIndices: changed,
      metrics: this.lastMetrics
    } satisfies EngineEvent;
    if (sheetDeleted) {
      this.events.emitAllWatched(event);
    } else {
      this.events.emit(event, changed, (cellIndex) => this.workbook.getQualifiedAddress(cellIndex));
    }
    if (source === "local") {
      this.emitBatch(batch);
    }
  }

  private materializeDependencies(
    currentSheet: string,
    compiled: ReturnType<typeof compileFormula>
  ): MaterializedDependencies {
    const deps = compiled.deps;
    this.ensureDependencyBuildCapacity(
      this.workbook.cellStore.size + 1,
      deps.length + 1,
      compiled.symbolicRefs.length + 1,
      compiled.symbolicRanges.length + 1
    );
    this.dependencyBuildEpoch += 1;
    if (this.dependencyBuildEpoch === 0xffff_ffff) {
      this.dependencyBuildEpoch = 1;
      this.dependencyBuildSeen.fill(0);
    }

    let dependencyIndexCount = 0;
    let dependencyEntityCount = 0;
    let rangeDependencyCount = 0;
    let newRangeCount = 0;
    this.symbolicRangeBindings.fill(0, 0, compiled.symbolicRanges.length);
    for (const dep of deps) {
      if (dep.includes(":")) {
        const range = parseRangeAddress(dep, currentSheet);
        const sheetName = range.sheetName ?? currentSheet;
        if (range.sheetName && !this.workbook.getSheet(sheetName)) {
          continue;
        }
        const sheet = this.workbook.getSheet(sheetName);
        if (!sheet) {
          continue;
        }
        const registered = this.ranges.intern(sheet.id, range, {
          ensureCell: (sheetId, row, col) => this.ensureCellTrackedByCoords(sheetId, row, col),
          forEachSheetCell: (sheetId, fn) => this.forEachSheetCell(sheetId, fn)
        });
        const symbolicRangeIndex = compiled.symbolicRanges.indexOf(dep);
        if (symbolicRangeIndex !== -1) {
          this.symbolicRangeBindings[symbolicRangeIndex] = registered.rangeIndex;
        }
        const rangeEntity = makeRangeEntity(registered.rangeIndex);
        this.dependencyBuildEntities[dependencyEntityCount] = rangeEntity;
        dependencyEntityCount += 1;
        this.dependencyBuildRanges[rangeDependencyCount] = registered.rangeIndex;
        rangeDependencyCount += 1;
        const memberIndices = this.ranges.expandToCells(registered.rangeIndex);
        for (let memberIndex = 0; memberIndex < memberIndices.length; memberIndex += 1) {
          const cellIndex = memberIndices[memberIndex]!;
          if (this.dependencyBuildSeen[cellIndex] === this.dependencyBuildEpoch) {
            continue;
          }
          this.dependencyBuildSeen[cellIndex] = this.dependencyBuildEpoch;
          this.dependencyBuildCells[dependencyIndexCount] = cellIndex;
          dependencyIndexCount += 1;
        }
        if (registered.materialized) {
          this.dependencyBuildNewRanges[newRangeCount] = registered.rangeIndex;
          newRangeCount += 1;
        }
        continue;
      }
      const parsed = parseCellAddress(dep, currentSheet);
      const sheetName = parsed.sheetName ?? currentSheet;
      if (parsed.sheetName && !this.workbook.getSheet(sheetName)) {
        continue;
      }
      const cellIndex = this.ensureCellTracked(sheetName, parsed.text);
      if (this.dependencyBuildSeen[cellIndex] !== this.dependencyBuildEpoch) {
        this.dependencyBuildSeen[cellIndex] = this.dependencyBuildEpoch;
        this.dependencyBuildCells[dependencyIndexCount] = cellIndex;
        dependencyIndexCount += 1;
      }
      this.dependencyBuildEntities[dependencyEntityCount] = makeCellEntity(cellIndex);
      dependencyEntityCount += 1;
    }
    return {
      dependencyIndices: this.dependencyBuildCells.slice(0, dependencyIndexCount),
      dependencyEntities: this.dependencyBuildEntities.slice(0, dependencyEntityCount),
      rangeDependencies: this.dependencyBuildRanges.slice(0, rangeDependencyCount),
      symbolicRangeIndices: this.symbolicRangeBindings,
      symbolicRangeCount: compiled.symbolicRanges.length,
      newRangeIndices: this.dependencyBuildNewRanges,
      newRangeCount
    };
  }

  private setFormula(
    cellIndex: number,
    source: string,
    compiled: ReturnType<typeof compileFormula>,
    dependencies: MaterializedDependencies
  ): void {
    this.removeFormula(cellIndex);

    this.ensureDependencyBuildCapacity(
      this.workbook.cellStore.size + 1,
      compiled.deps.length + 1,
      compiled.symbolicRefs.length + 1,
      compiled.symbolicRanges.length + 1
    );
    let hasUnresolvedSymbolicRef = false;
    for (let index = 0; index < compiled.symbolicRefs.length; index += 1) {
      const ref = compiled.symbolicRefs[index]!;
      const [qualifiedSheetName, qualifiedAddress] = ref.includes("!") ? ref.split("!") : [undefined, ref];
      const parsed = parseCellAddress(qualifiedAddress!, qualifiedSheetName);
      const sheetName = qualifiedSheetName ?? this.workbook.getSheetNameById(this.workbook.cellStore.sheetIds[cellIndex]!);
      if (qualifiedSheetName && !this.workbook.getSheet(sheetName)) {
        hasUnresolvedSymbolicRef = true;
        this.symbolicRefBindings[index] = 0;
        continue;
      }
      this.symbolicRefBindings[index] = this.ensureCellTracked(sheetName, parsed.text);
    }

    const effectiveCompiled =
      hasUnresolvedSymbolicRef && compiled.mode === FormulaMode.WasmFastPath
        ? { ...compiled, mode: FormulaMode.JsOnly }
        : compiled;

    const runtimeProgram = new Uint32Array(compiled.program.length);
    runtimeProgram.set(compiled.program);
    compiled.program.forEach((instruction, index) => {
      const opcode = instruction >>> 24;
      const operand = instruction & 0x00ff_ffff;
      if (opcode === 3) {
        const targetIndex = operand < compiled.symbolicRefs.length ? this.symbolicRefBindings[operand] ?? 0 : 0;
        runtimeProgram[index] = (opcode << 24) | (targetIndex & 0x00ff_ffff);
        return;
      }
      if (opcode === 4) {
        const targetIndex = operand < dependencies.symbolicRangeCount ? dependencies.symbolicRangeIndices[operand] ?? 0 : 0;
        runtimeProgram[index] = (opcode << 24) | (targetIndex & 0x00ff_ffff);
      }
    });

    const dependencyEntities = this.edgeArena.replace(this.edgeArena.empty(), dependencies.dependencyEntities);
    const runtimeFormula: RuntimeFormula = {
      cellIndex,
      source,
      compiled: {
        ...effectiveCompiled,
        depsPtr: dependencyEntities.ptr,
        depsLen: dependencyEntities.len
      },
      dependencyIndices: dependencies.dependencyIndices,
      dependencyEntities,
      rangeDependencies: dependencies.rangeDependencies,
      runtimeProgram,
      constants: effectiveCompiled.constants,
      programOffset: 0,
      programLength: runtimeProgram.length,
      constNumberOffset: 0,
      constNumberLength: effectiveCompiled.constants.length,
      rangeListOffset: 0,
      rangeListLength: dependencies.rangeDependencies.length
    };
    const formulaId = this.formulas.set(cellIndex, runtimeFormula);
    runtimeFormula.compiled.id = formulaId;
    this.workbook.cellStore.flags[cellIndex] = (this.workbook.cellStore.flags[cellIndex] ?? 0) | CellFlags.HasFormula;
    if (effectiveCompiled.mode === FormulaMode.JsOnly) {
      this.workbook.cellStore.flags[cellIndex] = (this.workbook.cellStore.flags[cellIndex] ?? 0) | CellFlags.JsOnly;
    } else {
      this.workbook.cellStore.flags[cellIndex] = (this.workbook.cellStore.flags[cellIndex] ?? 0) & ~CellFlags.JsOnly;
    }

    for (let rangeCursor = 0; rangeCursor < dependencies.newRangeCount; rangeCursor += 1) {
      const rangeIndex = dependencies.newRangeIndices[rangeCursor]!;
      const memberIndices = this.ranges.expandToCells(rangeIndex);
      const rangeEntity = makeRangeEntity(rangeIndex);
      for (let index = 0; index < memberIndices.length; index += 1) {
        this.appendReverseEdge(makeCellEntity(memberIndices[index]!), rangeEntity);
      }
    }
    const formulaEntity = makeCellEntity(cellIndex);
    for (let index = 0; index < dependencies.dependencyEntities.length; index += 1) {
      this.appendReverseEdge(dependencies.dependencyEntities[index]!, formulaEntity);
    }
    this.scheduleWasmProgramSync();
  }

  private removeFormula(cellIndex: number): boolean {
    const existing = this.formulas.get(cellIndex);
    if (existing) {
      const dependencyEntities = this.edgeArena.readView(existing.dependencyEntities);
      const formulaEntity = makeCellEntity(cellIndex);
      for (let index = 0; index < dependencyEntities.length; index += 1) {
        this.removeReverseEdge(dependencyEntities[index]!, formulaEntity);
      }
      for (let index = 0; index < existing.rangeDependencies.length; index += 1) {
        const rangeIndex = existing.rangeDependencies[index]!;
        const released = this.ranges.release(rangeIndex);
        if (!released.removed) {
          continue;
        }
        const rangeEntity = makeRangeEntity(rangeIndex);
        for (let memberIndex = 0; memberIndex < released.members.length; memberIndex += 1) {
          this.removeReverseEdge(makeCellEntity(released.members[memberIndex]!), rangeEntity);
        }
        this.setReverseEdgeSlice(rangeEntity, this.edgeArena.empty());
      }
      this.edgeArena.free(existing.dependencyEntities);
    }
    this.formulas.delete(cellIndex);
    this.workbook.cellStore.flags[cellIndex] =
      (this.workbook.cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle);
    this.scheduleWasmProgramSync();
    return existing !== undefined;
  }

  private rebuildTopoRanks(): void {
    const requiredCellCapacity = this.workbook.cellStore.size + 1;
    const requiredEntityCapacity = this.workbook.cellStore.size + this.ranges.size + 1;
    this.ensureTopoScratchCapacity(requiredCellCapacity, requiredEntityCapacity, this.ranges.size + 1);

    let queueLength = 0;
    this.formulas.forEach((_formula, cellIndex) => {
      this.topoIndegree[cellIndex] = 0;
      this.workbook.cellStore.topoRanks[cellIndex] = 0;
    });
    this.formulas.forEach((formula, cellIndex) => {
      for (let index = 0; index < formula.dependencyIndices.length; index += 1) {
        const dependency = formula.dependencyIndices[index]!;
        if ((this.workbook.cellStore.formulaIds[dependency] ?? 0) !== 0) {
          this.topoIndegree[cellIndex] = (this.topoIndegree[cellIndex] ?? 0) + 1;
        }
      }
    });
    this.formulas.forEach((_formula, cellIndex) => {
      if ((this.topoIndegree[cellIndex] ?? 0) === 0) {
        this.topoQueue[queueLength] = cellIndex;
        queueLength += 1;
      }
    });

    let rank = 0;
    for (let queueIndex = 0; queueIndex < queueLength; queueIndex += 1) {
      const cellIndex = this.topoQueue[queueIndex]!;
      this.workbook.cellStore.topoRanks[cellIndex] = rank++;
      const dependentCount = this.collectFormulaDependentsForEntityInto(makeCellEntity(cellIndex));
      for (let dependentIndex = 0; dependentIndex < dependentCount; dependentIndex += 1) {
        const dependent = this.topoFormulaBuffer[dependentIndex]!;
        if ((this.workbook.cellStore.formulaIds[dependent] ?? 0) === 0) {
          continue;
        }
        const next = (this.topoIndegree[dependent] ?? 0) - 1;
        this.topoIndegree[dependent] = next;
        if (next === 0) {
          this.topoQueue[queueLength] = dependent;
          queueLength += 1;
        }
      }
    }
  }

  private detectCycles(): void {
    const result = this.cycleDetector.detect(
      this.formulas.keys(),
      this.workbook.cellStore.size + 1,
      (cellIndex, fn) => this.forEachFormulaDependencyCell(cellIndex, fn),
      (cellIndex) => this.formulas.has(cellIndex)
    );

    this.formulas.forEach((_formula, cellIndex) => {
      this.workbook.cellStore.flags[cellIndex] = (this.workbook.cellStore.flags[cellIndex] ?? 0) & ~CellFlags.InCycle;
      this.workbook.cellStore.cycleGroupIds[cellIndex] = -1;
    });

    for (let index = 0; index < result.cycleMemberCount; index += 1) {
      const cellIndex = result.cycleMembers[index]!;
      this.workbook.cellStore.flags[cellIndex] = (this.workbook.cellStore.flags[cellIndex] ?? 0) | CellFlags.InCycle;
      this.workbook.cellStore.cycleGroupIds[cellIndex] = result.cycleGroups[cellIndex] ?? -1;
      this.workbook.cellStore.setValue(cellIndex, errorValue(ErrorCode.Cycle));
    }
  }

  private recalculate(
    changedRoots: readonly number[] | U32,
    kernelSyncRoots: readonly number[] | U32 = changedRoots
  ): Uint32Array {
    const started = performance.now();
    const scheduled = this.scheduler.collectDirty(
      changedRoots,
      { getDependents: (entityId) => this.getEntityDependents(entityId) },
      this.workbook.cellStore,
      (cellIndex) => this.formulas.has(cellIndex),
      this.ranges.size
    );
    const ordered = scheduled.orderedFormulaCellIndices;
    const orderedCount = scheduled.orderedFormulaCount;

    this.ensureRecalcScratchCapacity(Math.max(this.workbook.cellStore.size + 1, changedRoots.length + orderedCount + 1));

    let pendingKernelSyncCount = 0;
    for (let index = 0; index < kernelSyncRoots.length; index += 1) {
      this.pendingKernelSync[pendingKernelSyncCount] = kernelSyncRoots[index]!;
      pendingKernelSyncCount += 1;
    }

    const flushWasmBatch = (batchCount: number): number => {
      if (batchCount === 0) {
        return 0;
      }
      this.wasm.syncFromStore(this.workbook.cellStore, this.pendingKernelSync.subarray(0, pendingKernelSyncCount));
      pendingKernelSyncCount = 0;
      const batchIndices = this.wasmBatch.subarray(0, batchCount);
      this.wasm.evalBatch(batchIndices);
      this.wasm.syncToStore(this.workbook.cellStore, batchIndices);
      return batchCount;
    };

    let wasmCount = 0;
    let jsCount = 0;
    let wasmBatchCount = 0;
    for (let index = 0; index < orderedCount; index += 1) {
      const cellIndex = ordered[index]!;
      const formula = this.formulas.get(cellIndex);
      if (!formula) {
        continue;
      }
      if (((this.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0) {
        continue;
      }
      if (formula.compiled.mode === FormulaMode.WasmFastPath && this.wasm.ready) {
        this.wasmBatch[wasmBatchCount] = cellIndex;
        wasmBatchCount += 1;
        continue;
      }
      wasmCount += flushWasmBatch(wasmBatchCount);
      wasmBatchCount = 0;
      jsCount += 1;
      this.evaluateFormulaJs(cellIndex, formula);
      this.pendingKernelSync[pendingKernelSyncCount] = cellIndex;
      pendingKernelSyncCount += 1;
    }

    wasmCount += flushWasmBatch(wasmBatchCount);
    if (pendingKernelSyncCount > 0) {
      this.wasm.syncFromStore(this.workbook.cellStore, this.pendingKernelSync.subarray(0, pendingKernelSyncCount));
    }

    this.lastMetrics.dirtyFormulaCount = orderedCount;
    this.lastMetrics.jsFormulaCount = jsCount;
    this.lastMetrics.wasmFormulaCount = wasmCount;
    this.lastMetrics.rangeNodeVisits = scheduled.rangeNodeVisits;
    this.lastMetrics.recalcMs = performance.now() - started;
    return orderedCount === 0 && changedRoots.length === 0
      ? this.changedUnion.subarray(0, 0)
      : this.composeChangedRootsAndOrdered(changedRoots, ordered, orderedCount);
  }

  private ensureRecalcScratchCapacity(size: number): void {
    if (size > this.mutationRoots.length) {
      this.mutationRoots = growUint32(this.mutationRoots, size);
    }
    if (size > this.changedInputSeen.length) {
      this.changedInputSeen = growUint32(this.changedInputSeen, size);
    }
    if (size > this.changedInputBuffer.length) {
      this.changedInputBuffer = growUint32(this.changedInputBuffer, size);
    }
    if (size > this.changedFormulaSeen.length) {
      this.changedFormulaSeen = growUint32(this.changedFormulaSeen, size);
    }
    if (size > this.changedFormulaBuffer.length) {
      this.changedFormulaBuffer = growUint32(this.changedFormulaBuffer, size);
    }
    if (size > this.pendingKernelSync.length) {
      this.pendingKernelSync = growUint32(this.pendingKernelSync, size);
    }
    if (size > this.wasmBatch.length) {
      this.wasmBatch = growUint32(this.wasmBatch, size);
    }
    if (size > this.changedUnion.length) {
      this.changedUnion = growUint32(this.changedUnion, size);
    }
    if (size > this.changedUnionSeen.length) {
      this.changedUnionSeen = growUint32(this.changedUnionSeen, size);
    }
    if (size > this.explicitChangedSeen.length) {
      this.explicitChangedSeen = growUint32(this.explicitChangedSeen, size);
    }
    if (size > this.explicitChangedBuffer.length) {
      this.explicitChangedBuffer = growUint32(this.explicitChangedBuffer, size);
    }
    if (size > this.dependencyBuildSeen.length) {
      this.dependencyBuildSeen = growUint32(this.dependencyBuildSeen, size);
    }
    if (size > this.dependencyBuildCells.length) {
      this.dependencyBuildCells = growUint32(this.dependencyBuildCells, size);
    }
    if (size > this.impactedFormulaSeen.length) {
      this.impactedFormulaSeen = growUint32(this.impactedFormulaSeen, size);
    }
    if (size > this.impactedFormulaBuffer.length) {
      this.impactedFormulaBuffer = growUint32(this.impactedFormulaBuffer, size);
    }
  }

  private ensureDependencyBuildCapacity(
    cellCapacity: number,
    dependencyCapacity: number,
    symbolicRefCapacity = 0,
    symbolicRangeCapacity = 0
  ): void {
    if (cellCapacity > this.dependencyBuildSeen.length) {
      this.dependencyBuildSeen = growUint32(this.dependencyBuildSeen, cellCapacity);
    }
    if (cellCapacity > this.dependencyBuildCells.length) {
      this.dependencyBuildCells = growUint32(this.dependencyBuildCells, cellCapacity);
    }
    if (dependencyCapacity > this.dependencyBuildEntities.length) {
      this.dependencyBuildEntities = growUint32(this.dependencyBuildEntities, dependencyCapacity);
    }
    if (dependencyCapacity > this.dependencyBuildRanges.length) {
      this.dependencyBuildRanges = growUint32(this.dependencyBuildRanges, dependencyCapacity);
    }
    if (dependencyCapacity > this.dependencyBuildNewRanges.length) {
      this.dependencyBuildNewRanges = growUint32(this.dependencyBuildNewRanges, dependencyCapacity);
    }
    if (symbolicRefCapacity > this.symbolicRefBindings.length) {
      this.symbolicRefBindings = growUint32(this.symbolicRefBindings, symbolicRefCapacity);
    }
    if (symbolicRangeCapacity > this.symbolicRangeBindings.length) {
      this.symbolicRangeBindings = growUint32(this.symbolicRangeBindings, symbolicRangeCapacity);
    }
  }

  private ensureWasmProgramScratchCapacity(formulaSize: number, rangeSize: number): void {
    if (formulaSize > this.wasmProgramTargets.length) {
      this.wasmProgramTargets = growUint32(this.wasmProgramTargets, formulaSize);
    }
    if (formulaSize > this.wasmProgramOffsets.length) {
      this.wasmProgramOffsets = growUint32(this.wasmProgramOffsets, formulaSize);
    }
    if (formulaSize > this.wasmProgramLengths.length) {
      this.wasmProgramLengths = growUint32(this.wasmProgramLengths, formulaSize);
    }
    if (formulaSize > this.wasmConstantOffsets.length) {
      this.wasmConstantOffsets = growUint32(this.wasmConstantOffsets, formulaSize);
    }
    if (formulaSize > this.wasmConstantLengths.length) {
      this.wasmConstantLengths = growUint32(this.wasmConstantLengths, formulaSize);
    }
    if (rangeSize > this.wasmRangeOffsets.length) {
      this.wasmRangeOffsets = growUint32(this.wasmRangeOffsets, rangeSize);
    }
    if (rangeSize > this.wasmRangeLengths.length) {
      this.wasmRangeLengths = growUint32(this.wasmRangeLengths, rangeSize);
    }
  }

  private ensureTopoScratchCapacity(cellSize: number, entitySize: number, rangeSize: number): void {
    if (cellSize > this.topoIndegree.length) {
      this.topoIndegree = growUint32(this.topoIndegree, cellSize);
    }
    if (cellSize > this.topoQueue.length) {
      this.topoQueue = growUint32(this.topoQueue, cellSize);
    }
    if (cellSize > this.topoFormulaBuffer.length) {
      this.topoFormulaBuffer = growUint32(this.topoFormulaBuffer, cellSize);
    }
    if (cellSize > this.topoFormulaSeen.length) {
      this.topoFormulaSeen = growUint32(this.topoFormulaSeen, cellSize);
    }
    if (entitySize > this.topoEntityQueue.length) {
      this.topoEntityQueue = growUint32(this.topoEntityQueue, entitySize);
    }
    if (rangeSize > this.topoRangeSeen.length) {
      this.topoRangeSeen = growUint32(this.topoRangeSeen, rangeSize);
    }
  }

  private evaluateFormulaJs(cellIndex: number, formula: RuntimeFormula): void {
    const sheetName = this.workbook.getSheetNameById(this.workbook.cellStore.sheetIds[cellIndex]!);
    const value = evaluatePlan(formula.compiled.jsPlan, {
      sheetName,
      resolveCell: (targetSheetName, address) => {
        const targetCell = this.workbook.getCellIndex(targetSheetName, address);
        if (targetCell === undefined) {
          return targetSheetName === sheetName ? emptyValue() : errorValue(ErrorCode.Ref);
        }
        return this.workbook.cellStore.getValue(targetCell, (id) => this.strings.get(id));
      },
      resolveRange: (targetSheetName, start, end, _refKind) => {
        if (targetSheetName && !this.workbook.getSheet(targetSheetName)) {
          return [errorValue(ErrorCode.Ref)];
        }
        return this.resolveRangeValues(targetSheetName, parseRangeAddress(`${start}:${end}`, targetSheetName));
      }
    });
    this.workbook.cellStore.setValue(
      cellIndex,
      value,
      value.tag === ValueTag.String ? this.strings.intern(value.value) : 0
    );
  }

  private syncWasmPrograms(): void {
    this.programArena.reset();
    this.constantArena.reset();
    this.rangeListArena.reset();

    let wasmFormulaCount = 0;
    this.formulas.forEach((formula) => {
      if (formula.compiled.mode === FormulaMode.WasmFastPath) {
        wasmFormulaCount += 1;
      }
    });
    this.ensureWasmProgramScratchCapacity(Math.max(wasmFormulaCount, 1), Math.max(this.ranges.size, 1));

    let formulaIndex = 0;
    this.formulas.forEach((formula) => {
      if (formula.compiled.mode !== FormulaMode.WasmFastPath) {
        return;
      }
      const programSlice = this.programArena.append(formula.runtimeProgram);
      const constantSlice = this.constantArena.append(formula.constants);
      const rangeSlice = this.rangeListArena.append(formula.rangeDependencies);

      formula.programOffset = programSlice.offset;
      formula.programLength = programSlice.length;
      formula.constNumberOffset = constantSlice.offset;
      formula.constNumberLength = constantSlice.length;
      formula.rangeListOffset = rangeSlice.offset;
      formula.rangeListLength = rangeSlice.length;
      formula.compiled.programOffset = programSlice.offset;
      formula.compiled.programLength = programSlice.length;
      formula.compiled.constNumberOffset = constantSlice.offset;
      formula.compiled.constNumberLength = constantSlice.length;
      formula.compiled.rangeListOffset = rangeSlice.offset;
      formula.compiled.rangeListLength = rangeSlice.length;
      formula.compiled.depsPtr = formula.dependencyEntities.ptr;
      formula.compiled.depsLen = formula.dependencyEntities.len;

      this.wasmProgramTargets[formulaIndex] = formula.cellIndex;
      this.wasmProgramOffsets[formulaIndex] = programSlice.offset;
      this.wasmProgramLengths[formulaIndex] = programSlice.length;
      this.wasmConstantOffsets[formulaIndex] = constantSlice.offset;
      this.wasmConstantLengths[formulaIndex] = constantSlice.length;
      formulaIndex += 1;
    });

    this.wasm.uploadFormulas({
      targets: this.wasmProgramTargets.subarray(0, wasmFormulaCount),
      programs: this.programArena.view(),
      programOffsets: this.wasmProgramOffsets.subarray(0, wasmFormulaCount),
      programLengths: this.wasmProgramLengths.subarray(0, wasmFormulaCount),
      constants: this.constantArena.view(),
      constantOffsets: this.wasmConstantOffsets.subarray(0, wasmFormulaCount),
      constantLengths: this.wasmConstantLengths.subarray(0, wasmFormulaCount)
    });

    const rangeCapacity = Math.max(this.ranges.size, 1);
    if (this.ranges.size === 0) {
      this.wasmRangeOffsets[0] = 0;
      this.wasmRangeLengths[0] = 0;
    }
    for (let rangeIndex = 0; rangeIndex < this.ranges.size; rangeIndex += 1) {
      const descriptor = this.ranges.getDescriptor(rangeIndex);
      this.wasmRangeOffsets[rangeIndex] = descriptor.refCount > 0 ? descriptor.membersOffset : 0;
      this.wasmRangeLengths[rangeIndex] = descriptor.refCount > 0 ? descriptor.membersLength : 0;
    }

    this.wasm.uploadRanges({
      members: this.ranges.getMemberPoolView(),
      offsets: this.wasmRangeOffsets.subarray(0, rangeCapacity),
      lengths: this.wasmRangeLengths.subarray(0, rangeCapacity)
    });
  }

  private scheduleWasmProgramSync(): void {
    if (this.batchMutationDepth > 0) {
      this.wasmProgramSyncPending = true;
      return;
    }
    this.syncWasmPrograms();
  }

  private flushWasmProgramSync(): void {
    if (!this.wasmProgramSyncPending) {
      return;
    }
    this.wasmProgramSyncPending = false;
    this.syncWasmPrograms();
  }

  private emitBatch(batch: EngineOpBatch): void {
    this.batchListeners.forEach((listener) => listener(batch));
  }

  private estimatePotentialNewCells(ops: readonly EngineOp[]): number {
    let count = 0;
    for (let index = 0; index < ops.length; index += 1) {
      const op = ops[index]!;
      if (
        op.kind === "setCellValue" ||
        op.kind === "setCellFormula" ||
        op.kind === "setCellFormat"
      ) {
        count += 1;
      }
    }
    return count;
  }

  private getEntityDependents(entityId: number): Uint32Array {
    const slice = this.getReverseEdgeSlice(entityId) ?? this.edgeArena.empty();
    return this.edgeArena.readView(slice);
  }

  private setReverseEdgeSlice(entityId: number, slice: EdgeSlice): void {
    const empty = slice.ptr < 0 || slice.len === 0;
    if (isRangeEntity(entityId)) {
      this.reverseRangeEdges[entityPayload(entityId)] = empty ? undefined : slice;
      return;
    }
    this.reverseCellEdges[entityPayload(entityId)] = empty ? undefined : slice;
  }

  private appendReverseEdge(entityId: number, dependentEntityId: number): void {
    const slice = this.getReverseEdgeSlice(entityId) ?? this.edgeArena.empty();
    this.setReverseEdgeSlice(entityId, this.edgeArena.appendUnique(slice, dependentEntityId));
  }

  private removeReverseEdge(entityId: number, dependentEntityId: number): void {
    const slice = this.getReverseEdgeSlice(entityId);
    if (!slice) {
      return;
    }
    this.setReverseEdgeSlice(entityId, this.edgeArena.removeValue(slice, dependentEntityId));
  }

  private getReverseEdgeSlice(entityId: number): EdgeSlice | undefined {
    if (isRangeEntity(entityId)) {
      return this.reverseRangeEdges[entityPayload(entityId)];
    }
    return this.reverseCellEdges[entityPayload(entityId)];
  }

  private collectFormulaDependentsForEntityInto(entityId: number): number {
    this.topoFormulaSeenEpoch += 1;
    this.topoRangeSeenEpoch += 1;

    let entityQueueLength = 1;
    let formulaCount = 0;
    this.topoEntityQueue[0] = entityId;

    for (let queueIndex = 0; queueIndex < entityQueueLength; queueIndex += 1) {
      const currentEntity = this.topoEntityQueue[queueIndex]!;
      const dependents = this.getEntityDependents(currentEntity);
      for (let index = 0; index < dependents.length; index += 1) {
        const dependent = dependents[index]!;
        if (isRangeEntity(dependent)) {
          const rangeIndex = entityPayload(dependent);
          if (this.topoRangeSeen[rangeIndex] === this.topoRangeSeenEpoch) {
            continue;
          }
          this.topoRangeSeen[rangeIndex] = this.topoRangeSeenEpoch;
          this.topoEntityQueue[entityQueueLength] = dependent;
          entityQueueLength += 1;
          continue;
        }

        const formulaCellIndex = entityPayload(dependent);
        if (this.topoFormulaSeen[formulaCellIndex] === this.topoFormulaSeenEpoch) {
          continue;
        }
        this.topoFormulaSeen[formulaCellIndex] = this.topoFormulaSeenEpoch;
        this.topoFormulaBuffer[formulaCount] = formulaCellIndex;
        formulaCount += 1;
      }
    }

    return formulaCount;
  }

  private forEachFormulaDependencyCell(cellIndex: number, fn: (dependencyCellIndex: number) => void): void {
    const formula = this.formulas.get(cellIndex);
    if (!formula) {
      return;
    }
    for (let index = 0; index < formula.dependencyIndices.length; index += 1) {
      fn(formula.dependencyIndices[index]!);
    }
  }

  private shouldApplyOp(op: EngineOp, order: OpOrder): boolean {
    const sheetDeleteOrder = this.sheetDeleteBarrierForOp(op);
    if (sheetDeleteOrder && compareOpOrder(order, sheetDeleteOrder) <= 0) {
      return false;
    }
    const existingOrder = this.entityVersions.get(this.entityKeyForOp(op));
    if (existingOrder && compareOpOrder(order, existingOrder) <= 0) {
      return false;
    }
    return true;
  }

  private entityKeyForOp(op: EngineOp): string {
    switch (op.kind) {
      case "upsertWorkbook":
        return "workbook";
      case "upsertSheet":
      case "deleteSheet":
        return `sheet:${op.name}`;
      case "setCellFormat":
        return `format:${op.sheetName}!${op.address}`;
      case "setCellValue":
      case "setCellFormula":
      case "clearCell":
        return `cell:${op.sheetName}!${op.address}`;
    }
  }

  private sheetDeleteBarrierForOp(op: EngineOp): OpOrder | undefined {
    switch (op.kind) {
      case "setCellFormat":
        return this.sheetDeleteVersions.get(op.sheetName);
      case "setCellValue":
      case "setCellFormula":
      case "clearCell":
        return this.sheetDeleteVersions.get(op.sheetName);
      case "upsertSheet":
        return this.sheetDeleteVersions.get(op.name);
      default:
        return undefined;
    }
  }

  private removeSheetRuntime(
    sheetName: string,
    explicitChangedCount: number
  ): { changedInputCount: number; formulaChangedCount: number; explicitChangedCount: number } {
    const sheet = this.workbook.getSheet(sheetName);
    if (!sheet) {
      return { changedInputCount: 0, formulaChangedCount: 0, explicitChangedCount };
    }

    const cellIndices: number[] = [];
    sheet.grid.forEachCell((cellIndex) => {
      cellIndices.push(cellIndex);
    });
    const impactedCount = this.collectImpactedFormulasForCells(cellIndices);

    let changedInputCount = 0;
    let formulaChangedCount = 0;
    cellIndices.forEach((cellIndex) => {
      this.removeFormula(cellIndex);
      this.setReverseEdgeSlice(makeCellEntity(cellIndex), this.edgeArena.empty());
      this.workbook.cellStore.setValue(cellIndex, emptyValue());
      this.workbook.cellStore.flags[cellIndex] =
        (this.workbook.cellStore.flags[cellIndex] ?? 0) | CellFlags.PendingDelete;
      changedInputCount = this.markInputChanged(cellIndex, changedInputCount);
      explicitChangedCount = this.markExplicitChanged(cellIndex, explicitChangedCount);
    });

    this.workbook.deleteSheet(sheetName);
    if (this.selection.sheetName === sheetName) {
      const nextSheet = [...this.workbook.sheetsByName.values()].sort((left, right) => left.order - right.order)[0];
      this.setSelection(nextSheet?.name ?? sheetName, "A1");
    }
    formulaChangedCount = this.rebindFormulasForSheet(
      sheetName,
      formulaChangedCount,
      this.impactedFormulaBuffer.subarray(0, impactedCount)
    );
    return { changedInputCount, formulaChangedCount, explicitChangedCount };
  }

  private rebindFormulasForSheet(
    sheetName: string,
    formulaChangedCount: number,
    candidates?: readonly number[] | U32
  ): number {
    if (candidates) {
      for (let index = 0; index < candidates.length; index += 1) {
        const cellIndex = candidates[index]!;
        const formula = this.formulas.get(cellIndex);
        if (!formula) continue;
        const ownerSheetName = this.workbook.getSheetNameById(this.workbook.cellStore.sheetIds[cellIndex]!);
        if (!ownerSheetName) continue;
        const touchesSheet = formula.compiled.deps.some((dep) => {
          if (!dep.includes("!")) return false;
          const [qualifiedSheet] = dep.split("!");
          return qualifiedSheet?.replace(/^'(.*)'$/, "$1") === sheetName;
        });
        if (!touchesSheet) continue;
        const dependencies = this.materializeDependencies(ownerSheetName, formula.compiled);
        this.setFormula(cellIndex, formula.source, formula.compiled, dependencies);
        formulaChangedCount = this.markFormulaChanged(cellIndex, formulaChangedCount);
      }
      return formulaChangedCount;
    }

    this.formulas.forEach((formula, cellIndex) => {
      if (!formula) return;
      const ownerSheetName = this.workbook.getSheetNameById(this.workbook.cellStore.sheetIds[cellIndex]!);
      if (!ownerSheetName) return;
      const touchesSheet = formula.compiled.deps.some((dep) => {
        if (!dep.includes("!")) return false;
        const [qualifiedSheet] = dep.split("!");
        return qualifiedSheet?.replace(/^'(.*)'$/, "$1") === sheetName;
      });
      if (!touchesSheet) return;
      const dependencies = this.materializeDependencies(ownerSheetName, formula.compiled);
      this.setFormula(cellIndex, formula.source, formula.compiled, dependencies);
      formulaChangedCount = this.markFormulaChanged(cellIndex, formulaChangedCount);
    });

    return formulaChangedCount;
  }

  private collectImpactedFormulasForCells(cellIndices: readonly number[]): number {
    this.ensureRecalcScratchCapacity(this.workbook.cellStore.size + 1);
    this.impactedFormulaEpoch += 1;
    if (this.impactedFormulaEpoch === 0xffff_ffff) {
      this.impactedFormulaEpoch = 1;
      this.impactedFormulaSeen.fill(0);
    }

    let impactedCount = 0;
    for (let cellCursor = 0; cellCursor < cellIndices.length; cellCursor += 1) {
      const cellIndex = cellIndices[cellCursor]!;
      const dependentCount = this.collectFormulaDependentsForEntityInto(makeCellEntity(cellIndex));
      for (let dependentIndex = 0; dependentIndex < dependentCount; dependentIndex += 1) {
        const formulaCellIndex = this.topoFormulaBuffer[dependentIndex]!;
        if (this.impactedFormulaSeen[formulaCellIndex] === this.impactedFormulaEpoch) {
          continue;
        }
        this.impactedFormulaSeen[formulaCellIndex] = this.impactedFormulaEpoch;
        this.impactedFormulaBuffer[impactedCount] = formulaCellIndex;
        impactedCount += 1;
      }
    }

    return impactedCount;
  }

  private beginMutationCollection(): void {
    this.changedInputEpoch += 1;
    if (this.changedInputEpoch === 0xffff_ffff) {
      this.changedInputEpoch = 1;
      this.changedInputSeen.fill(0);
    }
    this.changedFormulaEpoch += 1;
    if (this.changedFormulaEpoch === 0xffff_ffff) {
      this.changedFormulaEpoch = 1;
      this.changedFormulaSeen.fill(0);
    }
    this.explicitChangedEpoch += 1;
    if (this.explicitChangedEpoch === 0xffff_ffff) {
      this.explicitChangedEpoch = 1;
      this.explicitChangedSeen.fill(0);
    }
    this.ensureRecalcScratchCapacity(this.workbook.cellStore.size + 1);
  }

  private markInputChanged(cellIndex: number, count: number): number {
    if (this.changedInputSeen[cellIndex] === this.changedInputEpoch) {
      return count;
    }
    this.changedInputSeen[cellIndex] = this.changedInputEpoch;
    this.changedInputBuffer[count] = cellIndex;
    return count + 1;
  }

  private markFormulaChanged(cellIndex: number, count: number): number {
    if (this.changedFormulaSeen[cellIndex] === this.changedFormulaEpoch) {
      return count;
    }
    this.changedFormulaSeen[cellIndex] = this.changedFormulaEpoch;
    this.changedFormulaBuffer[count] = cellIndex;
    return count + 1;
  }

  private markExplicitChanged(cellIndex: number, count: number): number {
    if (this.explicitChangedSeen[cellIndex] === this.explicitChangedEpoch) {
      return count;
    }
    this.explicitChangedSeen[cellIndex] = this.explicitChangedEpoch;
    this.explicitChangedBuffer[count] = cellIndex;
    return count + 1;
  }

  private composeMutationRoots(changedInputCount: number, formulaChangedCount: number): U32 {
    const total = changedInputCount + formulaChangedCount;
    this.ensureRecalcScratchCapacity(total + 1);
    for (let index = 0; index < changedInputCount; index += 1) {
      this.mutationRoots[index] = this.changedInputBuffer[index]!;
    }
    for (let index = 0; index < formulaChangedCount; index += 1) {
      this.mutationRoots[changedInputCount + index] = this.changedFormulaBuffer[index]!;
    }
    return this.mutationRoots.subarray(0, total);
  }

  private composeEventChanges(recalculated: U32, explicitChangedCount: number): U32 {
    this.changedUnionEpoch += 1;
    if (this.changedUnionEpoch === 0xffff_ffff) {
      this.changedUnionEpoch = 1;
      this.changedUnionSeen.fill(0);
    }
    let changedCount = 0;

    for (let index = 0; index < explicitChangedCount; index += 1) {
      const cellIndex = this.explicitChangedBuffer[index]!;
      if (this.changedUnionSeen[cellIndex] === this.changedUnionEpoch) {
        continue;
      }
      this.changedUnionSeen[cellIndex] = this.changedUnionEpoch;
      this.changedUnion[changedCount] = cellIndex;
      changedCount += 1;
    }

    for (let index = 0; index < recalculated.length; index += 1) {
      const cellIndex = recalculated[index]!;
      if (this.changedUnionSeen[cellIndex] === this.changedUnionEpoch) {
        continue;
      }
      this.changedUnionSeen[cellIndex] = this.changedUnionEpoch;
      this.changedUnion[changedCount] = cellIndex;
      changedCount += 1;
    }

    return this.changedUnion.subarray(0, changedCount);
  }

  private composeChangedRootsAndOrdered(changedRoots: readonly number[] | U32, ordered: U32, orderedCount: number): U32 {
    this.changedUnionEpoch += 1;
    if (this.changedUnionEpoch === 0xffff_ffff) {
      this.changedUnionEpoch = 1;
      this.changedUnionSeen.fill(0);
    }
    let changedCount = 0;

    for (let index = 0; index < changedRoots.length; index += 1) {
      const cellIndex = changedRoots[index]!;
      if (this.changedUnionSeen[cellIndex] === this.changedUnionEpoch) {
        continue;
      }
      this.changedUnionSeen[cellIndex] = this.changedUnionEpoch;
      this.changedUnion[changedCount] = cellIndex;
      changedCount += 1;
    }
    for (let index = 0; index < orderedCount; index += 1) {
      const cellIndex = ordered[index]!;
      if (this.changedUnionSeen[cellIndex] === this.changedUnionEpoch) {
        continue;
      }
      this.changedUnionSeen[cellIndex] = this.changedUnionEpoch;
      this.changedUnion[changedCount] = cellIndex;
      changedCount += 1;
    }

    return this.changedUnion.subarray(0, changedCount);
  }

  private ensureCellTracked(sheetName: string, address: string): number {
    const ensured = this.workbook.ensureCellRecord(sheetName, address);
    if (ensured.created) {
      this.pushMaterializedCell(ensured.cellIndex);
    }
    return ensured.cellIndex;
  }

  private ensureCellTrackedByCoords(sheetId: number, row: number, col: number): number {
    const ensured = this.workbook.ensureCellAt(sheetId, row, col);
    if (ensured.created) {
      this.pushMaterializedCell(ensured.cellIndex);
    }
    return ensured.cellIndex;
  }

  private forEachSheetCell(sheetId: number, fn: (cellIndex: number, row: number, col: number) => void): void {
    const sheet = this.workbook.getSheetById(sheetId);
    if (!sheet) {
      return;
    }
    sheet.grid.forEachCellEntry((cellIndex, row, col) => {
      fn(cellIndex, row, col);
    });
  }

  private resolveRangeValues(sheetName: string, range: ReturnType<typeof parseRangeAddress>): CellValue[] {
    if (range.kind === "cells") {
      const values: CellValue[] = [];
      for (let row = range.start.row; row <= range.end.row; row += 1) {
        for (let col = range.start.col; col <= range.end.col; col += 1) {
          const addr = formatAddress(row, col);
          const index = this.workbook.getCellIndex(sheetName, addr);
          values.push(index === undefined ? emptyValue() : this.workbook.cellStore.getValue(index, (id) => this.strings.get(id)));
        }
      }
      return values;
    }

    const sheet = this.workbook.getSheet(sheetName);
    if (!sheet) {
      return [errorValue(ErrorCode.Ref)];
    }

    const matches: Array<{ row: number; col: number; value: CellValue }> = [];
    sheet.grid.forEachCellEntry((cellIndex, row, col) => {
      const inRange =
        range.kind === "rows"
          ? row >= range.start.row && row <= range.end.row
          : col >= range.start.col && col <= range.end.col;
      if (!inRange) {
        return;
      }
      matches.push({
        row,
        col,
        value: this.workbook.cellStore.getValue(cellIndex, (id) => this.strings.get(id))
      });
    });
    matches.sort((left, right) => left.row - right.row || left.col - right.col);
    return matches.map((match) => match.value);
  }

  private syncDynamicRanges(formulaChangedCount: number): number {
    for (let index = 0; index < this.materializedCellCount; index += 1) {
      const cellIndex = this.materializedCells[index]!;
      const sheetId = this.workbook.cellStore.sheetIds[cellIndex] ?? 0;
      if (sheetId === 0) {
        continue;
      }
      const row = this.workbook.cellStore.rows[cellIndex] ?? 0;
      const col = this.workbook.cellStore.cols[cellIndex] ?? 0;
      const rangeIndices = this.ranges.addDynamicMember(sheetId, row, col, cellIndex);
      if (rangeIndices.length > 0) {
        this.scheduleWasmProgramSync();
      }
      for (let rangeCursor = 0; rangeCursor < rangeIndices.length; rangeCursor += 1) {
        const rangeIndex = rangeIndices[rangeCursor]!;
        const rangeEntity = makeRangeEntity(rangeIndex);
        this.appendReverseEdge(makeCellEntity(cellIndex), rangeEntity);
        const formulas = this.getEntityDependents(rangeEntity);
        for (let formulaCursor = 0; formulaCursor < formulas.length; formulaCursor += 1) {
          const formulaEntity = formulas[formulaCursor]!;
          if (isRangeEntity(formulaEntity)) {
            continue;
          }
          const formulaCellIndex = entityPayload(formulaEntity);
          const formula = this.formulas.get(formulaCellIndex);
          if (!formula) {
            continue;
          }
          const nextDependencyIndices = appendPackedCellIndex(formula.dependencyIndices, cellIndex);
          if (nextDependencyIndices !== formula.dependencyIndices) {
            formula.dependencyIndices = nextDependencyIndices;
            formulaChangedCount = this.markFormulaChanged(formulaCellIndex, formulaChangedCount);
          }
        }
      }
    }
    return formulaChangedCount;
  }

  private resetWorkbook(workbookName = "Workbook"): void {
    const previousBatchId = this.lastMetrics.batchId;
    this.workbook.reset(workbookName);
    this.formulas.clear();
    this.reverseCellEdges = [];
    this.reverseRangeEdges = [];
    this.ranges.reset();
    this.edgeArena.reset();
    this.entityVersions.clear();
    this.sheetDeleteVersions.clear();
    this.selection = { sheetName: "Sheet1", address: "A1" };
    this.lastMetrics = {
      batchId: previousBatchId,
      changedInputCount: 0,
      dirtyFormulaCount: 0,
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
      rangeNodeVisits: 0,
      recalcMs: 0,
      compileMs: 0
    };
    this.wasmProgramSyncPending = false;
    this.materializedCellCount = 0;
    this.syncWasmPrograms();
  }

  private resetMaterializedCellScratch(expectedSize: number): void {
    this.materializedCellCount = 0;
    if (expectedSize > this.materializedCells.length) {
      this.materializedCells = growUint32(this.materializedCells, expectedSize);
    }
  }

  private pushMaterializedCell(cellIndex: number): void {
    const nextCount = this.materializedCellCount + 1;
    if (nextCount > this.materializedCells.length) {
      this.materializedCells = growUint32(this.materializedCells, nextCount);
    }
    this.materializedCells[this.materializedCellCount] = cellIndex;
    this.materializedCellCount = nextCount;
  }
}

function growUint32(buffer: U32, required: number): U32 {
  let capacity = buffer.length;
  while (capacity < required) {
    capacity *= 2;
  }
  const next = new Uint32Array(capacity);
  next.set(buffer);
  return next as U32;
}

function appendPackedCellIndex(indices: Uint32Array, cellIndex: number): Uint32Array {
  for (let index = 0; index < indices.length; index += 1) {
    if (indices[index] === cellIndex) {
      return indices;
    }
  }
  const next = new Uint32Array(indices.length + 1);
  next.set(indices);
  next[indices.length] = cellIndex;
  return next;
}

export const selectors = {
  selectCellSnapshot,
  selectMetrics,
  selectSelectionState,
  selectViewportCells
};
