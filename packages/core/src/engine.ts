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
  type WorkbookSnapshot
} from "@bilig/protocol";
import { compileFormula, evaluatePlan, formatAddress, parseCellAddress, parseRangeAddress } from "@bilig/formula";
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
import { detectFormulaCycles } from "./cycle-detection.js";
import { EdgeArena, type EdgeSlice } from "./edge-arena.js";
import { entityPayload, isRangeEntity, makeCellEntity, makeRangeEntity } from "./entity-ids.js";
import { EngineEventBus } from "./events.js";
import { RangeRegistry } from "./range-registry.js";
import { RecalcScheduler } from "./scheduler.js";
import { selectCellSnapshot, selectViewportCells } from "./selectors.js";
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
  dependencyIndices: number[];
  dependencyEntities: EdgeSlice;
  rangeDependencies: Uint32Array;
  runtimeProgram: Uint32Array;
  constants: number[];
}

interface MaterializedCell {
  sheetName: string;
  address: string;
  cellIndex: number;
}

interface MaterializedDependencies {
  dependencyIndices: number[];
  dependencyEntities: Uint32Array;
  rangeDependencies: Uint32Array;
  rangeIndexByRef: Map<string, number>;
  newRangeLinks: Array<{ rangeIndex: number; memberIndices: Uint32Array }>;
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

  private readonly formulas = new Map<number, RuntimeFormula>();
  private readonly edgeArena = new EdgeArena();
  private readonly reverseEdges = new Map<number, EdgeSlice>();
  private readonly batchListeners = new Set<(batch: EngineOpBatch) => void>();
  private readonly entityVersions = new Map<string, OpOrder>();
  private readonly sheetDeleteVersions = new Map<string, OpOrder>();
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
    return this.events.subscribeCell(`${sheetName}!${address}`, listener);
  }

  subscribeCells(sheetName: string, addresses: readonly string[], listener: () => void): () => void {
    return this.events.subscribeCells(addresses.map((address) => `${sheetName}!${address}`), listener);
  }

  subscribeBatches(listener: (batch: EngineOpBatch) => void): () => void {
    this.batchListeners.add(listener);
    return () => {
      this.batchListeners.delete(listener);
    };
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
        flags: 0
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
      flags: this.workbook.cellStore.flags[cellIndex]!
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
    const directPrecedents = this.getFormulaDependencyCells(cellIndex);
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
      directPrecedents: directPrecedents.map((index) =>
        this.workbook.getQualifiedAddress(index)
      ),
      directDependents: [...directDependents].map((index) =>
        this.workbook.getQualifiedAddress(index)
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

    const changedInputs = new Set<number>();
    const formulaChanged = new Set<number>();
    const materializedCells: MaterializedCell[] = [];
    let compileMs = 0;

    this.batchMutationDepth += 1;
    try {
      snapshot.sheets.forEach((sheet) => {
        sheet.cells.forEach((cell) => {
          const cellIndex = this.ensureCellTracked(sheet.name, cell.address, materializedCells);
          if (cell.formula !== undefined) {
            const compileStarted = performance.now();
            const compiled = compileFormula(cell.formula);
            compileMs += performance.now() - compileStarted;
            const dependencies = this.materializeDependencies(sheet.name, compiled.deps, materializedCells);
            this.setFormula(cellIndex, cell.formula, compiled, dependencies, materializedCells);
            formulaChanged.add(cellIndex);
          } else {
            const value = literalToValue(cell.value ?? null, this.strings);
            this.workbook.cellStore.setValue(
              cellIndex,
              value,
              value.tag === ValueTag.String ? value.stringId : 0
            );
            this.workbook.cellStore.flags[cellIndex] =
              (this.workbook.cellStore.flags[cellIndex] ?? 0) & ~CellFlags.HasFormula;
            changedInputs.add(cellIndex);
          }
          if (cell.format !== undefined) {
            this.workbook.setCellFormat(cellIndex, cell.format);
          }
        });
      });

      this.syncDynamicRanges(materializedCells).forEach((cellIndex) => {
        formulaChanged.add(cellIndex);
      });
    } finally {
      this.batchMutationDepth -= 1;
      this.flushWasmProgramSync();
    }

    this.lastMetrics.compileMs = compileMs;
    if (formulaChanged.size > 0) {
      this.rebuildTopoRanks();
      this.detectCycles();
    }

    const changedInputArray = [...changedInputs];
    const changed = this.recalculate([...changedInputArray, ...formulaChanged], changedInputArray);
    this.lastMetrics.batchId += 1;
    this.lastMetrics.changedInputCount = changedInputs.size + formulaChanged.size;

    const event: EngineEvent = {
      kind: "batch",
      changedCellIndices: changed,
      metrics: this.lastMetrics
    };
    if (this.events.hasCellListeners()) {
      this.events.emitAllWatched(event);
      return;
    }
    this.events.emit(event);
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
    const changedInputs = new Set<number>();
    const formulaChanged = new Set<number>();
    const trackQualifiedAddresses = this.events.hasCellListeners();
    const changedQualifiedAddresses = trackQualifiedAddresses ? new Set<string>() : null;
    const materializedCells: MaterializedCell[] = [];
    let topologyChanged = false;
    let appliedOps = 0;
    const canSkipOrderChecks = source === "local";

    this.workbook.cellStore.ensureCapacity(
      this.workbook.cellStore.size + (potentialNewCells ?? this.estimatePotentialNewCells(batch.ops))
    );

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
            this.rebindFormulasForSheet(op.name).forEach((cellIndex) => {
              formulaChanged.add(cellIndex);
              topologyChanged = true;
            });
            break;
          case "deleteSheet":
            this.removeSheetRuntime(op.name, changedInputs, formulaChanged, changedQualifiedAddresses);
            this.entityVersions.set(this.entityKeyForOp(op), order);
            this.sheetDeleteVersions.set(op.name, order);
            topologyChanged = true;
            break;
          case "setCellValue": {
            const cellIndex = this.ensureCellTracked(op.sheetName, op.address, materializedCells);
            topologyChanged = this.removeFormula(cellIndex) || topologyChanged;
            const value = literalToValue(op.value, this.strings);
            this.workbook.cellStore.setValue(
              cellIndex,
              value,
              value.tag === ValueTag.String ? value.stringId : 0
            );
            this.workbook.cellStore.flags[cellIndex] =
              (this.workbook.cellStore.flags[cellIndex] ?? 0) & ~CellFlags.HasFormula;
            changedInputs.add(cellIndex);
            changedQualifiedAddresses?.add(`${op.sheetName}!${op.address}`);
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          }
          case "setCellFormula": {
            const cellIndex = this.ensureCellTracked(op.sheetName, op.address, materializedCells);
            const compileStarted = performance.now();
            const compiled = compileFormula(op.formula);
            this.lastMetrics.compileMs = performance.now() - compileStarted;
            const dependencies = this.materializeDependencies(op.sheetName, compiled.deps, materializedCells);
            this.setFormula(cellIndex, op.formula, compiled, dependencies, materializedCells);
            formulaChanged.add(cellIndex);
            changedQualifiedAddresses?.add(`${op.sheetName}!${op.address}`);
            this.entityVersions.set(this.entityKeyForOp(op), order);
            topologyChanged = true;
            break;
          }
          case "setCellFormat": {
            const cellIndex = this.ensureCellTracked(op.sheetName, op.address, materializedCells);
            this.workbook.setCellFormat(cellIndex, op.format);
            changedQualifiedAddresses?.add(`${op.sheetName}!${op.address}`);
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
            changedInputs.add(cellIndex);
            changedQualifiedAddresses?.add(`${op.sheetName}!${op.address}`);
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          }
        }
        appliedOps += 1;
      });

      this.syncDynamicRanges(materializedCells).forEach((cellIndex) => {
        formulaChanged.add(cellIndex);
        topologyChanged = true;
      });
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
    const changedInputArray = [...changedInputs];
    const changed = this.recalculate([...changedInputArray, ...formulaChanged], changedInputArray);
    this.lastMetrics.batchId += 1;
    this.lastMetrics.changedInputCount = changedInputs.size + formulaChanged.size;
    if (changedQualifiedAddresses) {
      changed.forEach((cellIndex) => {
        const qualifiedAddress = this.workbook.getQualifiedAddress(cellIndex);
        if (!qualifiedAddress.startsWith("!")) {
          changedQualifiedAddresses.add(qualifiedAddress);
        }
      });
    }
    this.events.emit({
      kind: "batch",
      changedCellIndices: changed,
      metrics: this.lastMetrics
    }, changedQualifiedAddresses ? [...changedQualifiedAddresses] : []);
    if (source === "local") {
      this.emitBatch(batch);
    }
  }

  private materializeDependencies(currentSheet: string, deps: string[], materializedCells: MaterializedCell[]): MaterializedDependencies {
    const indices = new Set<number>();
    const dependencyEntities: number[] = [];
    const rangeDependencies: number[] = [];
    const rangeIndexByRef = new Map<string, number>();
    const newRangeLinks: Array<{ rangeIndex: number; memberIndices: Uint32Array }> = [];
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
          ensureCell: (sheetId, row, col) => this.ensureCellTrackedByCoords(sheetId, row, col, materializedCells),
          listSheetCells: (sheetId) => this.listSheetCells(sheetId)
        });
        rangeIndexByRef.set(dep, registered.rangeIndex);
        const rangeEntity = makeRangeEntity(registered.rangeIndex);
        dependencyEntities.push(rangeEntity);
        rangeDependencies.push(registered.rangeIndex);
        const memberIndices = this.ranges.expandToCells(registered.rangeIndex);
        for (let memberIndex = 0; memberIndex < memberIndices.length; memberIndex += 1) {
          indices.add(memberIndices[memberIndex]!);
        }
        if (registered.materialized) {
          newRangeLinks.push({ rangeIndex: registered.rangeIndex, memberIndices });
        }
        continue;
      }
      const parsed = parseCellAddress(dep, currentSheet);
      const sheetName = parsed.sheetName ?? currentSheet;
      if (parsed.sheetName && !this.workbook.getSheet(sheetName)) {
        continue;
      }
      const cellIndex = this.ensureCellTracked(sheetName, parsed.text, materializedCells);
      indices.add(cellIndex);
      dependencyEntities.push(makeCellEntity(cellIndex));
    }
    return {
      dependencyIndices: [...indices],
      dependencyEntities: Uint32Array.from(dependencyEntities),
      rangeDependencies: Uint32Array.from(rangeDependencies),
      rangeIndexByRef,
      newRangeLinks
    };
  }

  private setFormula(
    cellIndex: number,
    source: string,
    compiled: ReturnType<typeof compileFormula>,
    dependencies: MaterializedDependencies,
    materializedCells: MaterializedCell[]
  ): void {
    this.removeFormula(cellIndex);

    const symbolicRefToIndex = new Map<string, number>();
    let hasUnresolvedSymbolicRef = false;
    compiled.symbolicRefs.forEach((ref) => {
      const [qualifiedSheetName, qualifiedAddress] = ref.includes("!") ? ref.split("!") : [undefined, ref];
      const parsed = parseCellAddress(qualifiedAddress!, qualifiedSheetName);
      const sheetName = qualifiedSheetName ?? this.workbook.getSheetNameById(this.workbook.cellStore.sheetIds[cellIndex]!);
      if (qualifiedSheetName && !this.workbook.getSheet(sheetName)) {
        hasUnresolvedSymbolicRef = true;
        return;
      }
      symbolicRefToIndex.set(ref, this.ensureCellTracked(sheetName, parsed.text, materializedCells));
    });

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
        const cellRef = compiled.symbolicRefs[operand];
        const targetIndex = cellRef ? symbolicRefToIndex.get(cellRef) ?? 0 : 0;
        runtimeProgram[index] = (opcode << 24) | (targetIndex & 0x00ff_ffff);
        return;
      }
      if (opcode === 4) {
        const rangeRef = compiled.symbolicRanges[operand];
        const targetIndex = rangeRef ? dependencies.rangeIndexByRef.get(rangeRef) ?? 0 : 0;
        runtimeProgram[index] = (opcode << 24) | (targetIndex & 0x00ff_ffff);
      }
    });

    this.formulas.set(cellIndex, {
      cellIndex,
      source,
      compiled: effectiveCompiled,
      dependencyIndices: dependencies.dependencyIndices,
      dependencyEntities: this.edgeArena.replace(this.edgeArena.empty(), dependencies.dependencyEntities),
      rangeDependencies: dependencies.rangeDependencies,
      runtimeProgram,
      constants: effectiveCompiled.constants
    });
    this.workbook.cellStore.flags[cellIndex] = (this.workbook.cellStore.flags[cellIndex] ?? 0) | CellFlags.HasFormula;
    this.workbook.cellStore.formulaIds[cellIndex] = cellIndex + 1;
    if (effectiveCompiled.mode === FormulaMode.JsOnly) {
      this.workbook.cellStore.flags[cellIndex] = (this.workbook.cellStore.flags[cellIndex] ?? 0) | CellFlags.JsOnly;
    } else {
      this.workbook.cellStore.flags[cellIndex] = (this.workbook.cellStore.flags[cellIndex] ?? 0) & ~CellFlags.JsOnly;
    }

    dependencies.newRangeLinks.forEach(({ rangeIndex, memberIndices }) => {
      const rangeEntity = makeRangeEntity(rangeIndex);
      for (let index = 0; index < memberIndices.length; index += 1) {
        this.appendReverseEdge(makeCellEntity(memberIndices[index]!), rangeEntity);
      }
    });
    const formulaEntity = makeCellEntity(cellIndex);
    for (let index = 0; index < dependencies.dependencyEntities.length; index += 1) {
      this.appendReverseEdge(dependencies.dependencyEntities[index]!, formulaEntity);
    }
    this.scheduleWasmProgramSync();
  }

  private removeFormula(cellIndex: number): boolean {
    const existing = this.formulas.get(cellIndex);
    if (existing) {
      const dependencyEntities = this.edgeArena.read(existing.dependencyEntities);
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
        this.reverseEdges.delete(rangeEntity);
      }
      this.edgeArena.free(existing.dependencyEntities);
    }
    this.formulas.delete(cellIndex);
    this.workbook.cellStore.flags[cellIndex] =
      (this.workbook.cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle);
    this.workbook.cellStore.formulaIds[cellIndex] = 0;
    this.scheduleWasmProgramSync();
    return existing !== undefined;
  }

  private rebuildTopoRanks(): void {
    const indegree = new Map<number, number>();
    this.formulas.forEach((_formula, cellIndex) => indegree.set(cellIndex, 0));
    this.formulas.forEach((formula, cellIndex) => {
      formula.dependencyIndices.forEach((dep) => {
        if (this.formulas.has(dep)) {
          indegree.set(cellIndex, (indegree.get(cellIndex) ?? 0) + 1);
        }
      });
    });

    const queue = [...indegree.entries()].filter(([, count]) => count === 0).map(([cellIndex]) => cellIndex);
    let rank = 0;
    for (let queueIndex = 0; queueIndex < queue.length; queueIndex += 1) {
      const cellIndex = queue[queueIndex]!;
      this.workbook.cellStore.topoRanks[cellIndex] = rank++;
      const dependents = new Set<number>();
      this.collectFormulaDependentsForEntity(makeCellEntity(cellIndex), dependents);
      for (const dependent of dependents) {
        if (!indegree.has(dependent)) {
          continue;
        }
        const next = (indegree.get(dependent) ?? 0) - 1;
        indegree.set(dependent, next);
        if (next === 0) queue.push(dependent);
      }
    }
  }

  private detectCycles(): void {
    const result = detectFormulaCycles(
      this.formulas.keys(),
      (cellIndex) => this.getFormulaDependencyCells(cellIndex).filter((dependency) => this.formulas.has(dependency))
    );

    this.formulas.forEach((_formula, cellIndex) => {
      this.workbook.cellStore.flags[cellIndex] = (this.workbook.cellStore.flags[cellIndex] ?? 0) & ~CellFlags.InCycle;
      this.workbook.cellStore.cycleGroupIds[cellIndex] = -1;
    });

    result.inCycle.forEach((cellIndex) => {
      this.workbook.cellStore.flags[cellIndex] = (this.workbook.cellStore.flags[cellIndex] ?? 0) | CellFlags.InCycle;
      this.workbook.cellStore.cycleGroupIds[cellIndex] = result.cycleGroups.get(cellIndex) ?? -1;
      this.workbook.cellStore.setValue(cellIndex, errorValue(ErrorCode.Cycle));
    });
  }

  private recalculate(changedRoots: number[], kernelSyncRoots: readonly number[] = changedRoots): Uint32Array {
    const started = performance.now();
    const scheduled = this.scheduler.collectDirty(
      changedRoots,
      { getDependents: (entityId) => this.getEntityDependents(entityId) },
      this.workbook.cellStore,
      (cellIndex) => this.formulas.has(cellIndex),
      this.ranges.size
    );
    const ordered = scheduled.orderedFormulaCellIndices;

    const pendingKernelSync = [...kernelSyncRoots];

    const flushWasmBatch = (batch: number[]): number => {
      if (batch.length === 0) {
        return 0;
      }
      this.wasm.syncFromStore(this.workbook.cellStore, pendingKernelSync);
      pendingKernelSync.length = 0;
      const batchIndices = Uint32Array.from(batch);
      this.wasm.evalBatch(batchIndices);
      this.wasm.syncToStore(this.workbook.cellStore, batchIndices);
      return batch.length;
    };

    let wasmCount = 0;
    let jsCount = 0;
    let wasmBatch: number[] = [];
    for (let index = 0; index < ordered.length; index += 1) {
      const cellIndex = ordered[index]!;
      const formula = this.formulas.get(cellIndex);
      if (!formula) {
        continue;
      }
      if (((this.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0) {
        continue;
      }
      if (formula.compiled.mode === FormulaMode.WasmFastPath && this.wasm.ready) {
        wasmBatch.push(cellIndex);
        continue;
      }
      wasmCount += flushWasmBatch(wasmBatch);
      wasmBatch = [];
      jsCount += 1;
      this.evaluateFormulaJs(cellIndex, formula);
      pendingKernelSync.push(cellIndex);
    }

    wasmCount += flushWasmBatch(wasmBatch);
    if (pendingKernelSync.length > 0) {
      this.wasm.syncFromStore(this.workbook.cellStore, pendingKernelSync);
    }

    this.lastMetrics.dirtyFormulaCount = ordered.length;
    this.lastMetrics.jsFormulaCount = jsCount;
    this.lastMetrics.wasmFormulaCount = wasmCount;
    this.lastMetrics.rangeNodeVisits = scheduled.rangeNodeVisits;
    this.lastMetrics.recalcMs = performance.now() - started;

    return Uint32Array.from([...new Set([...changedRoots, ...ordered])]);
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
    this.wasm.uploadFormulas(
      [...this.formulas.values()].map((formula) => ({
        cellIndex: formula.cellIndex,
        program: formula.runtimeProgram,
        constants: formula.constants,
        mode: formula.compiled.mode
      }))
    );
    this.wasm.uploadRanges(
      Array.from({ length: this.ranges.size }, (_, rangeIndex) => {
        const descriptor = this.ranges.getDescriptor(rangeIndex);
        return {
          rangeIndex,
          members: descriptor.refCount > 0 ? this.ranges.getMembers(rangeIndex) : new Uint32Array()
        };
      })
    );
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
    const slice = this.reverseEdges.get(entityId) ?? this.edgeArena.empty();
    return this.edgeArena.read(slice);
  }

  private setReverseEdgeSlice(entityId: number, slice: EdgeSlice): void {
    if (slice.ptr < 0 || slice.len === 0) {
      this.reverseEdges.delete(entityId);
      return;
    }
    this.reverseEdges.set(entityId, slice);
  }

  private appendReverseEdge(entityId: number, dependentEntityId: number): void {
    const slice = this.reverseEdges.get(entityId) ?? this.edgeArena.empty();
    this.setReverseEdgeSlice(entityId, this.edgeArena.appendUnique(slice, dependentEntityId));
  }

  private removeReverseEdge(entityId: number, dependentEntityId: number): void {
    const slice = this.reverseEdges.get(entityId);
    if (!slice) {
      return;
    }
    this.setReverseEdgeSlice(entityId, this.edgeArena.removeValue(slice, dependentEntityId));
  }

  private collectFormulaDependentsForEntity(entityId: number, output: Set<number>): void {
    const dependents = this.getEntityDependents(entityId);
    for (let index = 0; index < dependents.length; index += 1) {
      const dependent = dependents[index]!;
      if (isRangeEntity(dependent)) {
        this.collectFormulaDependentsForEntity(dependent, output);
        continue;
      }
      output.add(entityPayload(dependent));
    }
  }

  private getFormulaDependencyCells(cellIndex: number): number[] {
    const formula = this.formulas.get(cellIndex);
    if (!formula) {
      return [];
    }
    return [...formula.dependencyIndices];
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
    changedInputs: Set<number>,
    formulaChanged: Set<number>,
    changedQualifiedAddresses: Set<string> | null
  ): void {
    const sheet = this.workbook.getSheet(sheetName);
    if (!sheet) return;

    const cellIndices: number[] = [];
    const impacted = new Set<number>();
    sheet.grid.forEachCell((cellIndex) => {
      cellIndices.push(cellIndex);
      this.collectFormulaDependentsForEntity(makeCellEntity(cellIndex), impacted);
    });

    cellIndices.forEach((cellIndex) => {
      changedQualifiedAddresses?.add(`${sheetName}!${this.workbook.getAddress(cellIndex)}`);
      this.removeFormula(cellIndex);
      this.reverseEdges.delete(makeCellEntity(cellIndex));
      this.workbook.cellStore.setValue(cellIndex, emptyValue());
      this.workbook.cellStore.flags[cellIndex] =
        (this.workbook.cellStore.flags[cellIndex] ?? 0) | CellFlags.PendingDelete;
      changedInputs.add(cellIndex);
    });

    this.workbook.deleteSheet(sheetName);
    this.rebindFormulasForSheet(sheetName, impacted).forEach((cellIndex) => formulaChanged.add(cellIndex));
  }

  private rebindFormulasForSheet(sheetName: string, candidates?: Set<number>): Set<number> {
    const rebound = new Set<number>();
    const targetFormulas = candidates
      ? [...candidates].map((cellIndex) => [cellIndex, this.formulas.get(cellIndex)] as const)
      : [...this.formulas.entries()];

    targetFormulas.forEach(([cellIndex, formula]) => {
      if (!formula) return;
      const ownerSheetName = this.workbook.getSheetNameById(this.workbook.cellStore.sheetIds[cellIndex]!);
      if (!ownerSheetName) return;
      const touchesSheet = formula.compiled.deps.some((dep) => {
        if (!dep.includes("!")) return false;
        const [qualifiedSheet] = dep.split("!");
        return qualifiedSheet?.replace(/^'(.*)'$/, "$1") === sheetName;
      });
      if (!touchesSheet) return;
      const dependencies = this.materializeDependencies(ownerSheetName, formula.compiled.deps, []);
      this.setFormula(cellIndex, formula.source, formula.compiled, dependencies, []);
      rebound.add(cellIndex);
    });

    return rebound;
  }

  private ensureCellTracked(sheetName: string, address: string, materializedCells: MaterializedCell[]): number {
    const ensured = this.workbook.ensureCellRecord(sheetName, address);
    if (ensured.created) {
      materializedCells.push({ sheetName, address, cellIndex: ensured.cellIndex });
    }
    return ensured.cellIndex;
  }

  private ensureCellTrackedByCoords(sheetId: number, row: number, col: number, materializedCells: MaterializedCell[]): number {
    const sheet = this.workbook.getSheetById(sheetId);
    if (!sheet) {
      throw new Error(`Unknown sheet id: ${sheetId}`);
    }
    const ensured = this.workbook.ensureCellAt(sheetId, row, col);
    if (ensured.created) {
      materializedCells.push({ sheetName: sheet.name, address: formatAddress(row, col), cellIndex: ensured.cellIndex });
    }
    return ensured.cellIndex;
  }

  private listSheetCells(sheetId: number): Array<{ cellIndex: number; row: number; col: number }> {
    const sheet = this.workbook.getSheetById(sheetId);
    if (!sheet) {
      return [];
    }
    const indices: Array<{ cellIndex: number; row: number; col: number }> = [];
    sheet.grid.forEachCell((cellIndex) => {
      indices.push({
        cellIndex,
        row: this.workbook.cellStore.rows[cellIndex] ?? 0,
        col: this.workbook.cellStore.cols[cellIndex] ?? 0
      });
    });
    return indices;
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
    sheet.grid.forEachCell((cellIndex) => {
      const row = this.workbook.cellStore.rows[cellIndex] ?? 0;
      const col = this.workbook.cellStore.cols[cellIndex] ?? 0;
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

  private syncDynamicRanges(materializedCells: readonly MaterializedCell[]): Set<number> {
    const rebound = new Set<number>();
    for (let index = 0; index < materializedCells.length; index += 1) {
      const materialized = materializedCells[index]!;
      const sheet = this.workbook.getSheet(materialized.sheetName);
      if (!sheet) {
        continue;
      }
      const row = this.workbook.cellStore.rows[materialized.cellIndex] ?? 0;
      const col = this.workbook.cellStore.cols[materialized.cellIndex] ?? 0;
      const rangeIndices = this.ranges.addDynamicMember(sheet.id, row, col, materialized.cellIndex);
      if (rangeIndices.length > 0) {
        this.scheduleWasmProgramSync();
      }
      for (let rangeCursor = 0; rangeCursor < rangeIndices.length; rangeCursor += 1) {
        const rangeIndex = rangeIndices[rangeCursor]!;
        const rangeEntity = makeRangeEntity(rangeIndex);
        this.appendReverseEdge(makeCellEntity(materialized.cellIndex), rangeEntity);
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
          if (!formula.dependencyIndices.includes(materialized.cellIndex)) {
            formula.dependencyIndices.push(materialized.cellIndex);
            rebound.add(formulaCellIndex);
          }
        }
      }
    }
    return rebound;
  }

  private resetWorkbook(workbookName = "Workbook"): void {
    this.workbook.reset(workbookName);
    this.formulas.clear();
    this.reverseEdges.clear();
    this.ranges.reset();
    this.edgeArena.reset();
    this.entityVersions.clear();
    this.sheetDeleteVersions.clear();
    this.lastMetrics = {
      batchId: 0,
      changedInputCount: 0,
      dirtyFormulaCount: 0,
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
      rangeNodeVisits: 0,
      recalcMs: 0,
      compileMs: 0
    };
    this.wasmProgramSyncPending = false;
    this.syncWasmPrograms();
  }
}

export const selectors = {
  selectCellSnapshot,
  selectViewportCells
};
