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
import { compileFormula, evaluateAst, formatAddress, parseCellAddress, parseRangeAddress } from "@bilig/formula";
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
import { EngineEventBus } from "./events.js";
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
  runtimeProgram: Uint32Array;
  constants: number[];
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
  readonly wasm = new WasmKernelFacade();

  private readonly formulas = new Map<number, RuntimeFormula>();
  private readonly dependents = new Map<number, Set<number>>();
  private readonly dependencies = new Map<number, Set<number>>();
  private readonly batchListeners = new Set<(batch: EngineOpBatch) => void>();
  private readonly entityVersions = new Map<string, OpOrder>();
  private readonly sheetDeleteVersions = new Map<string, OpOrder>();
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
          return;
        }
        ops.push({ kind: "setCellValue", sheetName, address, value: parsed.value ?? null });
      });
    });

    this.applyLocalOps(ops);
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
    const formula = this.formulas.get(cellIndex)?.source;
    if (formula !== undefined) {
      snapshot.formula = formula;
    }
    return snapshot;
  }

  getDependencies(sheetName: string, address: string): DependencySnapshot {
    const cellIndex = this.workbook.getCellIndex(sheetName, address);
    if (cellIndex === undefined) return { directDependents: [], directPrecedents: [] };
    return {
      directPrecedents: [...(this.dependencies.get(cellIndex) ?? new Set())].map((index) =>
        this.workbook.getQualifiedAddress(index)
      ),
      directDependents: [...(this.dependents.get(cellIndex) ?? new Set())].map((index) =>
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
    const ops: EngineOp[] = [];
    snapshot.sheets.forEach((sheet) => {
      ops.push({ kind: "upsertSheet", name: sheet.name, order: sheet.order });
      sheet.cells.forEach((cell) => {
        if (cell.formula) {
          ops.push({ kind: "setCellFormula", sheetName: sheet.name, address: cell.address, formula: cell.formula });
        } else {
          ops.push({ kind: "setCellValue", sheetName: sheet.name, address: cell.address, value: cell.value ?? null });
        }
      });
    });
    this.applyLocalOps(ops);
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
            break;
          }
          engineOps.push({ kind: "setCellValue", sheetName: op.sheetName, address: op.addr, value: op.value ?? null });
          break;
        case "deleteCell":
          if (op.sheetName && op.addr) {
            engineOps.push({ kind: "clearCell", sheetName: op.sheetName, address: op.addr });
          }
          break;
      }
    });
    this.applyLocalOps(engineOps);
  }

  applyRemoteBatch(batch: EngineOpBatch): void {
    if (!shouldApplyBatch(this.replica, batch)) return;
    this.applyBatch(batch, "remote");
  }

  private applyLocalOps(ops: EngineOp[]): void {
    if (ops.length === 0) return;
    const batch = createBatch(this.replica, ops);
    this.applyBatch(batch, "local");
  }

  private applyBatch(batch: EngineOpBatch, source: "local" | "remote"): void {
    const changedInputs = new Set<number>();
    const formulaChanged = new Set<number>();
    let appliedOps = 0;

    batch.ops.forEach((op, opIndex) => {
      const order = batchOpOrder(batch, opIndex);
      if (!this.shouldApplyOp(op, order)) {
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
          this.rebindFormulasForSheet(op.name).forEach((cellIndex) => formulaChanged.add(cellIndex));
          break;
        case "deleteSheet":
          this.removeSheetRuntime(op.name, changedInputs, formulaChanged);
          this.entityVersions.set(this.entityKeyForOp(op), order);
          this.sheetDeleteVersions.set(op.name, order);
          break;
        case "setCellValue": {
          const cellIndex = this.workbook.ensureCell(op.sheetName, op.address);
          this.removeFormula(cellIndex);
          const value = literalToValue(op.value, this.strings);
          this.workbook.cellStore.setValue(
            cellIndex,
            value,
            value.tag === ValueTag.String ? value.stringId : 0
          );
          this.workbook.cellStore.flags[cellIndex] = (this.workbook.cellStore.flags[cellIndex] ?? 0) & ~CellFlags.HasFormula;
          changedInputs.add(cellIndex);
          this.entityVersions.set(this.entityKeyForOp(op), order);
          break;
        }
        case "setCellFormula": {
          const cellIndex = this.workbook.ensureCell(op.sheetName, op.address);
          const compileStarted = performance.now();
          const compiled = compileFormula(op.formula);
          this.lastMetrics.compileMs = performance.now() - compileStarted;
          const dependencyIndices = this.materializeDependencies(op.sheetName, compiled.deps);
          this.setFormula(cellIndex, op.formula, compiled, dependencyIndices);
          formulaChanged.add(cellIndex);
          this.entityVersions.set(this.entityKeyForOp(op), order);
          break;
        }
        case "clearCell": {
          const cellIndex = this.workbook.getCellIndex(op.sheetName, op.address);
          if (cellIndex === undefined) {
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          }
          this.removeFormula(cellIndex);
          this.workbook.cellStore.setValue(cellIndex, emptyValue());
          changedInputs.add(cellIndex);
          this.entityVersions.set(this.entityKeyForOp(op), order);
          break;
        }
      }
      appliedOps += 1;
    });

    markBatchApplied(this.replica, batch);
    if (appliedOps === 0) {
      if (source === "local") {
        this.emitBatch(batch);
      }
      return;
    }

    this.rebuildTopoRanks();
    this.detectCycles();
    const changed = this.recalculate([...changedInputs, ...formulaChanged]);
    this.lastMetrics.batchId += 1;
    this.lastMetrics.changedInputCount = changedInputs.size + formulaChanged.size;
    this.events.emit({
      kind: "batch",
      changedCellIndices: changed,
      metrics: this.lastMetrics
    });
    if (source === "local") {
      this.emitBatch(batch);
    }
  }

  private materializeDependencies(currentSheet: string, deps: string[]): number[] {
    const indices: number[] = [];
    for (const dep of deps) {
      if (dep.includes(":")) {
        const range = parseRangeAddress(dep, currentSheet);
        const sheetName = range.sheetName ?? currentSheet;
        if (range.sheetName && !this.workbook.getSheet(sheetName)) {
          continue;
        }
        for (let row = range.start.row; row <= range.end.row; row += 1) {
          for (let col = range.start.col; col <= range.end.col; col += 1) {
            indices.push(this.workbook.ensureCell(sheetName, formatAddress(row, col)));
          }
        }
        continue;
      }
      const parsed = parseCellAddress(dep, currentSheet);
      const sheetName = parsed.sheetName ?? currentSheet;
      if (parsed.sheetName && !this.workbook.getSheet(sheetName)) {
        continue;
      }
      indices.push(this.workbook.ensureCell(sheetName, parsed.text));
    }
    return [...new Set(indices)];
  }

  private setFormula(
    cellIndex: number,
    source: string,
    compiled: ReturnType<typeof compileFormula>,
    dependencyIndices: number[]
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
      symbolicRefToIndex.set(ref, this.workbook.ensureCell(sheetName, parsed.text));
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
      }
    });

    this.formulas.set(cellIndex, {
      cellIndex,
      source,
      compiled: effectiveCompiled,
      dependencyIndices,
      runtimeProgram,
      constants: effectiveCompiled.constants
    });
    this.workbook.cellStore.flags[cellIndex] = (this.workbook.cellStore.flags[cellIndex] ?? 0) | CellFlags.HasFormula;
    if (effectiveCompiled.mode === FormulaMode.JsOnly) {
      this.workbook.cellStore.flags[cellIndex] = (this.workbook.cellStore.flags[cellIndex] ?? 0) | CellFlags.JsOnly;
    } else {
      this.workbook.cellStore.flags[cellIndex] = (this.workbook.cellStore.flags[cellIndex] ?? 0) & ~CellFlags.JsOnly;
    }

    this.dependencies.set(cellIndex, new Set(dependencyIndices));
    dependencyIndices.forEach((dependencyIndex) => {
      let dependents = this.dependents.get(dependencyIndex);
      if (!dependents) {
        dependents = new Set<number>();
        this.dependents.set(dependencyIndex, dependents);
      }
      dependents.add(cellIndex);
    });
    this.syncWasmPrograms();
  }

  private removeFormula(cellIndex: number): void {
    const existingDeps = this.dependencies.get(cellIndex);
    if (existingDeps) {
      existingDeps.forEach((depIndex) => {
        this.dependents.get(depIndex)?.delete(cellIndex);
      });
    }
    this.dependencies.delete(cellIndex);
    this.formulas.delete(cellIndex);
    this.workbook.cellStore.flags[cellIndex] =
      (this.workbook.cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle);
    this.workbook.cellStore.formulaIds[cellIndex] = 0;
    this.syncWasmPrograms();
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
    while (queue.length > 0) {
      const cellIndex = queue.shift()!;
      this.workbook.cellStore.topoRanks[cellIndex] = rank++;
      const dependents = this.dependents.get(cellIndex);
      if (!dependents) continue;
      dependents.forEach((dependent) => {
        if (!indegree.has(dependent)) return;
        const next = (indegree.get(dependent) ?? 0) - 1;
        indegree.set(dependent, next);
        if (next === 0) queue.push(dependent);
      });
    }
  }

  private detectCycles(): void {
    const visiting = new Set<number>();
    const visited = new Set<number>();

    const visit = (cellIndex: number, stack: number[]): void => {
      if (visiting.has(cellIndex)) {
        stack.forEach((member) => {
          this.workbook.cellStore.flags[member] = (this.workbook.cellStore.flags[member] ?? 0) | CellFlags.InCycle;
          this.workbook.cellStore.setValue(member, errorValue(ErrorCode.Cycle));
        });
        return;
      }
      if (visited.has(cellIndex)) return;
      visited.add(cellIndex);
      visiting.add(cellIndex);
      const deps = this.dependencies.get(cellIndex) ?? new Set<number>();
      deps.forEach((dep) => {
        if (this.formulas.has(dep)) {
          visit(dep, [...stack, dep]);
        }
      });
      visiting.delete(cellIndex);
    };

    this.formulas.forEach((_formula, cellIndex) => {
      this.workbook.cellStore.flags[cellIndex] = (this.workbook.cellStore.flags[cellIndex] ?? 0) & ~CellFlags.InCycle;
    });

    this.formulas.forEach((_formula, cellIndex) => visit(cellIndex, [cellIndex]));
  }

  private recalculate(changedRoots: number[]): Uint32Array {
    const started = performance.now();
    const dirty = new Set<number>();
    const queue = [...changedRoots];

    changedRoots.forEach((cellIndex) => {
      if (this.formulas.has(cellIndex)) {
        dirty.add(cellIndex);
      }
    });

    while (queue.length > 0) {
      const cellIndex = queue.shift()!;
      const dependents = this.dependents.get(cellIndex);
      if (!dependents) continue;
      dependents.forEach((dependent) => {
        if (!dirty.has(dependent)) {
          dirty.add(dependent);
          queue.push(dependent);
        }
      });
    }

    const ordered = [...dirty].sort(
      (left, right) => (this.workbook.cellStore.topoRanks[left] ?? 0) - (this.workbook.cellStore.topoRanks[right] ?? 0)
    );

    const wasmBatch: number[] = [];
    let jsCount = 0;
    ordered.forEach((cellIndex) => {
      const formula = this.formulas.get(cellIndex);
      if (!formula) return;
      if (((this.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0) {
        return;
      }
      if (formula.compiled.mode === FormulaMode.WasmFastPath && this.wasm.ready) {
        wasmBatch.push(cellIndex);
        return;
      }
      jsCount += 1;
      this.evaluateFormulaJs(cellIndex, formula);
    });

    if (wasmBatch.length > 0) {
      this.wasm.syncFromStore(this.workbook.cellStore);
      this.wasm.evalBatch(Uint32Array.from(wasmBatch));
      this.wasm.syncToStore(this.workbook.cellStore, Uint32Array.from(wasmBatch));
    }

    this.lastMetrics.dirtyFormulaCount = ordered.length;
    this.lastMetrics.jsFormulaCount = jsCount;
    this.lastMetrics.wasmFormulaCount = wasmBatch.length;
    this.lastMetrics.recalcMs = performance.now() - started;

    return Uint32Array.from([...new Set([...changedRoots, ...ordered])]);
  }

  private evaluateFormulaJs(cellIndex: number, formula: RuntimeFormula): void {
    const sheetName = this.workbook.getSheetNameById(this.workbook.cellStore.sheetIds[cellIndex]!);
    const value = evaluateAst(formula.compiled.ast, {
      sheetName,
      resolveCell: (targetSheetName, address) => {
        const targetCell = this.workbook.getCellIndex(targetSheetName, address);
        if (targetCell === undefined) {
          return targetSheetName === sheetName ? emptyValue() : errorValue(ErrorCode.Ref);
        }
        return this.workbook.cellStore.getValue(targetCell, (id) => this.strings.get(id));
      },
      resolveRange: (targetSheetName, start, end) => {
        if (targetSheetName && !this.workbook.getSheet(targetSheetName)) {
          return [errorValue(ErrorCode.Ref)];
        }
        const range = parseRangeAddress(`${start}:${end}`, targetSheetName);
        const values: CellValue[] = [];
        for (let row = range.start.row; row <= range.end.row; row += 1) {
          for (let col = range.start.col; col <= range.end.col; col += 1) {
            const addr = formatAddress(row, col);
            const index = this.workbook.ensureCell(targetSheetName, addr);
            values.push(this.workbook.cellStore.getValue(index, (id) => this.strings.get(id)));
          }
        }
        return values;
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
  }

  private emitBatch(batch: EngineOpBatch): void {
    this.batchListeners.forEach((listener) => listener(batch));
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
      case "setCellValue":
      case "setCellFormula":
      case "clearCell":
        return `cell:${op.sheetName}!${op.address}`;
    }
  }

  private sheetDeleteBarrierForOp(op: EngineOp): OpOrder | undefined {
    switch (op.kind) {
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

  private removeSheetRuntime(sheetName: string, changedInputs: Set<number>, formulaChanged: Set<number>): void {
    const sheet = this.workbook.getSheet(sheetName);
    if (!sheet) return;

    const cellIndices: number[] = [];
    const impacted = new Set<number>();
    sheet.grid.forEachCell((cellIndex) => {
      cellIndices.push(cellIndex);
      this.dependents.get(cellIndex)?.forEach((dependent) => {
        if (this.formulas.has(dependent)) {
          impacted.add(dependent);
        }
      });
    });

    cellIndices.forEach((cellIndex) => {
      this.removeFormula(cellIndex);
      this.dependencies.delete(cellIndex);
      this.dependents.delete(cellIndex);
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
      const dependencyIndices = this.materializeDependencies(ownerSheetName, formula.compiled.deps);
      this.setFormula(cellIndex, formula.source, formula.compiled, dependencyIndices);
      rebound.add(cellIndex);
    });

    return rebound;
  }

  private resetWorkbook(workbookName = "Workbook"): void {
    this.workbook.reset(workbookName);
    this.formulas.clear();
    this.dependents.clear();
    this.dependencies.clear();
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
    this.syncWasmPrograms();
  }
}

export const selectors = {
  selectCellSnapshot,
  selectViewportCells
};
