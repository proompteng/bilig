import {
  createCellNumberFormatRecord,
  type CellNumberFormatInput,
  type CellNumberFormatRecord,
  type CellRangeRef,
  type CellStyleField,
  type CellStylePatch,
  type CellStyleRecord,
  ErrorCode,
  FormulaMode,
  MAX_COLS,
  MAX_ROWS,
  Opcode,
  ValueTag,
  type CellSnapshot,
  type CellValue,
  type DependencySnapshot,
  type EngineEvent,
  type ExplainCellSnapshot,
  type LiteralInput,
  type RecalcMetrics,
  type SyncState,
  type SelectionState,
  type SheetMetadataSnapshot,
  type SheetFormatRangeSnapshot,
  type SheetStyleRangeSnapshot,
  type WorkbookAxisEntrySnapshot,
  type WorkbookAxisMetadataSnapshot,
  type WorkbookCalculationSettingsSnapshot,
  type WorkbookDefinedNameValueSnapshot,
  type WorkbookFreezePaneSnapshot,
  type WorkbookPivotSnapshot,
  type WorkbookSortSnapshot,
  type WorkbookSnapshot,
} from "@bilig/protocol";
import {
  type FormulaNode,
  compileFormula,
  compileFormulaAst,
  evaluatePlanResult,
  formatAddress,
  isArrayValue,
  parseCellAddress,
  parseFormula,
  parseRangeAddress,
  renameFormulaSheetReferences,
  rewriteAddressForStructuralTransform,
  rewriteFormulaForStructuralTransform,
  rewriteRangeForStructuralTransform,
  type StructuralAxisTransform,
  translateFormulaReferences,
  utcDateToExcelSerial,
} from "@bilig/formula";
import { Float64Arena, Uint32Arena } from "@bilig/formula/program-arena";
import type { EngineOp, EngineOpBatch } from "@bilig/workbook-domain";
import {
  batchOpOrder,
  compareOpOrder,
  createBatch,
  createReplicaState,
  exportReplicaSnapshot as exportReplicaStateSnapshot,
  hydrateReplicaState,
  markBatchApplied,
  shouldApplyBatch,
  type OpOrder,
  type ReplicaSnapshot,
  type ReplicaVersionSnapshot,
  type ReplicaState,
} from "./replica-state.js";
import { CellFlags } from "./cell-store.js";
import { CycleDetector } from "./cycle-detection.js";
import { EdgeArena, type EdgeSlice } from "./edge-arena.js";
import { entityPayload, isRangeEntity, makeCellEntity, makeRangeEntity } from "./entity-ids.js";
import {
  applyStylePatch,
  clearStyleFields,
  cloneCellStyleRecord,
  normalizeCellStylePatch,
} from "./engine-style-utils.js";
import {
  mapStructuralAxisIndex,
  mapStructuralBoundary,
  structuralTransformForOp,
} from "./engine-structural-utils.js";
import { EngineEventBus } from "./events.js";
import { FormulaTable } from "./formula-table.js";
import { materializePivotTable, type PivotDefinitionInput } from "./pivot-engine.js";
import { RangeRegistry } from "./range-registry.js";
import { RecalcScheduler } from "./scheduler.js";
import {
  selectCellSnapshot,
  selectMetrics,
  selectSelectionState,
  selectViewportCells,
} from "./selectors.js";
import { StringPool } from "./string-pool.js";
import { WasmKernelFacade } from "./wasm-facade.js";
import {
  WorkbookStore,
  makeCellKey,
  normalizeDefinedName,
  normalizeWorkbookObjectName,
  pivotKey,
  type WorkbookAxisMetadataRecord,
  type WorkbookCalculationSettingsRecord,
  type WorkbookDefinedNameRecord,
  type WorkbookFilterRecord,
  type WorkbookPivotRecord,
  type WorkbookPropertyRecord,
  type WorkbookSortRecord,
  type WorkbookSpillRecord,
  type WorkbookTableRecord,
  type WorkbookVolatileContextRecord,
} from "./workbook-store.js";
import { cellToCsvValue, parseCsv, parseCsvCellInput, serializeCsv } from "./csv.js";

export interface CommitOp {
  kind:
    | "upsertWorkbook"
    | "upsertSheet"
    | "renameSheet"
    | "deleteSheet"
    | "upsertCell"
    | "deleteCell";
  name?: string;
  oldName?: string;
  newName?: string;
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

export interface EngineSyncClientConnection {
  send(batch: EngineOpBatch): void | Promise<void>;
  disconnect(): void | Promise<void>;
}

export interface EngineSyncClient {
  connect(handlers: {
    applyRemoteBatch(batch: EngineOpBatch): boolean;
    applyRemoteSnapshot?(snapshot: WorkbookSnapshot): void;
    setState(state: SyncState): void;
  }): EngineSyncClientConnection | Promise<EngineSyncClientConnection>;
}

export interface EngineReplicaSnapshot {
  replica: ReplicaSnapshot;
  entityVersions: ReplicaVersionSnapshot[];
  sheetDeleteVersions: Array<{ sheetName: string; order: OpOrder }>;
}

interface TransactionRecord {
  ops: EngineOp[];
  potentialNewCells?: number;
}

interface TransactionLogEntry {
  forward: TransactionRecord;
  inverse: TransactionRecord;
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

type U32 = Uint32Array;

const UNRESOLVED_WASM_OPERAND = 0x00ff_ffff;

interface MaterializedDependencies {
  dependencyIndices: Uint32Array;
  dependencyEntities: Uint32Array;
  rangeDependencies: Uint32Array;
  symbolicRangeIndices: U32;
  symbolicRangeCount: number;
  newRangeIndices: U32;
  newRangeCount: number;
}

interface SpillMaterialization {
  changedCellIndices: number[];
  ownerValue: CellValue;
}

export type PivotTableInput = Omit<
  WorkbookPivotSnapshot,
  "sheetName" | "address" | "rows" | "cols"
>;

interface RecalcVolatileState {
  nowSerial: number;
  randomValues: number[];
  randomCursor: number;
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

function areCellValuesEqual(left: CellValue, right: CellValue): boolean {
  if (left.tag !== right.tag) {
    return false;
  }
  switch (left.tag) {
    case ValueTag.Empty:
      return true;
    case ValueTag.Number:
      return right.tag === ValueTag.Number && Object.is(left.value, right.value);
    case ValueTag.Boolean:
      return right.tag === ValueTag.Boolean && left.value === right.value;
    case ValueTag.String:
      return right.tag === ValueTag.String && left.value === right.value;
    case ValueTag.Error:
      return right.tag === ValueTag.Error && left.code === right.code;
  }
}

function cellValueDisplayText(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return "";
    case ValueTag.Number:
      return Object.is(value.value, -0) ? "-0" : String(value.value);
    case ValueTag.Boolean:
      return value.value ? "TRUE" : "FALSE";
    case ValueTag.String:
      return value.value;
    case ValueTag.Error:
      return `#${ErrorCode[value.code] ?? "ERROR"}!`;
  }
}

function normalizePivotLookupText(value: string): string {
  return value.trim().toUpperCase();
}

function pivotItemMatches(cell: CellValue, item: CellValue): boolean {
  if (areCellValuesEqual(cell, item)) {
    return true;
  }
  if (item.tag === ValueTag.String) {
    return (
      normalizePivotLookupText(cellValueDisplayText(cell)) === normalizePivotLookupText(item.value)
    );
  }
  return false;
}

function assertNever(value: never): never {
  throw new Error(`Unexpected value: ${String(value)}`);
}

function literalToFormulaNode(input: LiteralInput): FormulaNode | null {
  if (typeof input === "number") {
    return { kind: "NumberLiteral", value: input };
  }
  if (typeof input === "string") {
    return { kind: "StringLiteral", value: input };
  }
  if (typeof input === "boolean") {
    return { kind: "BooleanLiteral", value: input };
  }
  return null;
}

function definedNameValueToFormulaNode(
  input: WorkbookDefinedNameValueSnapshot,
): FormulaNode | null {
  if (typeof input === "object" && input !== null && "kind" in input) {
    switch (input.kind) {
      case "scalar":
        return literalToFormulaNode(input.value);
      case "cell-ref":
        return { kind: "CellRef", ref: input.address, sheetName: input.sheetName };
      case "range-ref":
        return {
          kind: "RangeRef",
          refKind: "cells",
          start: input.startAddress,
          end: input.endAddress,
          sheetName: input.sheetName,
        };
      case "structured-ref":
        return {
          kind: "StructuredRef",
          tableName: input.tableName,
          columnName: input.columnName,
        };
      case "formula":
        try {
          return parseFormula(input.formula);
        } catch {
          return { kind: "ErrorLiteral", code: ErrorCode.Value };
        }
    }
  }
  if (typeof input === "string" && input.startsWith("=")) {
    try {
      return parseFormula(input);
    } catch {
      return { kind: "ErrorLiteral", code: ErrorCode.Value };
    }
  }
  return literalToFormulaNode(input);
}

interface MetadataResolutionContext {
  resolveName: (name: string) => WorkbookDefinedNameValueSnapshot | undefined;
  resolveStructuredReference: (tableName: string, columnName: string) => FormulaNode | undefined;
  resolveSpillReference: (
    sheetName: string | undefined,
    address: string,
  ) => FormulaNode | undefined;
}

function definedNameValuesEqual(
  left: WorkbookDefinedNameValueSnapshot,
  right: WorkbookDefinedNameValueSnapshot,
): boolean {
  if (left === right) {
    return true;
  }
  return JSON.stringify(left) === JSON.stringify(right);
}

function definedNameValueToCellValue(
  input: WorkbookDefinedNameValueSnapshot,
  stringPool: StringPool,
): CellValue {
  if (typeof input === "object" && input !== null && "kind" in input) {
    if (input.kind === "scalar") {
      return literalToValue(input.value, stringPool);
    }
    return errorValue(ErrorCode.Value);
  }
  return literalToValue(input, stringPool);
}

function renameDefinedNameValueSheet(
  input: WorkbookDefinedNameValueSnapshot,
  oldSheetName: string,
  newSheetName: string,
): WorkbookDefinedNameValueSnapshot {
  if (typeof input === "object" && input !== null && "kind" in input) {
    switch (input.kind) {
      case "scalar":
      case "structured-ref":
        return input;
      case "cell-ref":
        return input.sheetName === oldSheetName ? { ...input, sheetName: newSheetName } : input;
      case "range-ref":
        return input.sheetName === oldSheetName ? { ...input, sheetName: newSheetName } : input;
      case "formula":
        return {
          ...input,
          formula: renameFormulaTextForSheet(input.formula, oldSheetName, newSheetName),
        };
    }
  }
  if (typeof input === "string" && input.startsWith("=")) {
    return renameFormulaTextForSheet(input, oldSheetName, newSheetName);
  }
  return input;
}

function renameFormulaTextForSheet(
  input: string,
  oldSheetName: string,
  newSheetName: string,
): string {
  const hasLeadingEquals = input.startsWith("=");
  const source = hasLeadingEquals ? input.slice(1) : input;
  const rewritten = renameFormulaSheetReferences(source, oldSheetName, newSheetName);
  return hasLeadingEquals ? `=${rewritten}` : rewritten;
}

function resolveMetadataReferencesInAst(
  node: FormulaNode,
  context: MetadataResolutionContext,
  activeNames = new Set<string>(),
): { node: FormulaNode; fullyResolved: boolean; substituted: boolean } {
  switch (node.kind) {
    case "NumberLiteral":
    case "BooleanLiteral":
    case "StringLiteral":
    case "ErrorLiteral":
    case "CellRef":
    case "RowRef":
    case "ColumnRef":
    case "RangeRef":
      return { node, fullyResolved: true, substituted: false };
    case "NameRef": {
      const normalized = normalizeDefinedName(node.name);
      if (activeNames.has(normalized)) {
        return {
          node: { kind: "ErrorLiteral", code: ErrorCode.Cycle },
          fullyResolved: true,
          substituted: true,
        };
      }
      const literal = context.resolveName(node.name);
      const replacement =
        literal === undefined
          ? ({ kind: "ErrorLiteral", code: ErrorCode.Name } satisfies FormulaNode)
          : definedNameValueToFormulaNode(literal);
      if (!replacement) {
        return { node, fullyResolved: false, substituted: false };
      }
      const nextActiveNames = new Set(activeNames);
      nextActiveNames.add(normalized);
      const resolved = resolveMetadataReferencesInAst(replacement, context, nextActiveNames);
      return { node: resolved.node, fullyResolved: resolved.fullyResolved, substituted: true };
    }
    case "StructuredRef": {
      const replacement =
        context.resolveStructuredReference(node.tableName, node.columnName) ??
        ({ kind: "ErrorLiteral", code: ErrorCode.Ref } satisfies FormulaNode);
      return { node: replacement, fullyResolved: true, substituted: true };
    }
    case "SpillRef": {
      const replacement =
        context.resolveSpillReference(node.sheetName, node.ref) ??
        ({ kind: "ErrorLiteral", code: ErrorCode.Ref } satisfies FormulaNode);
      return { node: replacement, fullyResolved: true, substituted: true };
    }
    case "UnaryExpr": {
      const resolved = resolveMetadataReferencesInAst(node.argument, context, activeNames);
      return {
        node: resolved.substituted ? { ...node, argument: resolved.node } : node,
        fullyResolved: resolved.fullyResolved,
        substituted: resolved.substituted,
      };
    }
    case "BinaryExpr": {
      const left = resolveMetadataReferencesInAst(node.left, context, activeNames);
      const right = resolveMetadataReferencesInAst(node.right, context, activeNames);
      return {
        node:
          left.substituted || right.substituted
            ? { ...node, left: left.node, right: right.node }
            : node,
        fullyResolved: left.fullyResolved && right.fullyResolved,
        substituted: left.substituted || right.substituted,
      };
    }
    case "CallExpr": {
      let fullyResolved = true;
      let substituted = false;
      const args = node.args.map((arg) => {
        const resolved = resolveMetadataReferencesInAst(arg, context, activeNames);
        fullyResolved = fullyResolved && resolved.fullyResolved;
        substituted = substituted || resolved.substituted;
        return resolved.node;
      });
      return {
        node: substituted ? { ...node, args } : node,
        fullyResolved,
        substituted,
      };
    }
    case "InvokeExpr": {
      const callee = resolveMetadataReferencesInAst(node.callee, context, activeNames);
      let fullyResolved = callee.fullyResolved;
      let substituted = callee.substituted;
      const args = node.args.map((arg) => {
        const resolved = resolveMetadataReferencesInAst(arg, context, activeNames);
        fullyResolved = fullyResolved && resolved.fullyResolved;
        substituted = substituted || resolved.substituted;
        return resolved.node;
      });
      return {
        node: substituted ? { ...node, callee: callee.node, args } : node,
        fullyResolved,
        substituted,
      };
    }
  }
}

function tableDependencyKey(name: string): string {
  return normalizeWorkbookObjectName(name, "Table");
}

function spillDependencyKey(sheetName: string, address: string): string {
  return `${sheetName}!${parseCellAddress(address, sheetName).text}`;
}

function spillDependencyKeyFromRef(ref: string, ownerSheetName: string): string {
  if (ref.includes("!")) {
    const separator = ref.indexOf("!");
    const sheetName = ref.slice(0, separator).replace(/^'(.*)'$/, "$1");
    const address = ref.slice(separator + 1);
    return spillDependencyKey(sheetName, address);
  }
  return spillDependencyKey(ownerSheetName, ref);
}

export class SpreadsheetEngine {
  readonly workbook: WorkbookStore;
  readonly strings = new StringPool();
  readonly events = new EngineEventBus();
  private readonly replicaState: ReplicaState;
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
  private readonly reverseDefinedNameEdges = new Map<string, Set<number>>();
  private readonly reverseTableEdges = new Map<string, Set<number>>();
  private readonly reverseSpillEdges = new Map<string, Set<number>>();
  private readonly pivotOutputOwners = new Map<number, string>();
  private readonly batchListeners = new Set<(batch: EngineOpBatch) => void>();
  private readonly selectionListeners = new Set<() => void>();
  private readonly entityVersions = new Map<string, OpOrder>();
  private readonly sheetDeleteVersions = new Map<string, OpOrder>();
  private selection: SelectionState = {
    sheetName: "Sheet1",
    address: "A1",
    anchorAddress: "A1",
    range: { startAddress: "A1", endAddress: "A1" },
    editMode: "idle",
  };
  private syncState: SyncState = "local-only";
  private syncClientConnection: EngineSyncClientConnection | null = null;
  private readonly undoStack: TransactionLogEntry[] = [];
  private readonly redoStack: TransactionLogEntry[] = [];
  private transactionReplayDepth = 0;
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
  private wasmRangeRowCounts: U32 = new Uint32Array(128);
  private wasmRangeColCounts: U32 = new Uint32Array(128);
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
    compileMs: 0,
  };

  constructor(options: SpreadsheetEngineOptions = {}) {
    this.workbook = new WorkbookStore(options.workbookName ?? "Workbook");
    this.formulas = new FormulaTable(this.workbook.cellStore);
    this.replicaState = createReplicaState(options.replicaId ?? "local");
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

  subscribeCells(
    sheetName: string,
    addresses: readonly string[],
    listener: () => void,
  ): () => void {
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

  setSelection(
    sheetName: string,
    address: string | null,
    options: {
      anchorAddress?: string | null;
      range?: { startAddress: string; endAddress: string } | null;
      editMode?: SelectionState["editMode"];
    } = {},
  ): void {
    const nextSelection: SelectionState = {
      sheetName,
      address,
      anchorAddress: options.anchorAddress ?? address,
      range: options.range ?? (address ? { startAddress: address, endAddress: address } : null),
      editMode: options.editMode ?? this.selection.editMode,
    };

    if (
      this.selection.sheetName === nextSelection.sheetName &&
      this.selection.address === nextSelection.address &&
      this.selection.anchorAddress === nextSelection.anchorAddress &&
      this.selection.editMode === nextSelection.editMode &&
      this.selection.range?.startAddress === nextSelection.range?.startAddress &&
      this.selection.range?.endAddress === nextSelection.range?.endAddress
    ) {
      return;
    }

    this.selection = nextSelection;
    this.selectionListeners.forEach((listener) => listener());
  }

  getLastMetrics(): RecalcMetrics {
    return this.lastMetrics;
  }

  getSyncState(): SyncState {
    return this.syncState;
  }

  async connectSyncClient(client: EngineSyncClient): Promise<void> {
    await this.disconnectSyncClient();
    this.setSyncState("syncing");
    const connection = await client.connect({
      applyRemoteBatch: (batch) => {
        return this.applyRemoteBatch(batch);
      },
      applyRemoteSnapshot: (snapshot) => {
        this.importSnapshot(snapshot);
      },
      setState: (state) => {
        this.setSyncState(state);
      },
    });
    this.syncClientConnection = connection;
    if (this.syncState === "syncing") {
      this.setSyncState("live");
    }
  }

  async disconnectSyncClient(): Promise<void> {
    const connection = this.syncClientConnection;
    this.syncClientConnection = null;
    if (connection) {
      await connection.disconnect();
    }
    this.setSyncState("local-only");
  }

  private setSyncState(state: SyncState): void {
    this.syncState = state;
  }

  createSheet(name: string): void {
    this.executeLocalTransaction([
      { kind: "upsertSheet", name, order: this.workbook.sheetsByName.size },
    ]);
  }

  renameSheet(oldName: string, newName: string): void {
    const trimmedName = newName.trim();
    if (trimmedName.length === 0 || oldName === trimmedName) {
      return;
    }
    if (this.workbook.getSheet(trimmedName)) {
      return;
    }
    this.executeLocalTransaction([{ kind: "renameSheet", oldName, newName: trimmedName }]);
  }

  deleteSheet(name: string): void {
    this.executeLocalTransaction([{ kind: "deleteSheet", name }]);
  }

  setCellValue(sheetName: string, address: string, value: LiteralInput): CellValue {
    this.executeLocalTransaction([{ kind: "setCellValue", sheetName, address, value }]);
    return this.getCellValue(sheetName, address);
  }

  setCellFormula(sheetName: string, address: string, formula: string): CellValue {
    this.executeLocalTransaction([{ kind: "setCellFormula", sheetName, address, formula }]);
    return this.getCellValue(sheetName, address);
  }

  setCellFormat(sheetName: string, address: string, format: string | null): void {
    this.executeLocalTransaction([{ kind: "setCellFormat", sheetName, address, format }]);
  }

  setRangeNumberFormat(range: CellRangeRef, format: CellNumberFormatInput): void {
    const ops = this.buildFormatPatchOps(range, format);
    this.executeLocalTransaction(ops);
  }

  clearRangeNumberFormat(range: CellRangeRef): void {
    const ops = this.buildFormatClearOps(range);
    this.executeLocalTransaction(ops);
  }

  setRangeStyle(range: CellRangeRef, patch: CellStylePatch): void {
    const ops = this.buildStylePatchOps(range, patch);
    this.executeLocalTransaction(ops);
  }

  clearRangeStyle(range: CellRangeRef, fields?: readonly CellStyleField[]): void {
    const ops = this.buildStyleClearOps(range, fields);
    this.executeLocalTransaction(ops);
  }

  getCellStyle(styleId: string | undefined): CellStyleRecord | undefined {
    return this.workbook.getCellStyle(styleId);
  }

  getCellNumberFormat(id: string | undefined): CellNumberFormatRecord | undefined {
    return this.workbook.getCellNumberFormat(id);
  }

  setDefinedName(name: string, value: WorkbookDefinedNameValueSnapshot): void {
    const normalizedName = normalizeDefinedName(name);
    const previous = this.workbook.getDefinedName(normalizedName);
    const trimmedName = name.trim();
    if (previous?.name === trimmedName && definedNameValuesEqual(previous.value, value)) {
      return;
    }
    this.executeLocalTransaction([{ kind: "upsertDefinedName", name: trimmedName, value }]);
  }

  deleteDefinedName(name: string): boolean {
    if (!this.workbook.getDefinedName(name)) {
      return false;
    }
    this.executeLocalTransaction([{ kind: "deleteDefinedName", name }]);
    return true;
  }

  getDefinedName(name: string): WorkbookDefinedNameRecord | undefined {
    return this.workbook.getDefinedName(name);
  }

  getDefinedNames(): WorkbookDefinedNameRecord[] {
    return this.workbook.listDefinedNames();
  }

  setWorkbookMetadata(key: string, value: LiteralInput): void {
    const existing = this.workbook.getWorkbookProperty(key);
    if (existing?.value === value || (existing === undefined && value === null)) {
      return;
    }
    this.executeLocalTransaction([{ kind: "setWorkbookMetadata", key, value }]);
  }

  getWorkbookMetadata(key: string): WorkbookPropertyRecord | undefined {
    return this.workbook.getWorkbookProperty(key);
  }

  getWorkbookMetadataEntries(): WorkbookPropertyRecord[] {
    return this.workbook.listWorkbookProperties();
  }

  setCalculationSettings(settings: WorkbookCalculationSettingsSnapshot): void {
    const current = this.workbook.getCalculationSettings();
    const nextSettings = { compatibilityMode: "excel-modern" as const, ...settings };
    if (
      current.mode === nextSettings.mode &&
      current.compatibilityMode === nextSettings.compatibilityMode
    ) {
      return;
    }
    this.executeLocalTransaction([{ kind: "setCalculationSettings", settings: nextSettings }]);
  }

  getCalculationSettings(): WorkbookCalculationSettingsRecord {
    return this.workbook.getCalculationSettings();
  }

  getVolatileContext(): WorkbookVolatileContextRecord {
    return this.workbook.getVolatileContext();
  }

  recalculateNow(): number[] {
    this.workbook.setVolatileContext({
      recalcEpoch: this.workbook.getVolatileContext().recalcEpoch + 1,
    });
    let formulaChangedCount = 0;
    let explicitChangedCount = 0;
    this.formulas.forEach((_formula, cellIndex) => {
      formulaChangedCount = this.markFormulaChanged(cellIndex, formulaChangedCount);
      explicitChangedCount = this.markExplicitChanged(cellIndex, explicitChangedCount);
    });
    const recalculated = this.reconcilePivotOutputs(
      this.recalculate(
        this.composeMutationRoots(0, formulaChangedCount),
        this.changedInputBuffer.subarray(0, 0),
      ),
      true,
    );
    const changed = this.composeEventChanges(recalculated, explicitChangedCount);
    this.lastMetrics.batchId += 1;
    this.lastMetrics.changedInputCount = formulaChangedCount;
    this.events.emit(
      {
        kind: "batch",
        invalidation: "cells",
        changedCellIndices: changed,
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        metrics: this.lastMetrics,
      },
      changed,
      (cellIndex) => this.workbook.getQualifiedAddress(cellIndex),
    );
    return Array.from(changed);
  }

  recalculateDifferential(): { js: CellSnapshot[]; wasm: CellSnapshot[]; drift: string[] } {
    // 1. Snapshot original state
    const originalSnapshot = this.exportSnapshot();

    // 2. JS Only Recalc
    this.formulas.forEach((f) => {
      f.compiled.mode = FormulaMode.JsOnly;
    });
    const jsChanged = this.recalculateNow();
    const jsResults = jsChanged.map((idx) => this.getCellByIndex(idx));

    // 3. Restore and rerun using the engine's normal compiled formula modes.
    this.importSnapshot(originalSnapshot);
    const wasmChanged = this.recalculateNow();
    const wasmResults = wasmChanged.map((idx) => this.getCellByIndex(idx));

    // 4. Compare
    const drift: string[] = [];
    const jsMap = new Map(
      jsResults.map((result) => [`${result.sheetName}!${result.address}`, result]),
    );
    const wasmMap = new Map(
      wasmResults.map((result) => [`${result.sheetName}!${result.address}`, result]),
    );

    for (const [addr, jsCell] of jsMap) {
      const wasmCell = wasmMap.get(addr);
      if (!wasmCell) {
        drift.push(`${addr}: Calculated in JS but MISSING in WASM`);
        continue;
      }
      if (JSON.stringify(jsCell.value) !== JSON.stringify(wasmCell.value)) {
        drift.push(
          `${addr}: JS=${JSON.stringify(jsCell.value)} WASM=${JSON.stringify(wasmCell.value)}`,
        );
      }
    }

    for (const addr of wasmMap.keys()) {
      if (!jsMap.has(addr)) {
        drift.push(`${addr}: Calculated in WASM but MISSING in JS`);
      }
    }

    return { js: jsResults, wasm: wasmResults, drift };
  }

  recalculateDirty(
    dirtyRegions: Array<{
      sheetName: string;
      rowStart: number;
      rowEnd: number;
      colStart: number;
      colEnd: number;
    }>,
  ): number[] {
    this.beginMutationCollection();
    let changedInputCount = 0;
    let formulaChangedCount = 0;
    let explicitChangedCount = 0;

    for (const region of dirtyRegions) {
      const sheet = this.workbook.getSheet(region.sheetName);
      if (!sheet) continue;

      for (let row = region.rowStart; row <= region.rowEnd; row += 1) {
        for (let col = region.colStart; col <= region.colEnd; col += 1) {
          const cellIndex = this.workbook.cellKeyToIndex.get(makeCellKey(sheet.id, row, col));
          if (cellIndex !== undefined) {
            changedInputCount = this.markInputChanged(cellIndex, changedInputCount);
            explicitChangedCount = this.markExplicitChanged(cellIndex, explicitChangedCount);
          }
        }
      }
    }

    const recalculated = this.reconcilePivotOutputs(
      this.recalculate(
        this.composeMutationRoots(changedInputCount, formulaChangedCount),
        this.changedInputBuffer.subarray(0, changedInputCount),
      ),
      false,
    );
    const changed = this.composeEventChanges(recalculated, explicitChangedCount);
    this.lastMetrics.batchId += 1;
    this.lastMetrics.changedInputCount = changedInputCount + formulaChangedCount;
    this.events.emit(
      {
        kind: "batch",
        invalidation: "cells",
        changedCellIndices: changed,
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        metrics: this.lastMetrics,
      },
      changed,
      (cellIndex) => this.workbook.getQualifiedAddress(cellIndex),
    );
    return Array.from(changed);
  }

  updateRowMetadata(
    sheetName: string,
    start: number,
    count: number,
    size: number | null,
    hidden: boolean | null,
  ): void {
    const existing = this.workbook.getRowMetadata(sheetName, start, count);
    if (existing?.size === size && existing.hidden === hidden) {
      return;
    }
    if (existing === undefined && size === null && hidden === null) {
      return;
    }
    this.executeLocalTransaction([
      { kind: "updateRowMetadata", sheetName, start, count, size, hidden },
    ]);
  }

  getRowMetadata(sheetName: string): WorkbookAxisMetadataRecord[] {
    return this.workbook.listRowMetadata(sheetName);
  }

  getRowAxisEntries(sheetName: string): WorkbookAxisEntrySnapshot[] {
    return this.workbook.listRowAxisEntries(sheetName);
  }

  insertRows(sheetName: string, start: number, count: number): void {
    if (count <= 0) {
      return;
    }
    this.executeLocalTransaction([{ kind: "insertRows", sheetName, start, count }]);
  }

  deleteRows(sheetName: string, start: number, count: number): void {
    if (count <= 0) {
      return;
    }
    this.executeLocalTransaction([{ kind: "deleteRows", sheetName, start, count }]);
  }

  moveRows(sheetName: string, start: number, count: number, target: number): void {
    if (count <= 0 || start === target) {
      return;
    }
    this.executeLocalTransaction([{ kind: "moveRows", sheetName, start, count, target }]);
  }

  updateColumnMetadata(
    sheetName: string,
    start: number,
    count: number,
    size: number | null,
    hidden: boolean | null,
  ): void {
    const existing = this.workbook.getColumnMetadata(sheetName, start, count);
    if (existing?.size === size && existing.hidden === hidden) {
      return;
    }
    if (existing === undefined && size === null && hidden === null) {
      return;
    }
    this.executeLocalTransaction([
      { kind: "updateColumnMetadata", sheetName, start, count, size, hidden },
    ]);
  }

  getColumnMetadata(sheetName: string): WorkbookAxisMetadataRecord[] {
    return this.workbook.listColumnMetadata(sheetName);
  }

  getColumnAxisEntries(sheetName: string): WorkbookAxisEntrySnapshot[] {
    return this.workbook.listColumnAxisEntries(sheetName);
  }

  insertColumns(sheetName: string, start: number, count: number): void {
    if (count <= 0) {
      return;
    }
    this.executeLocalTransaction([{ kind: "insertColumns", sheetName, start, count }]);
  }

  deleteColumns(sheetName: string, start: number, count: number): void {
    if (count <= 0) {
      return;
    }
    this.executeLocalTransaction([{ kind: "deleteColumns", sheetName, start, count }]);
  }

  moveColumns(sheetName: string, start: number, count: number, target: number): void {
    if (count <= 0 || start === target) {
      return;
    }
    this.executeLocalTransaction([{ kind: "moveColumns", sheetName, start, count, target }]);
  }

  setFreezePane(sheetName: string, rows: number, cols: number): void {
    const existing = this.workbook.getFreezePane(sheetName);
    if (existing?.rows === rows && existing.cols === cols) {
      return;
    }
    this.executeLocalTransaction([{ kind: "setFreezePane", sheetName, rows, cols }]);
  }

  clearFreezePane(sheetName: string): boolean {
    if (!this.workbook.getFreezePane(sheetName)) {
      return false;
    }
    this.executeLocalTransaction([{ kind: "clearFreezePane", sheetName }]);
    return true;
  }

  getFreezePane(sheetName: string): WorkbookFreezePaneSnapshot | undefined {
    return this.workbook.getFreezePane(sheetName);
  }

  setFilter(sheetName: string, range: CellRangeRef): void {
    const existing = this.workbook.getFilter(sheetName, range);
    if (existing) {
      return;
    }
    this.executeLocalTransaction([{ kind: "setFilter", sheetName, range: { ...range } }]);
  }

  clearFilter(sheetName: string, range: CellRangeRef): boolean {
    if (!this.workbook.getFilter(sheetName, range)) {
      return false;
    }
    this.executeLocalTransaction([{ kind: "clearFilter", sheetName, range: { ...range } }]);
    return true;
  }

  getFilters(sheetName: string): WorkbookFilterRecord[] {
    return this.workbook.listFilters(sheetName);
  }

  setSort(sheetName: string, range: CellRangeRef, keys: WorkbookSortSnapshot["keys"]): void {
    const existing = this.workbook.getSort(sheetName, range);
    const normalizedKeys = keys.map((key) => Object.assign({}, key));
    if (
      existing &&
      existing.keys.length === normalizedKeys.length &&
      existing.keys.every(
        (key, index) =>
          key.keyAddress === normalizedKeys[index]?.keyAddress &&
          key.direction === normalizedKeys[index]?.direction,
      )
    ) {
      return;
    }
    this.executeLocalTransaction([
      { kind: "setSort", sheetName, range: { ...range }, keys: normalizedKeys },
    ]);
  }

  clearSort(sheetName: string, range: CellRangeRef): boolean {
    if (!this.workbook.getSort(sheetName, range)) {
      return false;
    }
    this.executeLocalTransaction([{ kind: "clearSort", sheetName, range: { ...range } }]);
    return true;
  }

  getSorts(sheetName: string): WorkbookSortRecord[] {
    return this.workbook.listSorts(sheetName);
  }

  setTable(table: WorkbookTableRecord): void {
    const existing = this.workbook.getTable(table.name);
    if (
      existing &&
      existing.sheetName === table.sheetName &&
      existing.startAddress === table.startAddress &&
      existing.endAddress === table.endAddress &&
      existing.headerRow === table.headerRow &&
      existing.totalsRow === table.totalsRow &&
      existing.columnNames.length === table.columnNames.length &&
      existing.columnNames.every((name, index) => name === table.columnNames[index])
    ) {
      return;
    }
    this.executeLocalTransaction([
      {
        kind: "upsertTable",
        table: Object.assign({}, table, { columnNames: [...table.columnNames] }),
      },
    ]);
  }

  deleteTable(name: string): boolean {
    if (!this.workbook.getTable(name)) {
      return false;
    }
    this.executeLocalTransaction([{ kind: "deleteTable", name }]);
    return true;
  }

  getTable(name: string): WorkbookTableRecord | undefined {
    return this.workbook.getTable(name);
  }

  getTables(): WorkbookTableRecord[] {
    return this.workbook.listTables();
  }

  setSpillRange(sheetName: string, address: string, rows: number, cols: number): void {
    const existing = this.workbook.getSpill(sheetName, address);
    if (existing?.rows === rows && existing.cols === cols) {
      return;
    }
    this.executeLocalTransaction([{ kind: "upsertSpillRange", sheetName, address, rows, cols }]);
  }

  deleteSpillRange(sheetName: string, address: string): boolean {
    if (!this.workbook.getSpill(sheetName, address)) {
      return false;
    }
    this.executeLocalTransaction([{ kind: "deleteSpillRange", sheetName, address }]);
    return true;
  }

  getSpillRanges(): WorkbookSpillRecord[] {
    return this.workbook.listSpills();
  }

  setPivotTable(sheetName: string, address: string, definition: PivotTableInput): void {
    this.executeLocalTransaction([
      {
        kind: "upsertPivotTable",
        name: definition.name.trim(),
        sheetName,
        address,
        source: { ...definition.source },
        groupBy: [...definition.groupBy],
        values: definition.values.map((v) => Object.assign({}, v)),
        rows: 1,
        cols: Math.max(definition.groupBy.length + definition.values.length, 1),
      },
    ]);
  }

  deletePivotTable(sheetName: string, address: string): boolean {
    if (!this.workbook.getPivot(sheetName, address)) {
      return false;
    }
    this.executeLocalTransaction([{ kind: "deletePivotTable", sheetName, address }]);
    return true;
  }

  getPivotTable(sheetName: string, address: string): WorkbookPivotSnapshot | undefined {
    return this.workbook.getPivot(sheetName, address);
  }

  getPivotTables(): WorkbookPivotSnapshot[] {
    return this.workbook.listPivots();
  }

  clearCell(sheetName: string, address: string): void {
    this.executeLocalTransaction([{ kind: "clearCell", sheetName, address }]);
  }

  setRangeValues(range: CellRangeRef, values: readonly (readonly LiteralInput[])[]): void {
    const bounds = normalizeRange(range);
    const expectedHeight = bounds.endRow - bounds.startRow + 1;
    const expectedWidth = bounds.endCol - bounds.startCol + 1;
    if (values.length !== expectedHeight || values.some((row) => row.length !== expectedWidth)) {
      throw new Error(
        "setRangeValues requires a value matrix that exactly matches the target range",
      );
    }

    const opCount = expectedHeight * expectedWidth;
    const ops = Array.from<EngineOp>({ length: opCount });
    let opIndex = 0;
    for (let rowOffset = 0; rowOffset < expectedHeight; rowOffset += 1) {
      for (let colOffset = 0; colOffset < expectedWidth; colOffset += 1) {
        ops[opIndex] = {
          kind: "setCellValue",
          sheetName: range.sheetName,
          address: formatAddress(bounds.startRow + rowOffset, bounds.startCol + colOffset),
          value: values[rowOffset]![colOffset] ?? null,
        };
        opIndex += 1;
      }
    }
    this.executeLocalTransaction(ops, opCount);
  }

  setRangeFormulas(range: CellRangeRef, formulas: readonly (readonly string[])[]): void {
    const bounds = normalizeRange(range);
    const expectedHeight = bounds.endRow - bounds.startRow + 1;
    const expectedWidth = bounds.endCol - bounds.startCol + 1;
    if (
      formulas.length !== expectedHeight ||
      formulas.some((row) => row.length !== expectedWidth)
    ) {
      throw new Error(
        "setRangeFormulas requires a formula matrix that exactly matches the target range",
      );
    }

    const opCount = expectedHeight * expectedWidth;
    const ops = Array.from<EngineOp>({ length: opCount });
    let opIndex = 0;
    for (let rowOffset = 0; rowOffset < expectedHeight; rowOffset += 1) {
      for (let colOffset = 0; colOffset < expectedWidth; colOffset += 1) {
        ops[opIndex] = {
          kind: "setCellFormula",
          sheetName: range.sheetName,
          address: formatAddress(bounds.startRow + rowOffset, bounds.startCol + colOffset),
          formula: formulas[rowOffset]![colOffset] ?? "",
        };
        opIndex += 1;
      }
    }
    this.executeLocalTransaction(ops, opCount);
  }

  clearRange(range: CellRangeRef): void {
    const bounds = normalizeRange(range);
    const opCount = (bounds.endRow - bounds.startRow + 1) * (bounds.endCol - bounds.startCol + 1);
    const ops = Array.from<EngineOp>({ length: opCount });
    let opIndex = 0;
    for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
      for (let col = bounds.startCol; col <= bounds.endCol; col += 1) {
        ops[opIndex] = {
          kind: "clearCell",
          sheetName: range.sheetName,
          address: formatAddress(row, col),
        };
        opIndex += 1;
      }
    }
    this.executeLocalTransaction(ops, opCount);
  }

  fillRange(source: CellRangeRef, target: CellRangeRef): void {
    const sourceMatrix = this.readRangeCells(source);
    const targetBounds = normalizeRange(target);
    const sourceBounds = normalizeRange(source);
    const sourceHeight = sourceMatrix.length;
    const sourceWidth = sourceMatrix[0]?.length ?? 0;
    if (sourceHeight === 0 || sourceWidth === 0) {
      return;
    }

    const ops: EngineOp[] = [];
    for (let row = targetBounds.startRow; row <= targetBounds.endRow; row += 1) {
      for (let col = targetBounds.startCol; col <= targetBounds.endCol; col += 1) {
        const sourceRowOffset = (row - targetBounds.startRow) % sourceHeight;
        const sourceColOffset = (col - targetBounds.startCol) % sourceWidth;
        const sourceCell = sourceMatrix[sourceRowOffset]![sourceColOffset]!;
        const sourceAddress = formatAddress(
          sourceBounds.startRow + sourceRowOffset,
          sourceBounds.startCol + sourceColOffset,
        );
        ops.push(
          ...this.toCellStateOps(
            target.sheetName,
            formatAddress(row, col),
            sourceCell,
            source.sheetName,
            sourceAddress,
          ),
        );
      }
    }
    this.executeLocalTransaction(ops, ops.length);
  }

  copyRange(source: CellRangeRef, target: CellRangeRef): void {
    const sourceMatrix = this.readRangeCells(source);
    const targetBounds = normalizeRange(target);
    const sourceBounds = normalizeRange(source);
    const sourceHeight = sourceBounds.endRow - sourceBounds.startRow + 1;
    const sourceWidth = sourceBounds.endCol - sourceBounds.startCol + 1;
    const targetHeight = targetBounds.endRow - targetBounds.startRow + 1;
    const targetWidth = targetBounds.endCol - targetBounds.startCol + 1;
    if (sourceHeight !== targetHeight || sourceWidth !== targetWidth) {
      throw new Error("copyRange requires source and target dimensions to match exactly");
    }

    const ops: EngineOp[] = [];
    for (let rowOffset = 0; rowOffset < targetHeight; rowOffset += 1) {
      for (let colOffset = 0; colOffset < targetWidth; colOffset += 1) {
        const nextAddress = formatAddress(
          targetBounds.startRow + rowOffset,
          targetBounds.startCol + colOffset,
        );
        const sourceAddress = formatAddress(
          sourceBounds.startRow + rowOffset,
          sourceBounds.startCol + colOffset,
        );
        ops.push(
          ...this.toCellStateOps(
            target.sheetName,
            nextAddress,
            sourceMatrix[rowOffset]![colOffset]!,
            source.sheetName,
            sourceAddress,
          ),
        );
      }
    }
    this.executeLocalTransaction(ops, ops.length);
  }

  moveRange(source: CellRangeRef, target: CellRangeRef): void {
    const sourceMatrix = this.readRangeCells(source);
    const targetBounds = normalizeRange(target);
    const sourceBounds = normalizeRange(source);
    const sourceHeight = sourceBounds.endRow - sourceBounds.startRow + 1;
    const sourceWidth = sourceBounds.endCol - sourceBounds.startCol + 1;
    const targetHeight = targetBounds.endRow - targetBounds.startRow + 1;
    const targetWidth = targetBounds.endCol - targetBounds.startCol + 1;
    if (sourceHeight !== targetHeight || sourceWidth !== targetWidth) {
      throw new Error("moveRange requires source and target dimensions to match exactly");
    }

    const ops: EngineOp[] = [];
    for (let row = sourceBounds.startRow; row <= sourceBounds.endRow; row += 1) {
      for (let col = sourceBounds.startCol; col <= sourceBounds.endCol; col += 1) {
        ops.push({
          kind: "clearCell",
          sheetName: source.sheetName,
          address: formatAddress(row, col),
        });
      }
    }
    for (let rowOffset = 0; rowOffset < targetHeight; rowOffset += 1) {
      for (let colOffset = 0; colOffset < targetWidth; colOffset += 1) {
        const nextAddress = formatAddress(
          targetBounds.startRow + rowOffset,
          targetBounds.startCol + colOffset,
        );
        const sourceAddress = formatAddress(
          sourceBounds.startRow + rowOffset,
          sourceBounds.startCol + colOffset,
        );
        ops.push(
          ...this.toCellStateOps(
            target.sheetName,
            nextAddress,
            sourceMatrix[rowOffset]![colOffset]!,
            source.sheetName,
            sourceAddress,
          ),
        );
      }
    }
    this.executeLocalTransaction(ops, ops.length);
  }

  pasteRange(source: CellRangeRef, target: CellRangeRef): void {
    this.copyRange(source, target);
  }

  undo(): boolean {
    const entry = this.undoStack.pop();
    if (!entry) {
      return false;
    }
    this.transactionReplayDepth += 1;
    try {
      this.executeTransaction(entry.inverse, "history");
    } finally {
      this.transactionReplayDepth -= 1;
    }
    this.redoStack.push(entry);
    return true;
  }

  redo(): boolean {
    const entry = this.redoStack.pop();
    if (!entry) {
      return false;
    }
    this.transactionReplayDepth += 1;
    try {
      this.executeTransaction(entry.forward, "history");
    } finally {
      this.transactionReplayDepth -= 1;
    }
    this.undoStack.push(entry);
    return true;
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

    const rows = Array.from({ length: maxRow + 1 }, (_rowEntry, row) =>
      Array.from({ length: maxCol + 1 }, (_colEntry, col) => cells.get(`${row}:${col}`) ?? ""),
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

    this.executeLocalTransaction(ops, potentialNewCells);
  }

  getCellValue(sheetName: string, address: string): CellValue {
    const cellIndex = this.workbook.getCellIndex(sheetName, address);
    if (cellIndex === undefined) {
      return emptyValue();
    }
    return this.workbook.cellStore.getValue(cellIndex, (id) => this.strings.get(id));
  }

  getRangeValues(range: CellRangeRef): CellValue[][] {
    return this.readRangeValueMatrix(range);
  }

  getCell(sheetName: string, address: string): CellSnapshot {
    const cellIndex = this.workbook.getCellIndex(sheetName, address);
    if (cellIndex === undefined) {
      const parsed = parseCellAddress(address, sheetName);
      const styleId = this.workbook.getStyleId(sheetName, parsed.row, parsed.col);
      const numberFormatId = this.workbook.getRangeFormatId(sheetName, parsed.row, parsed.col);
      const formatRecord = this.workbook.getCellNumberFormat(numberFormatId);
      return {
        sheetName,
        address,
        ...(styleId !== WorkbookStore.defaultStyleId ? { styleId } : {}),
        ...(numberFormatId !== WorkbookStore.defaultFormatId ? { numberFormatId } : {}),
        ...(formatRecord && numberFormatId !== WorkbookStore.defaultFormatId
          ? { format: formatRecord.code }
          : {}),
        value: emptyValue(),
        flags: 0,
        version: 0,
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
      version: this.workbook.cellStore.versions[cellIndex] ?? 0,
    };
    const styleId = this.workbook.getStyleId(
      sheetName,
      this.workbook.cellStore.rows[cellIndex]!,
      this.workbook.cellStore.cols[cellIndex]!,
    );
    if (styleId !== WorkbookStore.defaultStyleId) {
      snapshot.styleId = styleId;
    }
    const explicitFormat = this.workbook.getCellFormat(cellIndex);
    const numberFormatId =
      explicitFormat !== undefined
        ? this.workbook.internCellNumberFormat(explicitFormat).id
        : this.workbook.getRangeFormatId(
            sheetName,
            this.workbook.cellStore.rows[cellIndex]!,
            this.workbook.cellStore.cols[cellIndex]!,
          );
    const formatRecord = this.workbook.getCellNumberFormat(numberFormatId);
    if (numberFormatId !== WorkbookStore.defaultFormatId) {
      snapshot.numberFormatId = numberFormatId;
    }
    if (explicitFormat !== undefined) {
      snapshot.format = explicitFormat;
    } else if (formatRecord && numberFormatId !== WorkbookStore.defaultFormatId) {
      snapshot.format = formatRecord.code;
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
        this.workbook.getQualifiedAddress(dependencyCellIndex),
      ),
      directDependents: [...directDependents].map((dependentCellIndex) =>
        this.workbook.getQualifiedAddress(dependentCellIndex),
      ),
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
        directDependents: [],
      };
    }

    const snapshot = this.getCellByIndex(cellIndex);
    const formula = this.formulas.get(cellIndex);
    const flags = this.workbook.cellStore.flags[cellIndex] ?? 0;
    const isFormula = (flags & CellFlags.HasFormula) !== 0 && formula !== undefined;
    const dependencies = this.getDependencies(sheetName, address);

    const explanation: ExplainCellSnapshot = {
      ...snapshot,
      version: this.workbook.cellStore.versions[cellIndex] ?? 0,
      inCycle: (flags & CellFlags.InCycle) !== 0,
      directPrecedents: dependencies.directPrecedents,
      directDependents: dependencies.directDependents,
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
    const workbook: WorkbookSnapshot["workbook"] = {
      name: this.workbook.workbookName,
    };
    const properties = this.workbook
      .listWorkbookProperties()
      .map(({ key, value }) => ({ key, value }));
    const definedNames = this.workbook
      .listDefinedNames()
      .map(({ name, value }) => ({ name, value }));
    const calculationSettings = this.workbook.getCalculationSettings();
    const volatileContext = this.workbook.getVolatileContext();
    const tables = this.workbook.listTables().map((table) => ({
      name: table.name,
      sheetName: table.sheetName,
      startAddress: table.startAddress,
      endAddress: table.endAddress,
      columnNames: [...table.columnNames],
      headerRow: table.headerRow,
      totalsRow: table.totalsRow,
    }));
    const spills = this.workbook
      .listSpills()
      .map(({ sheetName, address, rows, cols }) => ({ sheetName, address, rows, cols }));
    const referencedStyleIds = new Set<string>();
    const referencedFormatIds = new Set<string>();
    this.workbook.sheetsByName.forEach((sheet) => {
      sheet.styleRanges.forEach((record) => referencedStyleIds.add(record.styleId));
      sheet.formatRanges.forEach((record) => referencedFormatIds.add(record.formatId));
    });
    for (let cellIndex = 0; cellIndex < this.workbook.cellStore.size; cellIndex += 1) {
      const explicitFormat = this.workbook.getCellFormat(cellIndex);
      if (explicitFormat !== undefined) {
        referencedFormatIds.add(this.workbook.internCellNumberFormat(explicitFormat).id);
      }
    }
    const styles = this.workbook
      .listCellStyles()
      .filter((style) => referencedStyleIds.has(style.id))
      .map((style) => cloneCellStyleRecord(style));
    const formats = this.workbook
      .listCellNumberFormats()
      .filter((format) => referencedFormatIds.has(format.id))
      .map((format) => Object.assign({}, format));
    const pivots = this.workbook.listPivots().map((pivot) => ({
      name: pivot.name,
      sheetName: pivot.sheetName,
      address: pivot.address,
      source: { ...pivot.source },
      groupBy: [...pivot.groupBy],
      values: pivot.values.map((value) => Object.assign({}, value)),
      rows: pivot.rows,
      cols: pivot.cols,
    }));
    if (
      properties.length > 0 ||
      definedNames.length > 0 ||
      tables.length > 0 ||
      spills.length > 0 ||
      pivots.length > 0 ||
      styles.length > 0 ||
      formats.length > 0 ||
      calculationSettings.mode !== "automatic" ||
      calculationSettings.compatibilityMode !== "excel-modern" ||
      volatileContext.recalcEpoch !== 0
    ) {
      workbook.metadata = {};
      if (properties.length > 0) {
        workbook.metadata.properties = properties;
      }
      if (definedNames.length > 0) {
        workbook.metadata.definedNames = definedNames;
      }
      if (tables.length > 0) {
        workbook.metadata.tables = tables;
      }
      if (spills.length > 0) {
        workbook.metadata.spills = spills;
      }
      if (pivots.length > 0) {
        workbook.metadata.pivots = pivots;
      }
      if (styles.length > 0) {
        workbook.metadata.styles = styles;
      }
      if (formats.length > 0) {
        workbook.metadata.formats = formats;
      }
      if (
        calculationSettings.mode !== "automatic" ||
        calculationSettings.compatibilityMode !== "excel-modern"
      ) {
        workbook.metadata.calculationSettings = calculationSettings;
      }
      if (volatileContext.recalcEpoch !== 0) {
        workbook.metadata.volatileContext = volatileContext;
      }
    }

    return {
      version: 1,
      workbook,
      sheets: [...this.workbook.sheetsByName.values()]
        .toSorted((left, right) => left.order - right.order)
        .map((sheet) => {
          const metadata = this.exportSheetMetadata(sheet.name);
          const cells: WorkbookSnapshot["sheets"][number]["cells"] = [];
          sheet.grid.forEachCell((cellIndex) => {
            const snapshot = this.getCellByIndex(cellIndex);
            if ((snapshot.flags & (CellFlags.SpillChild | CellFlags.PivotOutput)) !== 0) {
              return;
            }
            const cell: WorkbookSnapshot["sheets"][number]["cells"][number] = {
              address: snapshot.address,
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
          return metadata
            ? { id: sheet.id, name: sheet.name, order: sheet.order, metadata, cells }
            : { id: sheet.id, name: sheet.name, order: sheet.order, cells };
        }),
    };
  }

  importSnapshot(snapshot: WorkbookSnapshot): void {
    this.resetWorkbook();
    const ops: EngineOp[] = [{ kind: "upsertWorkbook", name: snapshot.workbook.name }];
    snapshot.workbook.metadata?.properties?.forEach(({ key, value }) => {
      ops.push({ kind: "setWorkbookMetadata", key, value });
    });
    if (snapshot.workbook.metadata?.calculationSettings) {
      ops.push({
        kind: "setCalculationSettings",
        settings: { ...snapshot.workbook.metadata.calculationSettings },
      });
    }
    if (snapshot.workbook.metadata?.volatileContext) {
      ops.push({
        kind: "setVolatileContext",
        context: { ...snapshot.workbook.metadata.volatileContext },
      });
    }
    snapshot.workbook.metadata?.definedNames?.forEach(({ name, value }) => {
      ops.push({ kind: "upsertDefinedName", name, value });
    });
    snapshot.workbook.metadata?.styles?.forEach((style) => {
      ops.push({ kind: "upsertCellStyle", style: cloneCellStyleRecord(style) });
    });
    snapshot.workbook.metadata?.formats?.forEach((format) => {
      ops.push({ kind: "upsertCellNumberFormat", format: { ...format } });
    });
    snapshot.sheets.forEach((sheet) => {
      ops.push({
        kind: "upsertSheet",
        name: sheet.name,
        order: sheet.order,
        ...(typeof sheet.id === "number" ? { id: sheet.id } : {}),
      });
    });
    snapshot.sheets.forEach((sheet) => {
      sheet.metadata?.rows?.forEach(({ index, id, size, hidden }) => {
        const entry = { index, id } as WorkbookAxisEntrySnapshot;
        if (size !== undefined) {
          entry.size = size;
        }
        if (hidden !== undefined) {
          entry.hidden = hidden;
        }
        ops.push({
          kind: "insertRows",
          sheetName: sheet.name,
          start: index,
          count: 1,
          entries: [entry],
        });
      });
      sheet.metadata?.columns?.forEach(({ index, id, size, hidden }) => {
        const entry = { index, id } as WorkbookAxisEntrySnapshot;
        if (size !== undefined) {
          entry.size = size;
        }
        if (hidden !== undefined) {
          entry.hidden = hidden;
        }
        ops.push({
          kind: "insertColumns",
          sheetName: sheet.name,
          start: index,
          count: 1,
          entries: [entry],
        });
      });
      sheet.metadata?.rowMetadata?.forEach(({ start, count, size, hidden }) => {
        ops.push({
          kind: "updateRowMetadata",
          sheetName: sheet.name,
          start,
          count,
          size: size ?? null,
          hidden: hidden ?? null,
        });
      });
      sheet.metadata?.columnMetadata?.forEach(({ start, count, size, hidden }) => {
        ops.push({
          kind: "updateColumnMetadata",
          sheetName: sheet.name,
          start,
          count,
          size: size ?? null,
          hidden: hidden ?? null,
        });
      });
      if (sheet.metadata?.freezePane) {
        ops.push({
          kind: "setFreezePane",
          sheetName: sheet.name,
          rows: sheet.metadata.freezePane.rows,
          cols: sheet.metadata.freezePane.cols,
        });
      }
      sheet.metadata?.styleRanges?.forEach((styleRange) => {
        ops.push({
          kind: "setStyleRange",
          range: { ...styleRange.range },
          styleId: styleRange.styleId,
        });
      });
      sheet.metadata?.formatRanges?.forEach((formatRange) => {
        ops.push({
          kind: "setFormatRange",
          range: { ...formatRange.range },
          formatId: formatRange.formatId,
        });
      });
      sheet.metadata?.filters?.forEach((range) => {
        ops.push({ kind: "setFilter", sheetName: sheet.name, range: { ...range } });
      });
      sheet.metadata?.sorts?.forEach((sort) => {
        ops.push({
          kind: "setSort",
          sheetName: sheet.name,
          range: { ...sort.range },
          keys: sort.keys.map((key) => Object.assign({}, key)),
        });
      });
    });
    snapshot.sheets.forEach((sheet) => {
      sheet.cells.forEach((cell) => {
        if (cell.formula !== undefined) {
          ops.push({
            kind: "setCellFormula",
            sheetName: sheet.name,
            address: cell.address,
            formula: cell.formula,
          });
        } else {
          ops.push({
            kind: "setCellValue",
            sheetName: sheet.name,
            address: cell.address,
            value: cell.value ?? null,
          });
        }
        if (cell.format !== undefined) {
          ops.push({
            kind: "setCellFormat",
            sheetName: sheet.name,
            address: cell.address,
            format: cell.format,
          });
        }
      });
    });
    snapshot.workbook.metadata?.tables?.forEach((table) => {
      ops.push({
        kind: "upsertTable",
        table: {
          name: table.name,
          sheetName: table.sheetName,
          startAddress: table.startAddress,
          endAddress: table.endAddress,
          columnNames: [...table.columnNames],
          headerRow: table.headerRow,
          totalsRow: table.totalsRow,
        },
      });
    });
    snapshot.workbook.metadata?.spills?.forEach((spill) => {
      ops.push({
        kind: "upsertSpillRange",
        sheetName: spill.sheetName,
        address: spill.address,
        rows: spill.rows,
        cols: spill.cols,
      });
    });
    snapshot.workbook.metadata?.pivots?.forEach((pivot) => {
      ops.push({
        kind: "upsertPivotTable",
        name: pivot.name,
        sheetName: pivot.sheetName,
        address: pivot.address,
        source: { ...pivot.source },
        groupBy: [...pivot.groupBy],
        values: pivot.values.map((value) => Object.assign({}, value)),
        rows: pivot.rows,
        cols: pivot.cols,
      });
    });
    const potentialNewCells = snapshot.sheets.reduce(
      (count, sheet) => count + sheet.cells.length,
      0,
    );
    this.executeTransaction(
      potentialNewCells > 0 ? { ops, potentialNewCells } : { ops },
      "restore",
    );
  }

  private exportSheetMetadata(sheetName: string): SheetMetadataSnapshot | undefined {
    const rows = this.workbook.listRowAxisEntries(sheetName);
    const columns = this.workbook.listColumnAxisEntries(sheetName);
    const rowMetadata = this.axisMetadataToSnapshot(this.workbook.listRowMetadata(sheetName));
    const columnMetadata = this.axisMetadataToSnapshot(this.workbook.listColumnMetadata(sheetName));
    const styleRanges = this.workbook.listStyleRanges(sheetName).map((record) => ({
      range: this.toSnapshotRangeRef(record.range),
      styleId: record.styleId,
    }));
    const formatRanges = this.workbook.listFormatRanges(sheetName).map((record) => ({
      range: this.toSnapshotRangeRef(record.range),
      formatId: record.formatId,
    }));
    const freezePane = this.freezePaneToSnapshot(this.workbook.getFreezePane(sheetName));
    const filters = this.workbook
      .listFilters(sheetName)
      .map((filter) => Object.assign({}, filter.range));
    const sorts = this.workbook.listSorts(sheetName).map((sort) => ({
      range: Object.assign({}, sort.range),
      keys: sort.keys.map((key) => Object.assign({}, key)),
    }));

    if (
      rows.length === 0 &&
      columns.length === 0 &&
      rowMetadata.length === 0 &&
      columnMetadata.length === 0 &&
      styleRanges.length === 0 &&
      formatRanges.length === 0 &&
      freezePane === undefined &&
      filters.length === 0 &&
      sorts.length === 0
    ) {
      return undefined;
    }

    const metadata: SheetMetadataSnapshot = {};
    if (rows.length > 0) {
      metadata.rows = rows;
    }
    if (columns.length > 0) {
      metadata.columns = columns;
    }
    if (rowMetadata.length > 0) {
      metadata.rowMetadata = rowMetadata;
    }
    if (columnMetadata.length > 0) {
      metadata.columnMetadata = columnMetadata;
    }
    if (styleRanges.length > 0) {
      metadata.styleRanges = styleRanges;
    }
    if (formatRanges.length > 0) {
      metadata.formatRanges = formatRanges;
    }
    if (freezePane) {
      metadata.freezePane = freezePane;
    }
    if (filters.length > 0) {
      metadata.filters = filters;
    }
    if (sorts.length > 0) {
      metadata.sorts = sorts;
    }
    return metadata;
  }

  private toSnapshotRangeRef(range: CellRangeRef): CellRangeRef {
    return {
      sheetName: range.sheetName,
      startAddress: range.startAddress,
      endAddress: range.endAddress,
    };
  }

  private sheetMetadataToOps(sheetName: string): EngineOp[] {
    const ops: EngineOp[] = [];
    this.workbook.listRowAxisEntries(sheetName).forEach((entry) => {
      ops.push({ kind: "insertRows", sheetName, start: entry.index, count: 1, entries: [entry] });
    });
    this.workbook.listColumnAxisEntries(sheetName).forEach((entry) => {
      ops.push({
        kind: "insertColumns",
        sheetName,
        start: entry.index,
        count: 1,
        entries: [entry],
      });
    });
    this.workbook.listRowMetadata(sheetName).forEach((record) => {
      ops.push({
        kind: "updateRowMetadata",
        sheetName,
        start: record.start,
        count: record.count,
        size: record.size,
        hidden: record.hidden,
      });
    });
    this.workbook.listColumnMetadata(sheetName).forEach((record) => {
      ops.push({
        kind: "updateColumnMetadata",
        sheetName,
        start: record.start,
        count: record.count,
        size: record.size,
        hidden: record.hidden,
      });
    });
    this.workbook.listStyleRanges(sheetName).forEach((record) => {
      ops.push({ kind: "setStyleRange", range: { ...record.range }, styleId: record.styleId });
    });
    this.workbook.listFormatRanges(sheetName).forEach((record) => {
      ops.push({ kind: "setFormatRange", range: { ...record.range }, formatId: record.formatId });
    });
    const freezePane = this.workbook.getFreezePane(sheetName);
    if (freezePane) {
      ops.push({ kind: "setFreezePane", sheetName, rows: freezePane.rows, cols: freezePane.cols });
    }
    this.workbook.listFilters(sheetName).forEach((record) => {
      ops.push({ kind: "setFilter", sheetName, range: { ...record.range } });
    });
    this.workbook.listSorts(sheetName).forEach((record) => {
      ops.push({
        kind: "setSort",
        sheetName,
        range: { ...record.range },
        keys: record.keys.map((key) => Object.assign({}, key)),
      });
    });
    return ops;
  }

  private axisMetadataToSnapshot(
    records: readonly WorkbookAxisMetadataRecord[],
  ): WorkbookAxisMetadataSnapshot[] {
    return records.map((record) => {
      const snapshot: WorkbookAxisMetadataSnapshot = {
        start: record.start,
        count: record.count,
      };
      if (record.size !== null) {
        snapshot.size = record.size;
      }
      if (record.hidden !== null) {
        snapshot.hidden = record.hidden;
      }
      return snapshot;
    });
  }

  private freezePaneToSnapshot(
    record: { rows: number; cols: number } | undefined,
  ): WorkbookFreezePaneSnapshot | undefined {
    if (!record) {
      return undefined;
    }
    return { rows: record.rows, cols: record.cols };
  }

  private captureRowRangeCellState(sheetName: string, start: number, count: number): EngineOp[] {
    return this.captureAxisRangeCellState(sheetName, "row", start, count);
  }

  private captureColumnRangeCellState(sheetName: string, start: number, count: number): EngineOp[] {
    return this.captureAxisRangeCellState(sheetName, "column", start, count);
  }

  private captureAxisRangeCellState(
    sheetName: string,
    axis: "row" | "column",
    start: number,
    count: number,
  ): EngineOp[] {
    const sheet = this.workbook.getSheet(sheetName);
    if (!sheet) {
      return [];
    }
    const captured: Array<{ cellIndex: number; row: number; col: number }> = [];
    sheet.grid.forEachCellEntry((cellIndex, row, col) => {
      const index = axis === "row" ? row : col;
      if (index >= start && index < start + count) {
        captured.push({ cellIndex, row, col });
      }
    });
    return captured
      .toSorted((left, right) => left.row - right.row || left.col - right.col)
      .flatMap(({ cellIndex, row, col }) =>
        this.toCellStateOps(sheetName, formatAddress(row, col), this.getCellByIndex(cellIndex)),
      );
  }

  private applyStructuralAxisOp(
    op: Extract<
      EngineOp,
      {
        kind:
          | "insertRows"
          | "deleteRows"
          | "moveRows"
          | "insertColumns"
          | "deleteColumns"
          | "moveColumns";
      }
    >,
  ): { changedCellIndices: number[]; formulaCellIndices: number[] } {
    const axis = op.kind.includes("Rows") ? "row" : "column";
    const transform = structuralTransformForOp(op);
    const sheetName = op.sheetName;

    this.rewriteDefinedNamesForStructuralTransform(sheetName, transform);
    this.rewriteCellFormulasForStructuralTransform(sheetName, transform);
    this.rewriteWorkbookMetadataForStructuralTransform(sheetName, transform);

    switch (op.kind) {
      case "insertRows":
        this.workbook.insertRows(sheetName, op.start, op.count, op.entries);
        break;
      case "deleteRows":
        this.workbook.deleteRows(sheetName, op.start, op.count);
        break;
      case "moveRows":
        this.workbook.moveRows(sheetName, op.start, op.count, op.target);
        break;
      case "insertColumns":
        this.workbook.insertColumns(sheetName, op.start, op.count, op.entries);
        break;
      case "deleteColumns":
        this.workbook.deleteColumns(sheetName, op.start, op.count);
        break;
      case "moveColumns":
        this.workbook.moveColumns(sheetName, op.start, op.count, op.target);
        break;
      default:
        return assertNever(op);
    }

    const remapped = this.workbook.remapSheetCells(sheetName, axis, (index) =>
      mapStructuralAxisIndex(index, transform),
    );
    remapped.removedCellIndices.forEach((cellIndex) => {
      this.clearDerivedCellArtifacts(cellIndex);
      this.removeFormula(cellIndex);
      this.workbook.setCellFormat(cellIndex, null);
      this.workbook.cellStore.setValue(cellIndex, emptyValue());
      this.workbook.cellStore.flags[cellIndex] =
        (this.workbook.cellStore.flags[cellIndex] ?? 0) &
        ~(
          CellFlags.HasFormula |
          CellFlags.JsOnly |
          CellFlags.InCycle |
          CellFlags.SpillChild |
          CellFlags.PivotOutput
        );
    });

    this.clearAllSpillMetadata();
    this.clearPivotOutputsForSheet(sheetName);
    const formulaCellIndices = this.rebuildAllFormulaBindings();
    return {
      changedCellIndices: [...remapped.changedCellIndices, ...remapped.removedCellIndices],
      formulaCellIndices,
    };
  }

  private rewriteDefinedNamesForStructuralTransform(
    sheetName: string,
    transform: StructuralAxisTransform,
  ): void {
    this.workbook.listDefinedNames().forEach((record) => {
      if (typeof record.value !== "string" || !record.value.startsWith("=")) {
        return;
      }
      const nextFormula = rewriteFormulaForStructuralTransform(
        record.value.slice(1),
        sheetName,
        sheetName,
        transform,
      );
      if (`=${nextFormula}` !== record.value) {
        this.workbook.setDefinedName(record.name, `=${nextFormula}`);
      }
    });
  }

  private rewriteDefinedNamesForSheetRename(oldSheetName: string, newSheetName: string): void {
    this.workbook.listDefinedNames().forEach((record) => {
      const nextValue = renameDefinedNameValueSheet(record.value, oldSheetName, newSheetName);
      if (!definedNameValuesEqual(record.value, nextValue)) {
        this.workbook.setDefinedName(record.name, nextValue);
      }
    });
  }

  private rewriteCellFormulasForStructuralTransform(
    sheetName: string,
    transform: StructuralAxisTransform,
  ): void {
    this.formulas.forEach((formula, cellIndex) => {
      const ownerSheetName = this.workbook.getSheetNameById(
        this.workbook.cellStore.sheetIds[cellIndex]!,
      );
      formula.source = rewriteFormulaForStructuralTransform(
        formula.source,
        ownerSheetName,
        sheetName,
        transform,
      );
    });
  }

  private rewriteCellFormulasForSheetRename(
    oldSheetName: string,
    newSheetName: string,
    formulaChangedCount: number,
  ): number {
    this.formulas.forEach((formula, cellIndex) => {
      if (!formula) {
        return;
      }
      const ownerSheetName = this.workbook.getSheetNameById(
        this.workbook.cellStore.sheetIds[cellIndex]!,
      );
      if (!ownerSheetName) {
        return;
      }
      const nextSource = renameFormulaSheetReferences(formula.source, oldSheetName, newSheetName);
      if (nextSource === formula.source && ownerSheetName !== newSheetName) {
        return;
      }
      const compiled = this.compileFormulaForSheet(ownerSheetName, nextSource);
      const dependencies = this.materializeDependencies(ownerSheetName, compiled);
      this.setFormula(cellIndex, nextSource, compiled, dependencies);
      formulaChangedCount = this.markFormulaChanged(cellIndex, formulaChangedCount);
    });
    return formulaChangedCount;
  }

  private rewriteWorkbookMetadataForStructuralTransform(
    sheetName: string,
    transform: StructuralAxisTransform,
  ): void {
    this.workbook
      .listTables()
      .filter((table) => table.sheetName === sheetName)
      .forEach((table) => {
        const range = rewriteRangeForStructuralTransform(
          table.startAddress,
          table.endAddress,
          transform,
        );
        if (!range) {
          this.workbook.deleteTable(table.name);
          return;
        }
        this.workbook.setTable({
          ...table,
          startAddress: range.startAddress,
          endAddress: range.endAddress,
        });
      });
    this.workbook.listFilters(sheetName).forEach((filter) => {
      const range = rewriteRangeForStructuralTransform(
        filter.range.startAddress,
        filter.range.endAddress,
        transform,
      );
      this.workbook.deleteFilter(sheetName, filter.range);
      if (range) {
        this.workbook.setFilter(sheetName, {
          ...filter.range,
          startAddress: range.startAddress,
          endAddress: range.endAddress,
        });
      }
    });
    this.workbook.listSorts(sheetName).forEach((sort) => {
      const range = rewriteRangeForStructuralTransform(
        sort.range.startAddress,
        sort.range.endAddress,
        transform,
      );
      this.workbook.deleteSort(sheetName, sort.range);
      if (!range) {
        return;
      }
      this.workbook.setSort(
        sheetName,
        { ...sort.range, startAddress: range.startAddress, endAddress: range.endAddress },
        sort.keys.map((key) => ({
          ...key,
          keyAddress:
            rewriteAddressForStructuralTransform(key.keyAddress, transform) ?? key.keyAddress,
        })),
      );
    });
    const rewrittenStyleRanges: SheetStyleRangeSnapshot[] = [];
    const rewrittenFormatRanges: SheetFormatRangeSnapshot[] = [];
    this.workbook.listStyleRanges(sheetName).forEach((record) => {
      const range = rewriteRangeForStructuralTransform(
        record.range.startAddress,
        record.range.endAddress,
        transform,
      );
      if (!range) {
        return;
      }
      rewrittenStyleRanges.push({
        range: {
          ...record.range,
          startAddress: range.startAddress,
          endAddress: range.endAddress,
        },
        styleId: record.styleId,
      });
    });
    this.workbook.setStyleRanges(sheetName, rewrittenStyleRanges);
    this.workbook.listFormatRanges(sheetName).forEach((record) => {
      const range = rewriteRangeForStructuralTransform(
        record.range.startAddress,
        record.range.endAddress,
        transform,
      );
      if (!range) {
        return;
      }
      rewrittenFormatRanges.push({
        range: {
          ...record.range,
          startAddress: range.startAddress,
          endAddress: range.endAddress,
        },
        formatId: record.formatId,
      });
    });
    this.workbook.setFormatRanges(sheetName, rewrittenFormatRanges);
    const freezePane = this.workbook.getFreezePane(sheetName);
    if (freezePane) {
      const nextRows =
        transform.axis === "row"
          ? mapStructuralBoundary(freezePane.rows, transform)
          : freezePane.rows;
      const nextCols =
        transform.axis === "column"
          ? mapStructuralBoundary(freezePane.cols, transform)
          : freezePane.cols;
      if (nextRows <= 0 && nextCols <= 0) {
        this.workbook.clearFreezePane(sheetName);
      } else {
        this.workbook.setFreezePane(sheetName, nextRows, nextCols);
      }
    }
    this.workbook.listPivots().forEach((pivot) => {
      const nextAddress =
        pivot.sheetName === sheetName
          ? rewriteAddressForStructuralTransform(pivot.address, transform)
          : pivot.address;
      const nextSource =
        pivot.source.sheetName === sheetName
          ? rewriteRangeForStructuralTransform(
              pivot.source.startAddress,
              pivot.source.endAddress,
              transform,
            )
          : { startAddress: pivot.source.startAddress, endAddress: pivot.source.endAddress };
      if (!nextAddress || !nextSource) {
        this.clearOwnedPivot(pivot);
        this.workbook.deletePivot(pivot.sheetName, pivot.address);
        return;
      }
      this.workbook.setPivot({
        ...pivot,
        address: nextAddress,
        source: {
          ...pivot.source,
          startAddress: nextSource.startAddress,
          endAddress: nextSource.endAddress,
        },
      });
    });
  }

  private clearAllSpillMetadata(): void {
    this.workbook.listSpills().forEach((spill) => {
      this.workbook.deleteSpill(spill.sheetName, spill.address);
    });
  }

  private clearPivotOutputsForSheet(sheetName: string): void {
    this.workbook
      .listPivots()
      .filter((pivot) => pivot.sheetName === sheetName)
      .forEach((pivot) => {
        this.clearOwnedPivot(pivot);
      });
  }

  private clearDerivedCellArtifacts(cellIndex: number): void {
    this.pivotOutputOwners.delete(cellIndex);
  }

  private rebuildAllFormulaBindings(): number[] {
    const pending = [...this.formulas.entries()].map(([cellIndex, formula]) => ({
      cellIndex,
      source: formula.source,
    }));
    this.formulas.clear();
    this.ranges.reset();
    this.edgeArena.reset();
    this.programArena.reset();
    this.constantArena.reset();
    this.rangeListArena.reset();
    this.reverseCellEdges = [];
    this.reverseRangeEdges = [];
    this.reverseDefinedNameEdges.clear();
    this.reverseTableEdges.clear();
    this.reverseSpillEdges.clear();

    const activeCellIndices: number[] = [];
    pending.forEach(({ cellIndex, source }) => {
      const ownerSheetName = this.workbook.getSheetNameById(
        this.workbook.cellStore.sheetIds[cellIndex]!,
      );
      if (!ownerSheetName || !this.workbook.getSheet(ownerSheetName)) {
        return;
      }
      try {
        const compiled = this.compileFormulaForSheet(ownerSheetName, source);
        const dependencies = this.materializeDependencies(ownerSheetName, compiled);
        this.setFormula(cellIndex, source, compiled, dependencies);
      } catch {
        this.setInvalidFormulaValue(cellIndex);
      }
      activeCellIndices.push(cellIndex);
    });
    return activeCellIndices;
  }

  private collectTrackedDependents(
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

  private rebindFormulaCells(candidates: readonly number[], formulaChangedCount: number): number {
    candidates.forEach((cellIndex) => {
      const formula = this.formulas.get(cellIndex);
      const ownerSheetName = this.workbook.getSheetNameById(
        this.workbook.cellStore.sheetIds[cellIndex]!,
      );
      if (formula && ownerSheetName) {
        const compiled = this.compileFormulaForSheet(ownerSheetName, formula.source);
        const dependencies = this.materializeDependencies(ownerSheetName, compiled);
        this.setFormula(cellIndex, formula.source, compiled, dependencies);
      }
      formulaChangedCount = this.markFormulaChanged(cellIndex, formulaChangedCount);
    });
    return formulaChangedCount;
  }

  private rebindTrackedDependents(
    registry: Map<string, Set<number>>,
    keys: readonly string[],
    formulaChangedCount: number,
  ): number {
    return this.rebindFormulaCells(
      this.collectTrackedDependents(registry, keys),
      formulaChangedCount,
    );
  }

  private rebindDefinedNameDependents(
    names: readonly string[],
    formulaChangedCount: number,
  ): number {
    return this.rebindTrackedDependents(this.reverseDefinedNameEdges, names, formulaChangedCount);
  }

  private rebindTableDependents(
    tableNames: readonly string[],
    formulaChangedCount: number,
  ): number {
    return this.rebindTrackedDependents(this.reverseTableEdges, tableNames, formulaChangedCount);
  }

  private reconcilePivotOutputs(baseChanged: U32, forceAllPivots = false): U32 {
    let aggregate = baseChanged;
    let pending = baseChanged;
    let forceAll = forceAllPivots;

    for (let iteration = 0; iteration < 4; iteration += 1) {
      const pivotChanged = this.refreshPivotOutputs(pending, forceAll);
      if (pivotChanged.length === 0) {
        break;
      }
      aggregate =
        aggregate.length === 0 ? pivotChanged : this.unionChangedSets(aggregate, pivotChanged);
      pending = this.recalculate(pivotChanged, pivotChanged);
      aggregate = pending.length === 0 ? aggregate : this.unionChangedSets(aggregate, pending);
      forceAll = false;
    }

    return aggregate;
  }

  private refreshPivotOutputs(changed: readonly number[] | U32, forceAll: boolean): U32 {
    const pivots = this.workbook.listPivots();
    if (pivots.length === 0 || (!forceAll && changed.length === 0)) {
      return this.changedUnion.subarray(0, 0);
    }

    const changedCellIndices: number[] = [];
    const changedSeen = new Set<number>();
    for (let index = 0; index < pivots.length; index += 1) {
      const pivot = pivots[index]!;
      if (!forceAll && !this.shouldRefreshPivot(pivot, changed)) {
        continue;
      }
      const pivotChanges = this.materializePivot(pivot);
      for (let changeIndex = 0; changeIndex < pivotChanges.length; changeIndex += 1) {
        const cellIndex = pivotChanges[changeIndex]!;
        if (changedSeen.has(cellIndex)) {
          continue;
        }
        changedSeen.add(cellIndex);
        changedCellIndices.push(cellIndex);
      }
    }

    return changedCellIndices.length === 0
      ? this.changedUnion.subarray(0, 0)
      : Uint32Array.from(changedCellIndices);
  }

  private shouldRefreshPivot(
    pivot: WorkbookPivotRecord,
    changed: readonly number[] | U32,
  ): boolean {
    const bounds = normalizeRange(pivot.source);
    for (let index = 0; index < changed.length; index += 1) {
      const cellIndex = changed[index]!;
      const sheetId = this.workbook.cellStore.sheetIds[cellIndex];
      if (sheetId === undefined) {
        continue;
      }
      if (this.workbook.getSheetNameById(sheetId) !== pivot.source.sheetName) {
        continue;
      }
      const row = this.workbook.cellStore.rows[cellIndex] ?? 0;
      const col = this.workbook.cellStore.cols[cellIndex] ?? 0;
      if (
        row >= bounds.startRow &&
        row <= bounds.endRow &&
        col >= bounds.startCol &&
        col <= bounds.endCol
      ) {
        return true;
      }
    }
    return false;
  }

  private materializePivot(pivot: WorkbookPivotRecord): number[] {
    const changedCellIndices = this.clearOwnedPivot(pivot);
    const sourceSheet = this.workbook.getSheet(pivot.source.sheetName);
    if (!sourceSheet) {
      return this.writePivotOutput(pivot, 1, 1, [errorValue(ErrorCode.Ref)], changedCellIndices);
    }

    if (this.wasm.ready) {
      const bounds = normalizeRange(pivot.source);
      const rangeAddr = parseRangeAddress(
        `${pivot.source.sheetName}!${pivot.source.startAddress}:${pivot.source.endAddress}`,
      );
      const registered = this.ranges.intern(sourceSheet.id, rangeAddr, {
        ensureCell: (sheetId, row, col) => this.ensureCellTrackedByCoords(sheetId, row, col),
        forEachSheetCell: (sheetId, fn) => this.forEachSheetCell(sheetId, fn),
      });
      this.scheduleWasmProgramSync();
      this.flushWasmProgramSync();

      const sourceHeader =
        this.readPivotSourceRows({
          sheetName: pivot.source.sheetName,
          startAddress: pivot.source.startAddress,
          endAddress: formatAddress(bounds.startRow, bounds.endCol),
        })[0] ?? [];
      const headerLookup = new Map<string, number>();
      sourceHeader.forEach((cell, index) => {
        let label = "";
        if (cell.tag === ValueTag.Number) {
          label = String(cell.value);
        } else if (cell.tag === ValueTag.Boolean) {
          label = cell.value ? "TRUE" : "FALSE";
        } else if (cell.tag === ValueTag.String) {
          label = cell.value.trim();
        }
        const normalized = label.trim().toUpperCase();
        if (normalized.length > 0 && !headerLookup.has(normalized)) {
          headerLookup.set(normalized, index);
        }
      });

      const groupByIndexValues = pivot.groupBy.map((g) => headerLookup.get(g.trim().toUpperCase()));
      const valueIndexValues = pivot.values.map((v) =>
        headerLookup.get(v.sourceColumn.trim().toUpperCase()),
      );
      if (
        groupByIndexValues.every((value): value is number => value !== undefined) &&
        valueIndexValues.every((value): value is number => value !== undefined)
      ) {
        const groupByIndices = new Uint32Array(groupByIndexValues);
        const valueIndices = new Uint32Array(valueIndexValues);
        const valueAggs = new Uint8Array(
          pivot.values.map((v) => (v.summarizeBy === "sum" ? 1 : 2)),
        );

        const materialized = this.wasm.materializePivotTable(
          registered.rangeIndex,
          bounds.endCol - bounds.startCol + 1,
          groupByIndices,
          valueIndices,
          valueAggs,
        );

        if (materialized) {
          const owner = parseCellAddress(pivot.address, pivot.sheetName);
          const blockedOutput = this.guardPivotOutputWrite(
            pivot,
            owner.row,
            owner.col,
            materialized.rows,
            materialized.cols,
            changedCellIndices,
          );
          if (blockedOutput) {
            return blockedOutput;
          }
          const values: CellValue[] = [];
          const count = materialized.rows * materialized.cols;
          const groupByCount = pivot.groupBy.length;
          for (let i = 0; i < count; i += 1) {
            if (i < materialized.cols && i >= groupByCount) {
              const valueIndex = i - groupByCount;
              const field = pivot.values[valueIndex]!;
              const label =
                field.outputLabel?.trim() ||
                `${field.summarizeBy.toUpperCase()} of ${field.sourceColumn}`;
              values.push({
                tag: ValueTag.String,
                value: label,
                stringId: this.strings.intern(label),
              });
              continue;
            }
            const tag = materialized.tags[i]! as ValueTag;
            if (tag === ValueTag.Empty) {
              values.push(emptyValue());
            } else if (tag === ValueTag.Number) {
              values.push({ tag: ValueTag.Number, value: materialized.numbers[i]! });
            } else if (tag === ValueTag.Boolean) {
              values.push({ tag: ValueTag.Boolean, value: materialized.numbers[i]! !== 0 });
            } else if (tag === ValueTag.String) {
              const stringId = materialized.stringIds[i]!;
              values.push({ tag: ValueTag.String, value: this.strings.get(stringId), stringId });
            } else if (tag === ValueTag.Error) {
              values.push({ tag: ValueTag.Error, code: materialized.errors[i]! });
            }
          }
          return this.writePivotOutput(
            pivot,
            materialized.rows,
            materialized.cols,
            values,
            changedCellIndices,
          );
        }
      }
    }

    const materialized = materializePivotTable(
      this.toPivotDefinition(pivot),
      this.readPivotSourceRows(pivot.source),
    );
    if (materialized.kind === "error") {
      return this.writePivotOutput(
        pivot,
        materialized.rows,
        materialized.cols,
        materialized.values,
        changedCellIndices,
      );
    }

    const owner = parseCellAddress(pivot.address, pivot.sheetName);
    const blockedOutput = this.guardPivotOutputWrite(
      pivot,
      owner.row,
      owner.col,
      materialized.rows,
      materialized.cols,
      changedCellIndices,
    );
    if (blockedOutput) {
      return blockedOutput;
    }

    return this.writePivotOutput(
      pivot,
      materialized.rows,
      materialized.cols,
      materialized.values,
      changedCellIndices,
    );
  }

  private toPivotDefinition(pivot: WorkbookPivotRecord): PivotDefinitionInput {
    return {
      groupBy: pivot.groupBy,
      values: pivot.values,
    };
  }

  private readPivotSourceRows(range: CellRangeRef): CellValue[][] {
    const bounds = normalizeRange(range);
    const rows: CellValue[][] = [];
    for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
      const values: CellValue[] = [];
      for (let col = bounds.startCol; col <= bounds.endCol; col += 1) {
        const cellIndex = this.workbook.getCellIndex(range.sheetName, formatAddress(row, col));
        values.push(
          cellIndex === undefined
            ? emptyValue()
            : this.workbook.cellStore.getValue(cellIndex, (id) => this.strings.get(id)),
        );
      }
      rows.push(values);
    }
    return rows;
  }

  private resolvePivotData(
    sheetName: string,
    address: string,
    dataField: string,
    filters: ReadonlyArray<{ field: string; item: CellValue }>,
  ): CellValue {
    const target = parseCellAddress(address, sheetName);
    const pivot = this.workbook.listPivots().find((candidate) => {
      if (candidate.sheetName !== sheetName || candidate.rows <= 0 || candidate.cols <= 0) {
        return false;
      }
      const owner = parseCellAddress(candidate.address, candidate.sheetName);
      return (
        target.row >= owner.row &&
        target.row < owner.row + candidate.rows &&
        target.col >= owner.col &&
        target.col < owner.col + candidate.cols
      );
    });
    if (!pivot) {
      return errorValue(ErrorCode.Ref);
    }

    const normalizedDataField = normalizePivotLookupText(dataField);
    const valueField = pivot.values.find((field) => {
      const defaultLabel = `${field.summarizeBy.toUpperCase()} of ${field.sourceColumn}`;
      return (
        normalizePivotLookupText(field.sourceColumn) === normalizedDataField ||
        normalizePivotLookupText(field.outputLabel?.trim() ?? "") === normalizedDataField ||
        normalizePivotLookupText(defaultLabel) === normalizedDataField
      );
    });
    if (!valueField) {
      return errorValue(ErrorCode.Ref);
    }

    const sourceRows = this.readPivotSourceRows(pivot.source);
    const headerRow = sourceRows[0];
    if (!headerRow || headerRow.length === 0) {
      return errorValue(ErrorCode.Ref);
    }

    const headerLookup = new Map<string, number>();
    headerRow.forEach((cell, index) => {
      const normalized = normalizePivotLookupText(cellValueDisplayText(cell));
      if (normalized.length > 0 && !headerLookup.has(normalized)) {
        headerLookup.set(normalized, index);
      }
    });

    const valueColumnIndex = headerLookup.get(normalizePivotLookupText(valueField.sourceColumn));
    if (valueColumnIndex === undefined) {
      return errorValue(ErrorCode.Ref);
    }

    const materializedFilters = filters.map((filter) => ({
      fieldIndex: headerLookup.get(normalizePivotLookupText(filter.field)),
      item: filter.item,
    }));
    if (materializedFilters.some((filter) => filter.fieldIndex === undefined)) {
      return errorValue(ErrorCode.Ref);
    }

    for (let filterIndex = 0; filterIndex < materializedFilters.length; filterIndex += 1) {
      const filter = materializedFilters[filterIndex]!;
      const fieldIndex = filter.fieldIndex!;
      const itemSeen = sourceRows
        .slice(1)
        .some((row) => pivotItemMatches(row[fieldIndex] ?? emptyValue(), filter.item));
      if (!itemSeen) {
        return errorValue(ErrorCode.Ref);
      }
    }

    let matched = filters.length === 0;
    let aggregate = 0;
    for (let rowIndex = 1; rowIndex < sourceRows.length; rowIndex += 1) {
      const row = sourceRows[rowIndex] ?? [];
      const matches = materializedFilters.every((filter) =>
        pivotItemMatches(row[filter.fieldIndex!] ?? emptyValue(), filter.item),
      );
      if (!matches) {
        continue;
      }
      matched = true;
      const value = row[valueColumnIndex] ?? emptyValue();
      if (valueField.summarizeBy === "count") {
        aggregate += value.tag === ValueTag.Empty ? 0 : 1;
      } else if (value.tag === ValueTag.Number) {
        aggregate += value.value;
      }
    }

    return matched ? { tag: ValueTag.Number, value: aggregate } : errorValue(ErrorCode.Ref);
  }

  private resolveMultipleOperations(request: {
    formulaSheetName: string;
    formulaAddress: string;
    rowCellSheetName: string;
    rowCellAddress: string;
    rowReplacementSheetName: string;
    rowReplacementAddress: string;
    columnCellSheetName?: string;
    columnCellAddress?: string;
    columnReplacementSheetName?: string;
    columnReplacementAddress?: string;
  }): CellValue {
    const replacements = new Map<string, { sheetName: string; address: string }>();
    replacements.set(
      this.referenceReplacementKey(request.rowCellSheetName, request.rowCellAddress),
      {
        sheetName: request.rowReplacementSheetName,
        address: request.rowReplacementAddress,
      },
    );
    if (
      request.columnCellSheetName &&
      request.columnCellAddress &&
      request.columnReplacementSheetName &&
      request.columnReplacementAddress
    ) {
      replacements.set(
        this.referenceReplacementKey(request.columnCellSheetName, request.columnCellAddress),
        {
          sheetName: request.columnReplacementSheetName,
          address: request.columnReplacementAddress,
        },
      );
    }
    return this.evaluateCellWithReferenceReplacements(
      request.formulaSheetName,
      request.formulaAddress,
      replacements,
      new Set<string>(),
    );
  }

  private referenceReplacementKey(sheetName: string, address: string): string {
    return `${sheetName.trim().toUpperCase()}!${address.trim().toUpperCase()}`;
  }

  private evaluateCellWithReferenceReplacements(
    sheetName: string,
    address: string,
    replacements: ReadonlyMap<string, { sheetName: string; address: string }>,
    visiting: Set<string>,
  ): CellValue {
    const replacementKey = this.referenceReplacementKey(sheetName, address);
    const replacement = replacements.get(replacementKey);
    if (replacement) {
      return this.evaluateCellWithReferenceReplacements(
        replacement.sheetName,
        replacement.address,
        replacements,
        visiting,
      );
    }

    const visitKey = this.referenceReplacementKey(sheetName, address);
    if (visiting.has(visitKey)) {
      return errorValue(ErrorCode.Cycle);
    }

    const cellIndex = this.workbook.getCellIndex(sheetName, address);
    if (cellIndex === undefined) {
      return emptyValue();
    }

    const formula = this.formulas.get(cellIndex);
    if (!formula) {
      return this.workbook.cellStore.getValue(cellIndex, (id) => this.strings.get(id));
    }

    visiting.add(visitKey);
    const evaluationContext = {
      sheetName,
      currentAddress: address,
      resolveCell: (targetSheetName, targetAddress) =>
        this.evaluateCellWithReferenceReplacements(
          targetSheetName,
          targetAddress,
          replacements,
          visiting,
        ),
      resolveRange: (targetSheetName, start, end, refKind) => {
        if (refKind !== "cells") {
          return [];
        }
        const range = parseRangeAddress(`${start}:${end}`, targetSheetName);
        if (range.kind !== "cells") {
          return [];
        }
        const values: CellValue[] = [];
        for (let row = range.start.row; row <= range.end.row; row += 1) {
          for (let col = range.start.col; col <= range.end.col; col += 1) {
            values.push(
              this.evaluateCellWithReferenceReplacements(
                targetSheetName,
                formatAddress(row, col),
                replacements,
                visiting,
              ),
            );
          }
        }
        return values;
      },
      resolveName: (name) => {
        const definedName = this.workbook.getDefinedName(name);
        if (!definedName) {
          return errorValue(ErrorCode.Name);
        }
        return definedNameValueToCellValue(definedName.value, this.strings);
      },
      resolveFormula: (targetSheetName: string, targetAddress: string) =>
        this.getCell(targetSheetName, targetAddress).formula,
      resolvePivotData: ({
        dataField,
        sheetName: pivotSheetName,
        address: pivotAddress,
        filters,
      }) => this.resolvePivotData(pivotSheetName, pivotAddress, dataField, filters),
      resolveMultipleOperations: (nested: {
        formulaSheetName: string;
        formulaAddress: string;
        rowCellSheetName: string;
        rowCellAddress: string;
        rowReplacementSheetName: string;
        rowReplacementAddress: string;
        columnCellSheetName?: string;
        columnCellAddress?: string;
        columnReplacementSheetName?: string;
        columnReplacementAddress?: string;
      }) => this.resolveMultipleOperations(nested),
      listSheetNames: () =>
        [...this.workbook.sheetsByName.values()]
          .toSorted((left, right) => left.order - right.order)
          .map((sheet) => sheet.name),
    } as Parameters<typeof evaluatePlanResult>[1];
    const result = evaluatePlanResult(formula.compiled.jsPlan, evaluationContext);
    visiting.delete(visitKey);
    return isArrayValue(result) ? (result.values[0] ?? emptyValue()) : result;
  }

  private isPivotOutputBlocked(
    pivot: WorkbookPivotRecord,
    startRow: number,
    startCol: number,
    rows: number,
    cols: number,
  ): boolean {
    const ownerKey = pivotKey(pivot.sheetName, pivot.address);
    for (let rowOffset = 0; rowOffset < rows; rowOffset += 1) {
      for (let colOffset = 0; colOffset < cols; colOffset += 1) {
        const targetIndex = this.workbook.getCellIndex(
          pivot.sheetName,
          formatAddress(startRow + rowOffset, startCol + colOffset),
        );
        if (targetIndex === undefined) {
          continue;
        }
        const pivotOwner = this.pivotOutputOwners.get(targetIndex);
        if (pivotOwner && pivotOwner !== ownerKey) {
          return true;
        }
        if (this.formulas.get(targetIndex)) {
          return true;
        }
        const targetFlags = this.workbook.cellStore.flags[targetIndex] ?? 0;
        if ((targetFlags & CellFlags.SpillChild) !== 0) {
          return true;
        }
        const targetValue = this.workbook.cellStore.getValue(targetIndex, (id) =>
          this.strings.get(id),
        );
        if (!pivotOwner && targetValue.tag !== ValueTag.Empty) {
          return true;
        }
      }
    }
    return false;
  }

  private guardPivotOutputWrite(
    pivot: WorkbookPivotRecord,
    startRow: number,
    startCol: number,
    rows: number,
    cols: number,
    changedCellIndices: number[],
  ): number[] | undefined {
    if (startRow + rows > MAX_ROWS || startCol + cols > MAX_COLS) {
      return this.writePivotOutput(pivot, 1, 1, [errorValue(ErrorCode.Spill)], changedCellIndices);
    }
    if (this.isPivotOutputBlocked(pivot, startRow, startCol, rows, cols)) {
      return this.writePivotOutput(
        pivot,
        1,
        1,
        [errorValue(ErrorCode.Blocked)],
        changedCellIndices,
      );
    }
    return undefined;
  }

  private writePivotOutput(
    pivot: WorkbookPivotRecord,
    rows: number,
    cols: number,
    values: readonly CellValue[],
    changedCellIndices: number[],
  ): number[] {
    const sheet = this.workbook.getOrCreateSheet(pivot.sheetName);
    const owner = parseCellAddress(pivot.address, pivot.sheetName);
    const ownerKey = pivotKey(pivot.sheetName, pivot.address);
    const changedSeen = new Set(changedCellIndices);

    for (let rowOffset = 0; rowOffset < rows; rowOffset += 1) {
      for (let colOffset = 0; colOffset < cols; colOffset += 1) {
        const valueIndex = rowOffset * cols + colOffset;
        const cellValue = values[valueIndex] ?? emptyValue();
        const cellIndex = this.ensureCellTrackedByCoords(
          sheet.id,
          owner.row + rowOffset,
          owner.col + colOffset,
        );
        if (this.setPivotOutputCellValue(cellIndex, cellValue, ownerKey)) {
          if (!changedSeen.has(cellIndex)) {
            changedSeen.add(cellIndex);
            changedCellIndices.push(cellIndex);
          }
        }
      }
    }

    if (pivot.rows !== rows || pivot.cols !== cols) {
      this.applyDerivedOp({
        kind: "upsertPivotTable",
        name: pivot.name,
        sheetName: pivot.sheetName,
        address: pivot.address,
        source: { ...pivot.source },
        groupBy: [...pivot.groupBy],
        values: pivot.values.map((value) => Object.assign({}, value)),
        rows,
        cols,
      });
    }
    return changedCellIndices;
  }

  private clearOwnedPivot(pivot: WorkbookPivotRecord): number[] {
    const changedCellIndices: number[] = [];
    const ownerKey = pivotKey(pivot.sheetName, pivot.address);
    const owner = parseCellAddress(pivot.address, pivot.sheetName);
    for (let rowOffset = 0; rowOffset < pivot.rows; rowOffset += 1) {
      for (let colOffset = 0; colOffset < pivot.cols; colOffset += 1) {
        const cellIndex = this.workbook.getCellIndex(
          pivot.sheetName,
          formatAddress(owner.row + rowOffset, owner.col + colOffset),
        );
        if (cellIndex === undefined || this.pivotOutputOwners.get(cellIndex) !== ownerKey) {
          continue;
        }
        if (this.clearPivotOutputCell(cellIndex)) {
          changedCellIndices.push(cellIndex);
        }
      }
    }
    return changedCellIndices;
  }

  private clearPivotForCell(cellIndex: number): number[] {
    const ownerKey = this.pivotOutputOwners.get(cellIndex);
    if (!ownerKey) {
      return [];
    }
    const pivot = this.workbook.getPivotByKey(ownerKey);
    if (!pivot) {
      this.pivotOutputOwners.delete(cellIndex);
      return [];
    }
    return this.applyDerivedOp({
      kind: "deletePivotTable",
      sheetName: pivot.sheetName,
      address: pivot.address,
    });
  }

  private clearPivotOutputCell(cellIndex: number): boolean {
    const currentFlags = this.workbook.cellStore.flags[cellIndex] ?? 0;
    const currentValue = this.workbook.cellStore.getValue(cellIndex, (id) => this.strings.get(id));
    if (currentValue.tag === ValueTag.Empty && (currentFlags & CellFlags.PivotOutput) === 0) {
      this.pivotOutputOwners.delete(cellIndex);
      return false;
    }
    this.pivotOutputOwners.delete(cellIndex);
    this.workbook.cellStore.setValue(cellIndex, emptyValue());
    this.workbook.cellStore.flags[cellIndex] =
      currentFlags &
      ~(
        CellFlags.HasFormula |
        CellFlags.JsOnly |
        CellFlags.InCycle |
        CellFlags.SpillChild |
        CellFlags.PivotOutput
      );
    return true;
  }

  private setPivotOutputCellValue(cellIndex: number, value: CellValue, ownerKey: string): boolean {
    const currentFlags = this.workbook.cellStore.flags[cellIndex] ?? 0;
    const currentValue = this.workbook.cellStore.getValue(cellIndex, (id) => this.strings.get(id));
    const nextFlags =
      (currentFlags &
        ~(CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle | CellFlags.SpillChild)) |
      CellFlags.PivotOutput;
    if (
      areCellValuesEqual(currentValue, value) &&
      currentFlags === nextFlags &&
      this.pivotOutputOwners.get(cellIndex) === ownerKey
    ) {
      return false;
    }
    this.workbook.cellStore.setValue(
      cellIndex,
      value,
      value.tag === ValueTag.String ? this.strings.intern(value.value) : 0,
    );
    this.workbook.cellStore.flags[cellIndex] = nextFlags;
    this.pivotOutputOwners.set(cellIndex, ownerKey);
    return true;
  }

  exportReplicaSnapshot(): EngineReplicaSnapshot {
    return {
      replica: exportReplicaStateSnapshot(this.replicaState),
      entityVersions: [...this.entityVersions.entries()].map(([entityKey, order]) => ({
        entityKey,
        order,
      })),
      sheetDeleteVersions: [...this.sheetDeleteVersions.entries()].map(([sheetName, order]) => ({
        sheetName,
        order,
      })),
    };
  }

  importReplicaSnapshot(snapshot: EngineReplicaSnapshot): void {
    hydrateReplicaState(this.replicaState, snapshot.replica);
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
        case "renameSheet":
          if (op.oldName && op.newName) {
            engineOps.push({ kind: "renameSheet", oldName: op.oldName, newName: op.newName });
          }
          break;
        case "deleteSheet":
          if (op.name) engineOps.push({ kind: "deleteSheet", name: op.name });
          break;
        case "upsertCell":
          if (!op.sheetName || !op.addr) break;
          if (op.formula !== undefined) {
            engineOps.push({
              kind: "setCellFormula",
              sheetName: op.sheetName,
              address: op.addr,
              formula: op.formula,
            });
          } else {
            engineOps.push({
              kind: "setCellValue",
              sheetName: op.sheetName,
              address: op.addr,
              value: op.value ?? null,
            });
          }
          potentialNewCells += 1;
          if (op.format !== undefined) {
            engineOps.push({
              kind: "setCellFormat",
              sheetName: op.sheetName,
              address: op.addr,
              format: op.format,
            });
          }
          break;
        case "deleteCell":
          if (op.sheetName && op.addr) {
            engineOps.push({ kind: "clearCell", sheetName: op.sheetName, address: op.addr });
            engineOps.push({
              kind: "setCellFormat",
              sheetName: op.sheetName,
              address: op.addr,
              format: null,
            });
          }
          break;
      }
    });
    this.executeLocalTransaction(engineOps, potentialNewCells);
  }

  applyRemoteBatch(batch: EngineOpBatch): boolean {
    if (!shouldApplyBatch(this.replicaState, batch)) {
      return false;
    }
    this.applyBatch(batch, "remote");
    return true;
  }

  captureUndoOps<T>(mutate: () => T): {
    result: T;
    undoOps: readonly EngineOp[] | null;
  } {
    const previousUndoDepth = this.undoStack.length;
    const result = mutate();
    if (this.undoStack.length === previousUndoDepth) {
      return {
        result,
        undoOps: null,
      };
    }
    if (this.undoStack.length === previousUndoDepth + 1) {
      return {
        result,
        undoOps: structuredClone(this.undoStack.at(-1)?.inverse.ops ?? null),
      };
    }
    throw new Error("Expected a single local transaction while capturing undo ops");
  }

  applyOps(
    ops: readonly EngineOp[],
    options: {
      captureUndo?: boolean;
      potentialNewCells?: number;
    } = {},
  ): readonly EngineOp[] | null {
    const nextOps = structuredClone([...ops]);
    if (nextOps.length === 0) {
      return null;
    }
    if (options.captureUndo) {
      return this.executeLocalTransaction(nextOps, options.potentialNewCells);
    }
    this.executeTransaction(
      options.potentialNewCells === undefined
        ? { ops: nextOps }
        : { ops: nextOps, potentialNewCells: options.potentialNewCells },
      "restore",
    );
    return null;
  }

  private executeLocalTransaction(
    ops: EngineOp[],
    potentialNewCells?: number,
  ): readonly EngineOp[] | null {
    if (ops.length === 0) {
      return null;
    }
    const forward: TransactionRecord =
      potentialNewCells === undefined ? { ops } : { ops, potentialNewCells };
    const inverse: TransactionRecord = {
      ops: this.buildInverseOps(ops),
      potentialNewCells: ops.length,
    };
    this.executeTransaction(forward, "local");
    if (this.transactionReplayDepth === 0) {
      this.undoStack.push({ forward, inverse });
      this.redoStack.length = 0;
    }
    return structuredClone(inverse.ops);
  }

  private executeTransaction(
    record: TransactionRecord,
    source: "local" | "restore" | "history",
  ): void {
    if (record.ops.length === 0) {
      return;
    }
    const batch = createBatch(this.replicaState, record.ops);
    this.applyBatch(batch, source, record.potentialNewCells);
  }

  private applyDerivedOp(
    op: Extract<
      EngineOp,
      { kind: "upsertSpillRange" | "deleteSpillRange" | "upsertPivotTable" | "deletePivotTable" }
    >,
  ): number[] {
    const batch = createBatch(this.replicaState, [op]);
    const order = batchOpOrder(batch, 0);
    switch (op.kind) {
      case "upsertSpillRange":
      case "deleteSpillRange": {
        const candidates = this.applySpillRangeOp(op, order);
        this.rebindFormulaCells(candidates, 0);
        return candidates;
      }
      case "upsertPivotTable":
        this.applyPivotUpsertOp(op, order);
        return [];
      case "deletePivotTable":
        return this.applyPivotDeleteOp(op, order);
      default:
        return assertNever(op);
    }
  }

  private applySpillRangeOp(
    op: Extract<EngineOp, { kind: "upsertSpillRange" | "deleteSpillRange" }>,
    order: OpOrder,
  ): number[] {
    if (op.kind === "upsertSpillRange") {
      this.workbook.setSpill(op.sheetName, op.address, op.rows, op.cols);
    } else {
      this.workbook.deleteSpill(op.sheetName, op.address);
    }
    this.entityVersions.set(this.entityKeyForOp(op), order);
    const spillKey = spillDependencyKey(op.sheetName, op.address);
    return this.collectTrackedDependents(this.reverseSpillEdges, [spillKey]);
  }

  private applyPivotUpsertOp(
    op: Extract<EngineOp, { kind: "upsertPivotTable" }>,
    order: OpOrder,
  ): void {
    this.workbook.setPivot({
      name: op.name,
      sheetName: op.sheetName,
      address: op.address,
      source: op.source,
      groupBy: op.groupBy,
      values: op.values,
      rows: op.rows,
      cols: op.cols,
    });
    this.entityVersions.set(this.entityKeyForOp(op), order);
  }

  private applyPivotDeleteOp(
    op: Extract<EngineOp, { kind: "deletePivotTable" }>,
    order: OpOrder,
  ): number[] {
    const pivot = this.workbook.getPivot(op.sheetName, op.address);
    if (!pivot) {
      this.entityVersions.set(this.entityKeyForOp(op), order);
      return [];
    }
    const changedPivotOutputs = this.clearOwnedPivot(pivot);
    this.workbook.deletePivot(op.sheetName, op.address);
    this.entityVersions.set(this.entityKeyForOp(op), order);
    return changedPivotOutputs;
  }

  private applyBatch(
    batch: EngineOpBatch,
    source: "local" | "remote" | "restore" | "history",
    potentialNewCells?: number,
  ): void {
    this.beginMutationCollection();
    let changedInputCount = 0;
    let formulaChangedCount = 0;
    let explicitChangedCount = 0;
    let topologyChanged = false;
    let sheetDeleted = false;
    let structuralInvalidation = false;
    const invalidatedRanges: CellRangeRef[] = [];
    const invalidatedRows: { sheetName: string; startIndex: number; endIndex: number }[] = [];
    const invalidatedColumns: { sheetName: string; startIndex: number; endIndex: number }[] = [];
    let refreshAllPivots = false;
    let appliedOps = 0;
    const canSkipOrderChecks = source !== "remote";

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
          case "setWorkbookMetadata":
            this.workbook.setWorkbookProperty(op.key, op.value);
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          case "setCalculationSettings":
            this.workbook.setCalculationSettings(op.settings);
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          case "setVolatileContext":
            this.workbook.setVolatileContext(op.context);
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          case "upsertSheet":
            this.workbook.createSheet(op.name, op.order, op.id);
            this.entityVersions.set(this.entityKeyForOp(op), order);
            const tombstone = this.sheetDeleteVersions.get(op.name);
            if (!tombstone || compareOpOrder(order, tombstone) > 0) {
              this.sheetDeleteVersions.delete(op.name);
            }
            const sheetReboundCount = formulaChangedCount;
            formulaChangedCount = this.rebindFormulasForSheet(op.name, formulaChangedCount);
            topologyChanged = topologyChanged || formulaChangedCount !== sheetReboundCount;
            refreshAllPivots = true;
            break;
          case "renameSheet": {
            const renamedSheet = this.workbook.renameSheet(op.oldName, op.newName);
            this.entityVersions.set(`sheet:${op.oldName}`, order);
            this.entityVersions.set(`sheet:${op.newName}`, order);
            this.sheetDeleteVersions.set(op.oldName, order);
            const renamedTombstone = this.sheetDeleteVersions.get(op.newName);
            if (!renamedTombstone || compareOpOrder(order, renamedTombstone) > 0) {
              this.sheetDeleteVersions.delete(op.newName);
            }
            if (!renamedSheet) {
              break;
            }
            if (this.selection.sheetName === op.oldName) {
              this.setSelection(op.newName, this.selection.address);
            }
            this.rewriteDefinedNamesForSheetRename(op.oldName, op.newName);
            formulaChangedCount = this.rewriteCellFormulasForSheetRename(
              op.oldName,
              op.newName,
              formulaChangedCount,
            );
            topologyChanged = true;
            sheetDeleted = true;
            structuralInvalidation = true;
            refreshAllPivots = true;
            break;
          }
          case "deleteSheet":
            const removal = this.removeSheetRuntime(op.name, explicitChangedCount);
            changedInputCount += removal.changedInputCount;
            formulaChangedCount += removal.formulaChangedCount;
            explicitChangedCount = removal.explicitChangedCount;
            this.entityVersions.set(this.entityKeyForOp(op), order);
            this.sheetDeleteVersions.set(op.name, order);
            topologyChanged = true;
            sheetDeleted = true;
            structuralInvalidation = true;
            refreshAllPivots = true;
            break;
          case "insertRows":
          case "deleteRows":
          case "moveRows":
          case "insertColumns":
          case "deleteColumns":
          case "moveColumns": {
            const structural = this.applyStructuralAxisOp(op);
            structural.changedCellIndices.forEach((cellIndex) => {
              explicitChangedCount = this.markExplicitChanged(cellIndex, explicitChangedCount);
            });
            structural.formulaCellIndices.forEach((cellIndex) => {
              formulaChangedCount = this.markFormulaChanged(cellIndex, formulaChangedCount);
            });
            topologyChanged = true;
            structuralInvalidation = true;
            refreshAllPivots = true;
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          }
          case "updateRowMetadata":
            this.workbook.setRowMetadata(op.sheetName, op.start, op.count, op.size, op.hidden);
            invalidatedRows.push({
              sheetName: op.sheetName,
              startIndex: op.start,
              endIndex: op.start + op.count - 1,
            });
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          case "updateColumnMetadata":
            this.workbook.setColumnMetadata(op.sheetName, op.start, op.count, op.size, op.hidden);
            invalidatedColumns.push({
              sheetName: op.sheetName,
              startIndex: op.start,
              endIndex: op.start + op.count - 1,
            });
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          case "setFreezePane":
            this.workbook.setFreezePane(op.sheetName, op.rows, op.cols);
            structuralInvalidation = true;
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          case "clearFreezePane":
            this.workbook.clearFreezePane(op.sheetName);
            structuralInvalidation = true;
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          case "setFilter":
            this.workbook.setFilter(op.sheetName, op.range);
            structuralInvalidation = true;
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          case "clearFilter":
            this.workbook.deleteFilter(op.sheetName, op.range);
            structuralInvalidation = true;
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          case "setSort":
            this.workbook.setSort(op.sheetName, op.range, op.keys);
            structuralInvalidation = true;
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          case "clearSort":
            this.workbook.deleteSort(op.sheetName, op.range);
            structuralInvalidation = true;
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          case "upsertTable":
            this.workbook.setTable(op.table);
            {
              const tableReboundCount = formulaChangedCount;
              formulaChangedCount = this.rebindTableDependents(
                [tableDependencyKey(op.table.name)],
                formulaChangedCount,
              );
              topologyChanged = topologyChanged || formulaChangedCount !== tableReboundCount;
            }
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          case "deleteTable":
            this.workbook.deleteTable(op.name);
            {
              const tableReboundCount = formulaChangedCount;
              formulaChangedCount = this.rebindTableDependents(
                [tableDependencyKey(op.name)],
                formulaChangedCount,
              );
              topologyChanged = topologyChanged || formulaChangedCount !== tableReboundCount;
            }
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          case "upsertSpillRange":
            {
              const spillReboundCount = formulaChangedCount;
              formulaChangedCount = this.rebindFormulaCells(
                this.applySpillRangeOp(op, order),
                formulaChangedCount,
              );
              topologyChanged = topologyChanged || formulaChangedCount !== spillReboundCount;
            }
            break;
          case "deleteSpillRange":
            {
              const spillReboundCount = formulaChangedCount;
              formulaChangedCount = this.rebindFormulaCells(
                this.applySpillRangeOp(op, order),
                formulaChangedCount,
              );
              topologyChanged = topologyChanged || formulaChangedCount !== spillReboundCount;
            }
            break;
          case "setCellValue": {
            const existingIndex = this.workbook.getCellIndex(op.sheetName, op.address);
            if (existingIndex !== undefined) {
              changedInputCount = this.markPivotRootsChanged(
                this.clearPivotForCell(existingIndex),
                changedInputCount,
              );
            }
            const cellIndex = this.ensureCellTracked(op.sheetName, op.address);
            changedInputCount = this.markSpillRootsChanged(
              this.clearOwnedSpill(cellIndex),
              changedInputCount,
            );
            topologyChanged = this.removeFormula(cellIndex) || topologyChanged;
            const value = literalToValue(op.value, this.strings);
            this.workbook.cellStore.setValue(
              cellIndex,
              value,
              value.tag === ValueTag.String ? value.stringId : 0,
            );
            this.workbook.cellStore.flags[cellIndex] =
              (this.workbook.cellStore.flags[cellIndex] ?? 0) &
              ~(
                CellFlags.HasFormula |
                CellFlags.JsOnly |
                CellFlags.InCycle |
                CellFlags.SpillChild |
                CellFlags.PivotOutput
              );
            changedInputCount = this.markInputChanged(cellIndex, changedInputCount);
            explicitChangedCount = this.markExplicitChanged(cellIndex, explicitChangedCount);
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          }
          case "setCellFormula": {
            const existingIndex = this.workbook.getCellIndex(op.sheetName, op.address);
            if (existingIndex !== undefined) {
              changedInputCount = this.markPivotRootsChanged(
                this.clearPivotForCell(existingIndex),
                changedInputCount,
              );
            }
            const cellIndex = this.ensureCellTracked(op.sheetName, op.address);
            changedInputCount = this.markSpillRootsChanged(
              this.clearOwnedSpill(cellIndex),
              changedInputCount,
            );
            const compileStarted = performance.now();
            try {
              const compiled = this.compileFormulaForSheet(op.sheetName, op.formula);
              this.lastMetrics.compileMs = performance.now() - compileStarted;
              const dependencies = this.materializeDependencies(op.sheetName, compiled);
              this.setFormula(cellIndex, op.formula, compiled, dependencies);
              formulaChangedCount = this.markFormulaChanged(cellIndex, formulaChangedCount);
              topologyChanged = true;
            } catch {
              this.lastMetrics.compileMs = performance.now() - compileStarted;
              topologyChanged = this.removeFormula(cellIndex) || topologyChanged;
              this.setInvalidFormulaValue(cellIndex);
              changedInputCount = this.markInputChanged(cellIndex, changedInputCount);
            }
            explicitChangedCount = this.markExplicitChanged(cellIndex, explicitChangedCount);
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          }
          case "setCellFormat": {
            const cellIndex = this.ensureCellTracked(op.sheetName, op.address);
            this.workbook.setCellFormat(cellIndex, op.format);
            explicitChangedCount = this.markExplicitChanged(cellIndex, explicitChangedCount);
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          }
          case "upsertCellStyle":
            this.workbook.upsertCellStyle(op.style);
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          case "upsertCellNumberFormat":
            this.workbook.upsertCellNumberFormat(op.format);
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          case "setStyleRange":
            this.workbook.setStyleRange(op.range, op.styleId);
            invalidatedRanges.push(op.range);
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          case "setFormatRange":
            this.workbook.setFormatRange(op.range, op.formatId);
            invalidatedRanges.push(op.range);
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          case "clearCell": {
            const cellIndex = this.workbook.getCellIndex(op.sheetName, op.address);
            if (cellIndex === undefined) {
              this.entityVersions.set(this.entityKeyForOp(op), order);
              break;
            }
            changedInputCount = this.markPivotRootsChanged(
              this.clearPivotForCell(cellIndex),
              changedInputCount,
            );
            changedInputCount = this.markSpillRootsChanged(
              this.clearOwnedSpill(cellIndex),
              changedInputCount,
            );
            topologyChanged = this.removeFormula(cellIndex) || topologyChanged;
            this.workbook.cellStore.setValue(cellIndex, emptyValue());
            this.workbook.cellStore.flags[cellIndex] =
              (this.workbook.cellStore.flags[cellIndex] ?? 0) &
              ~(
                CellFlags.HasFormula |
                CellFlags.JsOnly |
                CellFlags.InCycle |
                CellFlags.SpillChild |
                CellFlags.PivotOutput
              );
            changedInputCount = this.markInputChanged(cellIndex, changedInputCount);
            explicitChangedCount = this.markExplicitChanged(cellIndex, explicitChangedCount);
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          }
          case "upsertDefinedName": {
            const normalizedName = normalizeDefinedName(op.name);
            this.workbook.setDefinedName(op.name, op.value);
            const nameReboundCount = formulaChangedCount;
            formulaChangedCount = this.rebindDefinedNameDependents(
              [normalizedName],
              formulaChangedCount,
            );
            topologyChanged = topologyChanged || formulaChangedCount !== nameReboundCount;
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          }
          case "deleteDefinedName": {
            const normalizedName = normalizeDefinedName(op.name);
            this.workbook.deleteDefinedName(op.name);
            const nameReboundCount = formulaChangedCount;
            formulaChangedCount = this.rebindDefinedNameDependents(
              [normalizedName],
              formulaChangedCount,
            );
            topologyChanged = topologyChanged || formulaChangedCount !== nameReboundCount;
            this.entityVersions.set(this.entityKeyForOp(op), order);
            break;
          }
          case "upsertPivotTable": {
            this.applyPivotUpsertOp(op, order);
            refreshAllPivots = true;
            break;
          }
          case "deletePivotTable": {
            const changedPivotOutputs = this.applyPivotDeleteOp(op, order);
            changedInputCount = this.markPivotRootsChanged(changedPivotOutputs, changedInputCount);
            changedPivotOutputs.forEach((cellIndex) => {
              explicitChangedCount = this.markExplicitChanged(cellIndex, explicitChangedCount);
            });
            refreshAllPivots = true;
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

    markBatchApplied(this.replicaState, batch);
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
    formulaChangedCount = this.markVolatileFormulasChanged(formulaChangedCount);
    const changedInputArray = this.changedInputBuffer.subarray(0, changedInputCount);
    let recalculated = this.recalculate(
      this.composeMutationRoots(changedInputCount, formulaChangedCount),
      changedInputArray,
    );
    recalculated = this.reconcilePivotOutputs(recalculated, refreshAllPivots);
    const changed = this.composeEventChanges(recalculated, explicitChangedCount);
    this.lastMetrics.batchId += 1;
    this.lastMetrics.changedInputCount = changedInputCount + formulaChangedCount;
    const event = {
      kind: "batch",
      invalidation: sheetDeleted || structuralInvalidation ? "full" : "cells",
      changedCellIndices: changed,
      invalidatedRanges,
      invalidatedRows,
      invalidatedColumns,
      metrics: this.lastMetrics,
    } satisfies EngineEvent;
    if (event.invalidation === "full") {
      this.events.emitAllWatched(event);
    } else {
      this.events.emit(event, changed, (cellIndex) => this.workbook.getQualifiedAddress(cellIndex));
    }
    if (source === "local") {
      void this.syncClientConnection?.send(batch);
      this.emitBatch(batch);
    } else if (source === "remote" && this.redoStack.length > 0) {
      this.redoStack.length = 0;
    }
  }

  private buildInverseOps(ops: readonly EngineOp[]): EngineOp[] {
    const inverseOps: EngineOp[] = [];

    for (let index = ops.length - 1; index >= 0; index -= 1) {
      inverseOps.push(...this.inverseOpsFor(ops[index]!));
    }

    return inverseOps;
  }

  private inverseOpsFor(op: EngineOp): EngineOp[] {
    switch (op.kind) {
      case "upsertWorkbook":
        return [{ kind: "upsertWorkbook", name: this.workbook.workbookName }];
      case "setWorkbookMetadata": {
        const existing = this.workbook.getWorkbookProperty(op.key);
        return [{ kind: "setWorkbookMetadata", key: op.key, value: existing?.value ?? null }];
      }
      case "setCalculationSettings":
        return [
          { kind: "setCalculationSettings", settings: this.workbook.getCalculationSettings() },
        ];
      case "setVolatileContext":
        return [{ kind: "setVolatileContext", context: this.workbook.getVolatileContext() }];
      case "upsertSheet": {
        const existing = this.workbook.getSheet(op.name);
        if (!existing) {
          return [{ kind: "deleteSheet", name: op.name }];
        }
        return [{ kind: "upsertSheet", name: existing.name, order: existing.order }];
      }
      case "renameSheet": {
        const existing = this.workbook.getSheet(op.newName);
        if (!existing) {
          return [];
        }
        return [{ kind: "renameSheet", oldName: op.newName, newName: op.oldName }];
      }
      case "deleteSheet": {
        const sheet = this.workbook.getSheet(op.name);
        if (!sheet) {
          return [];
        }
        const restoredOps: EngineOp[] = [
          { kind: "upsertSheet", name: sheet.name, order: sheet.order },
        ];
        restoredOps.push(...this.sheetMetadataToOps(sheet.name));
        this.workbook
          .listTables()
          .filter((table) => table.sheetName === sheet.name)
          .forEach((table) => {
            restoredOps.push({
              kind: "upsertTable",
              table: {
                name: table.name,
                sheetName: table.sheetName,
                startAddress: table.startAddress,
                endAddress: table.endAddress,
                columnNames: [...table.columnNames],
                headerRow: table.headerRow,
                totalsRow: table.totalsRow,
              },
            });
          });
        this.workbook
          .listSpills()
          .filter((spill) => spill.sheetName === sheet.name)
          .forEach((spill) => {
            restoredOps.push({
              kind: "upsertSpillRange",
              sheetName: spill.sheetName,
              address: spill.address,
              rows: spill.rows,
              cols: spill.cols,
            });
          });
        this.workbook
          .listPivots()
          .filter((pivot) => pivot.sheetName === sheet.name)
          .forEach((pivot) => {
            restoredOps.push({
              kind: "upsertPivotTable",
              name: pivot.name,
              sheetName: pivot.sheetName,
              address: pivot.address,
              source: { ...pivot.source },
              groupBy: [...pivot.groupBy],
              values: pivot.values.map((value) => Object.assign({}, value)),
              rows: pivot.rows,
              cols: pivot.cols,
            });
          });
        const cellIndices: number[] = [];
        sheet.grid.forEachCell((cellIndex) => {
          cellIndices.push(cellIndex);
        });
        const orderedCellIndices = cellIndices.toSorted((left, right) => {
          const leftRow = this.workbook.cellStore.rows[left] ?? 0;
          const rightRow = this.workbook.cellStore.rows[right] ?? 0;
          const leftCol = this.workbook.cellStore.cols[left] ?? 0;
          const rightCol = this.workbook.cellStore.cols[right] ?? 0;
          return leftRow - rightRow || leftCol - rightCol;
        });
        for (const cellIndex of orderedCellIndices) {
          restoredOps.push(
            ...this.toCellStateOps(
              sheet.name,
              this.workbook.getAddress(cellIndex),
              this.getCellByIndex(cellIndex),
            ),
          );
        }
        return restoredOps;
      }
      case "insertRows":
        return [{ kind: "deleteRows", sheetName: op.sheetName, start: op.start, count: op.count }];
      case "deleteRows": {
        const entries = this.workbook.materializeRowAxisEntries(op.sheetName, op.start, op.count);
        return [
          {
            kind: "insertRows",
            sheetName: op.sheetName,
            start: op.start,
            count: op.count,
            entries,
          },
          ...this.captureRowRangeCellState(op.sheetName, op.start, op.count),
        ];
      }
      case "moveRows":
        return [
          {
            kind: "moveRows",
            sheetName: op.sheetName,
            start: op.target,
            count: op.count,
            target: op.start,
          },
        ];
      case "insertColumns":
        return [
          { kind: "deleteColumns", sheetName: op.sheetName, start: op.start, count: op.count },
        ];
      case "deleteColumns": {
        const entries = this.workbook.materializeColumnAxisEntries(
          op.sheetName,
          op.start,
          op.count,
        );
        return [
          {
            kind: "insertColumns",
            sheetName: op.sheetName,
            start: op.start,
            count: op.count,
            entries,
          },
          ...this.captureColumnRangeCellState(op.sheetName, op.start, op.count),
        ];
      }
      case "moveColumns":
        return [
          {
            kind: "moveColumns",
            sheetName: op.sheetName,
            start: op.target,
            count: op.count,
            target: op.start,
          },
        ];
      case "updateRowMetadata": {
        const existing = this.workbook.getRowMetadata(op.sheetName, op.start, op.count);
        return [
          {
            kind: "updateRowMetadata",
            sheetName: op.sheetName,
            start: op.start,
            count: op.count,
            size: existing?.size ?? null,
            hidden: existing?.hidden ?? null,
          },
        ];
      }
      case "updateColumnMetadata": {
        const existing = this.workbook.getColumnMetadata(op.sheetName, op.start, op.count);
        return [
          {
            kind: "updateColumnMetadata",
            sheetName: op.sheetName,
            start: op.start,
            count: op.count,
            size: existing?.size ?? null,
            hidden: existing?.hidden ?? null,
          },
        ];
      }
      case "setFreezePane": {
        const existing = this.workbook.getFreezePane(op.sheetName);
        if (!existing) {
          return [{ kind: "clearFreezePane", sheetName: op.sheetName }];
        }
        return [
          {
            kind: "setFreezePane",
            sheetName: op.sheetName,
            rows: existing.rows,
            cols: existing.cols,
          },
        ];
      }
      case "clearFreezePane": {
        const existing = this.workbook.getFreezePane(op.sheetName);
        if (!existing) {
          return [];
        }
        return [
          {
            kind: "setFreezePane",
            sheetName: op.sheetName,
            rows: existing.rows,
            cols: existing.cols,
          },
        ];
      }
      case "setFilter": {
        const existing = this.workbook.getFilter(op.sheetName, op.range);
        if (!existing) {
          return [{ kind: "clearFilter", sheetName: op.sheetName, range: { ...op.range } }];
        }
        return [{ kind: "setFilter", sheetName: op.sheetName, range: { ...existing.range } }];
      }
      case "clearFilter": {
        const existing = this.workbook.getFilter(op.sheetName, op.range);
        if (!existing) {
          return [];
        }
        return [{ kind: "setFilter", sheetName: op.sheetName, range: { ...existing.range } }];
      }
      case "setSort": {
        const existing = this.workbook.getSort(op.sheetName, op.range);
        if (!existing) {
          return [{ kind: "clearSort", sheetName: op.sheetName, range: { ...op.range } }];
        }
        return [
          {
            kind: "setSort",
            sheetName: op.sheetName,
            range: { ...existing.range },
            keys: existing.keys.map((key) => Object.assign({}, key)),
          },
        ];
      }
      case "clearSort": {
        const existing = this.workbook.getSort(op.sheetName, op.range);
        if (!existing) {
          return [];
        }
        return [
          {
            kind: "setSort",
            sheetName: op.sheetName,
            range: { ...existing.range },
            keys: existing.keys.map((key) => Object.assign({}, key)),
          },
        ];
      }
      case "setCellValue":
      case "setCellFormula":
      case "clearCell":
        return this.restoreCellOps(op.sheetName, op.address);
      case "setCellFormat": {
        const cellIndex = this.workbook.getCellIndex(op.sheetName, op.address);
        return [
          {
            kind: "setCellFormat",
            sheetName: op.sheetName,
            address: op.address,
            format:
              cellIndex === undefined ? null : (this.workbook.getCellFormat(cellIndex) ?? null),
          },
        ];
      }
      case "upsertCellStyle": {
        const existing = this.workbook.getCellStyle(op.style.id);
        if (!existing || existing.id !== op.style.id) {
          return [];
        }
        return [{ kind: "upsertCellStyle", style: cloneCellStyleRecord(existing) }];
      }
      case "upsertCellNumberFormat": {
        const existing = this.workbook.getCellNumberFormat(op.format.id);
        if (!existing || existing.id !== op.format.id) {
          return [];
        }
        return [{ kind: "upsertCellNumberFormat", format: { ...existing } }];
      }
      case "setStyleRange":
        return this.restoreStyleRangeOps(op.range);
      case "setFormatRange":
        return this.restoreFormatRangeOps(op.range);
      case "upsertDefinedName": {
        const existing = this.workbook.getDefinedName(op.name);
        if (!existing) {
          return [{ kind: "deleteDefinedName", name: op.name }];
        }
        return [{ kind: "upsertDefinedName", name: existing.name, value: existing.value }];
      }
      case "deleteDefinedName": {
        const existing = this.workbook.getDefinedName(op.name);
        if (!existing) {
          return [];
        }
        return [{ kind: "upsertDefinedName", name: existing.name, value: existing.value }];
      }
      case "upsertTable": {
        const existing = this.workbook.getTable(op.table.name);
        if (!existing) {
          return [{ kind: "deleteTable", name: op.table.name }];
        }
        return [
          {
            kind: "upsertTable",
            table: {
              name: existing.name,
              sheetName: existing.sheetName,
              startAddress: existing.startAddress,
              endAddress: existing.endAddress,
              columnNames: [...existing.columnNames],
              headerRow: existing.headerRow,
              totalsRow: existing.totalsRow,
            },
          },
        ];
      }
      case "deleteTable": {
        const existing = this.workbook.getTable(op.name);
        if (!existing) {
          return [];
        }
        return [
          {
            kind: "upsertTable",
            table: {
              name: existing.name,
              sheetName: existing.sheetName,
              startAddress: existing.startAddress,
              endAddress: existing.endAddress,
              columnNames: [...existing.columnNames],
              headerRow: existing.headerRow,
              totalsRow: existing.totalsRow,
            },
          },
        ];
      }
      case "upsertSpillRange": {
        const existing = this.workbook.getSpill(op.sheetName, op.address);
        if (!existing) {
          return [{ kind: "deleteSpillRange", sheetName: op.sheetName, address: op.address }];
        }
        return [
          {
            kind: "upsertSpillRange",
            sheetName: existing.sheetName,
            address: existing.address,
            rows: existing.rows,
            cols: existing.cols,
          },
        ];
      }
      case "deleteSpillRange": {
        const existing = this.workbook.getSpill(op.sheetName, op.address);
        if (!existing) {
          return [];
        }
        return [
          {
            kind: "upsertSpillRange",
            sheetName: existing.sheetName,
            address: existing.address,
            rows: existing.rows,
            cols: existing.cols,
          },
        ];
      }
      case "upsertPivotTable": {
        const existing = this.workbook.getPivot(op.sheetName, op.address);
        if (!existing) {
          return [{ kind: "deletePivotTable", sheetName: op.sheetName, address: op.address }];
        }
        return [
          {
            kind: "upsertPivotTable",
            name: existing.name,
            sheetName: existing.sheetName,
            address: existing.address,
            source: { ...existing.source },
            groupBy: [...existing.groupBy],
            values: existing.values.map((v) => Object.assign({}, v)),
            rows: existing.rows,
            cols: existing.cols,
          },
        ];
      }
      case "deletePivotTable": {
        const existing = this.workbook.getPivot(op.sheetName, op.address);
        if (!existing) {
          return [];
        }
        return [
          {
            kind: "upsertPivotTable",
            name: existing.name,
            sheetName: existing.sheetName,
            address: existing.address,
            source: { ...existing.source },
            groupBy: [...existing.groupBy],
            values: existing.values.map((value) => Object.assign({}, value)),
            rows: existing.rows,
            cols: existing.cols,
          },
        ];
      }
    }
    return assertNever(op);
  }

  private restoreCellOps(sheetName: string, address: string): EngineOp[] {
    const cellIndex = this.workbook.getCellIndex(sheetName, address);
    if (cellIndex === undefined) {
      return [{ kind: "clearCell", sheetName, address }];
    }
    return this.toCellStateOps(sheetName, address, this.getCellByIndex(cellIndex)).filter(
      (op) => op.kind !== "setCellFormat",
    );
  }

  private readRangeCells(range: CellRangeRef): CellSnapshot[][] {
    const bounds = normalizeRange(range);
    const rows: CellSnapshot[][] = [];
    for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
      const cells: CellSnapshot[] = [];
      for (let col = bounds.startCol; col <= bounds.endCol; col += 1) {
        cells.push(this.getCell(range.sheetName, formatAddress(row, col)));
      }
      rows.push(cells);
    }
    return rows;
  }

  private buildStylePatchOps(range: CellRangeRef, patch: CellStylePatch): EngineOp[] {
    const normalizedPatch = normalizeCellStylePatch(patch);
    if (
      !normalizedPatch.fill &&
      !normalizedPatch.font &&
      !normalizedPatch.alignment &&
      !normalizedPatch.borders
    ) {
      return [];
    }
    return this.materializeStyleRangeOps(range, (baseStyle) =>
      this.workbook.internCellStyle(applyStylePatch(baseStyle, normalizedPatch)),
    );
  }

  private buildStyleClearOps(range: CellRangeRef, fields?: readonly CellStyleField[]): EngineOp[] {
    return this.materializeStyleRangeOps(range, (baseStyle) =>
      this.workbook.internCellStyle(clearStyleFields(baseStyle, fields)),
    );
  }

  private restoreStyleRangeOps(range: CellRangeRef): EngineOp[] {
    return this.materializeStyleRangeOps(range, (baseStyle, currentStyleId) => ({
      id: currentStyleId,
      ...baseStyle,
    }));
  }

  private buildFormatPatchOps(range: CellRangeRef, format: CellNumberFormatInput): EngineOp[] {
    const normalized = this.workbook.internCellNumberFormat(
      typeof format === "string"
        ? format
        : createCellNumberFormatRecord(WorkbookStore.defaultFormatId, format),
    );
    return this.materializeFormatRangeOps(range, () => normalized.id, normalized);
  }

  private buildFormatClearOps(range: CellRangeRef): EngineOp[] {
    return this.materializeFormatRangeOps(range, () => WorkbookStore.defaultFormatId);
  }

  private restoreFormatRangeOps(range: CellRangeRef): EngineOp[] {
    return this.materializeFormatRangeOps(range, (_currentFormatId, tile) => tile.formatId);
  }

  private materializeStyleRangeOps(
    range: CellRangeRef,
    resolveStyle: (
      baseStyle: Omit<CellStyleRecord, "id">,
      currentStyleId: string,
    ) => CellStyleRecord,
  ): EngineOp[] {
    const tiles = this.resolveStyleTiles(range);
    const ops: EngineOp[] = [];
    tiles.forEach((tile) => {
      const current = this.workbook.getCellStyle(tile.styleId) ?? {
        id: WorkbookStore.defaultStyleId,
      };
      const next = resolveStyle(current, tile.styleId);
      const normalizedId = next.id || WorkbookStore.defaultStyleId;
      if (normalizedId === tile.styleId) {
        return;
      }
      if (normalizedId !== WorkbookStore.defaultStyleId) {
        ops.push({
          kind: "upsertCellStyle",
          style: cloneCellStyleRecord(next),
        });
      }
      ops.push({
        kind: "setStyleRange",
        range: tile.range,
        styleId: normalizedId,
      });
    });
    return ops;
  }

  private materializeFormatRangeOps(
    range: CellRangeRef,
    resolveFormatId: (
      currentFormatId: string,
      tile: { range: CellRangeRef; formatId: string },
    ) => string,
    upsertFormat?: CellNumberFormatRecord,
  ): EngineOp[] {
    const tiles = this.resolveFormatTiles(range);
    const ops: EngineOp[] = [];
    if (upsertFormat && upsertFormat.id !== WorkbookStore.defaultFormatId) {
      ops.push({ kind: "upsertCellNumberFormat", format: { ...upsertFormat } });
    }
    tiles.forEach((tile) => {
      const nextFormatId = resolveFormatId(tile.formatId, tile);
      if (nextFormatId === tile.formatId) {
        return;
      }
      ops.push({
        kind: "setFormatRange",
        range: tile.range,
        formatId: nextFormatId,
      });
    });
    return ops;
  }

  private resolveStyleTiles(range: CellRangeRef): Array<{ range: CellRangeRef; styleId: string }> {
    const bounds = normalizeRange(range);
    const sheetRanges = this.workbook.listStyleRanges(range.sheetName);
    const rowBoundaries = new Set<number>([bounds.startRow, bounds.endRow + 1]);
    const colBoundaries = new Set<number>([bounds.startCol, bounds.endCol + 1]);

    sheetRanges.forEach((record) => {
      const clipped = intersectRangeBounds(record.range, bounds);
      if (!clipped) {
        return;
      }
      rowBoundaries.add(clipped.startRow);
      rowBoundaries.add(clipped.endRow + 1);
      colBoundaries.add(clipped.startCol);
      colBoundaries.add(clipped.endCol + 1);
    });

    const rows = [...rowBoundaries].toSorted((left, right) => left - right);
    const cols = [...colBoundaries].toSorted((left, right) => left - right);
    const tiles: Array<{ range: CellRangeRef; styleId: string }> = [];

    for (let rowIndex = 0; rowIndex < rows.length - 1; rowIndex += 1) {
      const startRow = rows[rowIndex]!;
      const endRow = rows[rowIndex + 1]! - 1;
      for (let colIndex = 0; colIndex < cols.length - 1; colIndex += 1) {
        const startCol = cols[colIndex]!;
        const endCol = cols[colIndex + 1]! - 1;
        tiles.push({
          range: {
            sheetName: range.sheetName,
            startAddress: formatAddress(startRow, startCol),
            endAddress: formatAddress(endRow, endCol),
          },
          styleId: this.workbook.getStyleId(range.sheetName, startRow, startCol),
        });
      }
    }

    return tiles;
  }

  private resolveFormatTiles(
    range: CellRangeRef,
  ): Array<{ range: CellRangeRef; formatId: string }> {
    const bounds = normalizeRange(range);
    const sheetRanges = this.workbook.listFormatRanges(range.sheetName);
    const rowBoundaries = new Set<number>([bounds.startRow, bounds.endRow + 1]);
    const colBoundaries = new Set<number>([bounds.startCol, bounds.endCol + 1]);

    sheetRanges.forEach((record) => {
      const clipped = intersectRangeBounds(record.range, bounds);
      if (!clipped) {
        return;
      }
      rowBoundaries.add(clipped.startRow);
      rowBoundaries.add(clipped.endRow + 1);
      colBoundaries.add(clipped.startCol);
      colBoundaries.add(clipped.endCol + 1);
    });

    const rows = [...rowBoundaries].toSorted((left, right) => left - right);
    const cols = [...colBoundaries].toSorted((left, right) => left - right);
    const tiles: Array<{ range: CellRangeRef; formatId: string }> = [];

    for (let rowIndex = 0; rowIndex < rows.length - 1; rowIndex += 1) {
      const startRow = rows[rowIndex]!;
      const endRow = rows[rowIndex + 1]! - 1;
      for (let colIndex = 0; colIndex < cols.length - 1; colIndex += 1) {
        const startCol = cols[colIndex]!;
        const endCol = cols[colIndex + 1]! - 1;
        tiles.push({
          range: {
            sheetName: range.sheetName,
            startAddress: formatAddress(startRow, startCol),
            endAddress: formatAddress(endRow, endCol),
          },
          formatId: this.workbook.getRangeFormatId(range.sheetName, startRow, startCol),
        });
      }
    }

    return tiles;
  }

  private toCellStateOps(
    sheetName: string,
    address: string,
    snapshot: CellSnapshot,
    sourceSheetName?: string,
    sourceAddress?: string,
  ): EngineOp[] {
    const ops: EngineOp[] = [];
    if (snapshot.formula !== undefined) {
      const translatedFormula =
        sourceSheetName && sourceAddress
          ? this.translateFormulaForTarget(
              snapshot.formula,
              sourceSheetName,
              sourceAddress,
              sheetName,
              address,
            )
          : snapshot.formula;
      ops.push({ kind: "setCellFormula", sheetName, address, formula: translatedFormula });
    } else {
      switch (snapshot.value.tag) {
        case ValueTag.Empty:
          ops.push({ kind: "clearCell", sheetName, address });
          break;
        case ValueTag.Number:
          ops.push({ kind: "setCellValue", sheetName, address, value: snapshot.value.value });
          break;
        case ValueTag.Boolean:
          ops.push({ kind: "setCellValue", sheetName, address, value: snapshot.value.value });
          break;
        case ValueTag.String:
          ops.push({ kind: "setCellValue", sheetName, address, value: snapshot.value.value });
          break;
        case ValueTag.Error:
          ops.push({ kind: "clearCell", sheetName, address });
          break;
      }
    }
    ops.push({
      kind: "setCellFormat",
      sheetName,
      address,
      format: snapshot.format ?? null,
    });
    return ops;
  }

  private translateFormulaForTarget(
    formula: string,
    sourceSheetName: string,
    sourceAddress: string,
    targetSheetName: string,
    targetAddress: string,
  ): string {
    const source = parseCellAddress(sourceAddress, sourceSheetName);
    const target = parseCellAddress(targetAddress, targetSheetName);
    return translateFormulaReferences(formula, target.row - source.row, target.col - source.col);
  }

  private materializeDependencies(
    currentSheet: string,
    compiled: ReturnType<typeof compileFormula>,
  ): MaterializedDependencies {
    const deps = compiled.deps;
    this.ensureDependencyBuildCapacity(
      this.workbook.cellStore.size + 1,
      deps.length + 1,
      compiled.symbolicRefs.length + 1,
      compiled.symbolicRanges.length + 1,
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
    this.symbolicRangeBindings.fill(UNRESOLVED_WASM_OPERAND, 0, compiled.symbolicRanges.length);
    for (const dep of deps) {
      if (dep.includes(":")) {
        const range = parseRangeAddress(dep, currentSheet);
        const sheetName = range.sheetName ?? currentSheet;
        const symbolicRangeIndex = compiled.symbolicRanges.indexOf(dep);
        if (range.sheetName && !this.workbook.getSheet(sheetName)) {
          continue;
        }
        const sheet = this.workbook.getSheet(sheetName);
        if (!sheet) {
          continue;
        }
        const registered = this.ranges.intern(sheet.id, range, {
          ensureCell: (sheetId, row, col) => this.ensureCellTrackedByCoords(sheetId, row, col),
          forEachSheetCell: (sheetId, fn) => this.forEachSheetCell(sheetId, fn),
        });
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
      newRangeCount,
    };
  }

  private setFormula(
    cellIndex: number,
    source: string,
    compiled: ReturnType<typeof compileFormula>,
    dependencies: MaterializedDependencies,
  ): void {
    this.removeFormula(cellIndex);

    this.ensureDependencyBuildCapacity(
      this.workbook.cellStore.size + 1,
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
        this.workbook.getSheetNameById(this.workbook.cellStore.sheetIds[cellIndex]!);
      if (qualifiedSheetName && !this.workbook.getSheet(sheetName)) {
        this.symbolicRefBindings[index] = UNRESOLVED_WASM_OPERAND;
        continue;
      }
      this.symbolicRefBindings[index] = this.ensureCellTracked(sheetName, parsed.text);
    }

    const literalStringIds = compiled.symbolicStrings.map((value) => this.strings.intern(value));
    const runtimeProgram = new Uint32Array(compiled.program.length);
    runtimeProgram.set(compiled.program);
    compiled.program.forEach((instruction, index) => {
      const opcode = (instruction >>> 24) as Opcode;
      const operand = instruction & 0x00ff_ffff;
      if (opcode === Opcode.PushCell) {
        const targetIndex =
          operand < compiled.symbolicRefs.length ? (this.symbolicRefBindings[operand] ?? 0) : 0;
        runtimeProgram[index] = (opcode << 24) | (targetIndex & 0x00ff_ffff);
        return;
      }
      if (opcode === Opcode.PushRange) {
        const targetIndex =
          operand < dependencies.symbolicRangeCount
            ? (dependencies.symbolicRangeIndices[operand] ?? 0)
            : 0;
        runtimeProgram[index] = (opcode << 24) | (targetIndex & 0x00ff_ffff);
        return;
      }
      if (opcode === Opcode.PushString) {
        const stringId = operand < literalStringIds.length ? (literalStringIds[operand] ?? 0) : 0;
        runtimeProgram[index] = (opcode << 24) | (stringId & 0x00ff_ffff);
      }
    });

    const dependencyEntities = this.edgeArena.replace(
      this.edgeArena.empty(),
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
    const formulaId = this.formulas.set(cellIndex, runtimeFormula);
    runtimeFormula.compiled.id = formulaId;
    this.workbook.cellStore.flags[cellIndex] =
      ((this.workbook.cellStore.flags[cellIndex] ?? 0) &
        ~(CellFlags.SpillChild | CellFlags.PivotOutput)) |
      CellFlags.HasFormula;
    if (runtimeFormula.compiled.mode === FormulaMode.JsOnly) {
      this.workbook.cellStore.flags[cellIndex] =
        (this.workbook.cellStore.flags[cellIndex] ?? 0) | CellFlags.JsOnly;
    } else {
      this.workbook.cellStore.flags[cellIndex] =
        (this.workbook.cellStore.flags[cellIndex] ?? 0) & ~CellFlags.JsOnly;
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
    runtimeFormula.compiled.symbolicNames.forEach((name) => {
      this.appendDefinedNameReverseEdge(name, cellIndex);
    });
    runtimeFormula.compiled.symbolicTables.forEach((name) => {
      this.appendTrackedReverseEdge(this.reverseTableEdges, tableDependencyKey(name), cellIndex);
    });
    runtimeFormula.compiled.symbolicSpills.forEach((key) => {
      this.appendTrackedReverseEdge(
        this.reverseSpillEdges,
        spillDependencyKeyFromRef(
          key,
          this.workbook.getSheetNameById(this.workbook.cellStore.sheetIds[cellIndex]!),
        ),
        cellIndex,
      );
    });
    this.scheduleWasmProgramSync();
  }

  private setInvalidFormulaValue(cellIndex: number): void {
    this.removeFormula(cellIndex);
    this.workbook.cellStore.setValue(cellIndex, errorValue(ErrorCode.Value));
    this.workbook.cellStore.flags[cellIndex] =
      (this.workbook.cellStore.flags[cellIndex] ?? 0) &
      ~(
        CellFlags.HasFormula |
        CellFlags.JsOnly |
        CellFlags.InCycle |
        CellFlags.SpillChild |
        CellFlags.PivotOutput
      );
  }

  private removeFormula(cellIndex: number): boolean {
    const existing = this.formulas.get(cellIndex);
    if (existing) {
      const dependencyEntities = this.edgeArena.readView(existing.dependencyEntities);
      const formulaEntity = makeCellEntity(cellIndex);
      for (let index = 0; index < dependencyEntities.length; index += 1) {
        this.removeReverseEdge(dependencyEntities[index]!, formulaEntity);
      }
      existing.compiled.symbolicNames.forEach((name) => {
        this.removeDefinedNameReverseEdge(name, cellIndex);
      });
      existing.compiled.symbolicTables.forEach((name) => {
        this.removeTrackedReverseEdge(this.reverseTableEdges, tableDependencyKey(name), cellIndex);
      });
      const ownerSheetName = this.workbook.getSheetNameById(
        this.workbook.cellStore.sheetIds[cellIndex]!,
      );
      existing.compiled.symbolicSpills.forEach((key) => {
        this.removeTrackedReverseEdge(
          this.reverseSpillEdges,
          spillDependencyKeyFromRef(key, ownerSheetName),
          cellIndex,
        );
      });
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
      (this.workbook.cellStore.flags[cellIndex] ?? 0) &
      ~(
        CellFlags.HasFormula |
        CellFlags.JsOnly |
        CellFlags.InCycle |
        CellFlags.SpillChild |
        CellFlags.PivotOutput
      );
    this.scheduleWasmProgramSync();
    return existing !== undefined;
  }

  private rebuildTopoRanks(): void {
    const requiredCellCapacity = this.workbook.cellStore.size + 1;
    const requiredEntityCapacity = this.workbook.cellStore.size + this.ranges.size + 1;
    this.ensureTopoScratchCapacity(
      requiredCellCapacity,
      requiredEntityCapacity,
      this.ranges.size + 1,
    );

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
      (cellIndex) => this.formulas.has(cellIndex),
    );

    this.formulas.forEach((_formula, cellIndex) => {
      this.workbook.cellStore.flags[cellIndex] =
        (this.workbook.cellStore.flags[cellIndex] ?? 0) & ~CellFlags.InCycle;
      this.workbook.cellStore.cycleGroupIds[cellIndex] = -1;
    });

    for (let index = 0; index < result.cycleMemberCount; index += 1) {
      const cellIndex = result.cycleMembers[index]!;
      this.workbook.cellStore.flags[cellIndex] =
        (this.workbook.cellStore.flags[cellIndex] ?? 0) | CellFlags.InCycle;
      this.workbook.cellStore.cycleGroupIds[cellIndex] = result.cycleGroups[cellIndex] ?? -1;
      this.workbook.cellStore.setValue(cellIndex, errorValue(ErrorCode.Cycle));
    }
  }

  private recalculate(
    changedRoots: readonly number[] | U32,
    kernelSyncRoots: readonly number[] | U32 = changedRoots,
  ): Uint32Array {
    const started = performance.now();
    this.ensureRecalcScratchCapacity(this.workbook.cellStore.size + 1);
    if (this.wasm.ready) {
      this.wasm.syncStringPool(this.strings.exportLayout());
    }

    const allChangedRoots = [...changedRoots];
    const allOrdered: number[] = [];
    let singlePassOrdered: U32 | null = null;
    let singlePassOrderedCount = 0;
    let passRoots = [...changedRoots];
    let passKernelRoots = [...kernelSyncRoots];
    let totalOrderedCount = 0;
    let totalRangeNodeVisits = 0;
    let wasmCount = 0;
    let jsCount = 0;
    let pendingKernelSyncCount = 0;
    const volatileState = this.createRecalcVolatileState();

    const flushWasmBatch = (
      batchCount: number,
      hasVolatile: boolean,
      randCount: number,
    ): number => {
      if (batchCount === 0) {
        return 0;
      }
      this.wasm.syncFromStore(
        this.workbook.cellStore,
        this.pendingKernelSync.subarray(0, pendingKernelSyncCount),
      );
      pendingKernelSyncCount = 0;
      if (hasVolatile) {
        this.wasm.uploadVolatileNowSerial(volatileState.nowSerial);
        this.wasm.uploadVolatileRandomValues(
          this.consumeVolatileRandomValues(volatileState, randCount),
        );
      }
      const batchIndices = this.wasmBatch.subarray(0, batchCount);
      this.wasm.evalBatch(batchIndices);
      this.wasm.syncToStore(this.workbook.cellStore, batchIndices, this.strings);
      return batchCount;
    };

    while (passRoots.length > 0) {
      const scheduled = this.scheduler.collectDirty(
        passRoots,
        { getDependents: (entityId) => this.getEntityDependents(entityId) },
        this.workbook.cellStore,
        (cellIndex) => this.formulas.has(cellIndex),
        this.ranges.size,
      );
      const ordered = scheduled.orderedFormulaCellIndices;
      const orderedCount = scheduled.orderedFormulaCount;
      totalOrderedCount += orderedCount;
      totalRangeNodeVisits += scheduled.rangeNodeVisits;
      if (singlePassOrdered === null && allOrdered.length === 0) {
        singlePassOrdered = ordered;
        singlePassOrderedCount = orderedCount;
      } else {
        if (singlePassOrdered !== null) {
          const orderedChunk = singlePassOrdered;
          for (let orderedIndex = 0; orderedIndex < singlePassOrderedCount; orderedIndex += 1) {
            const cellIndex = orderedChunk[orderedIndex];
            if (cellIndex === undefined) {
              continue;
            }
            allOrdered.push(cellIndex);
          }
          singlePassOrdered = null;
          singlePassOrderedCount = 0;
        }
        for (let orderedIndex = 0; orderedIndex < orderedCount; orderedIndex += 1) {
          allOrdered.push(ordered[orderedIndex]!);
        }
      }

      pendingKernelSyncCount = 0;
      for (let index = 0; index < passKernelRoots.length; index += 1) {
        this.pendingKernelSync[pendingKernelSyncCount] = passKernelRoots[index]!;
        pendingKernelSyncCount += 1;
      }

      let wasmBatchCount = 0;
      let wasmBatchHasVolatile = false;
      let wasmBatchRandCount = 0;
      const spillChangedRoots: number[] = [];
      const spillChangedSeen = new Set<number>();
      const noteSpillChanges = (changedCellIndices: readonly number[]): void => {
        for (let spillIndex = 0; spillIndex < changedCellIndices.length; spillIndex += 1) {
          const changedCellIndex = changedCellIndices[spillIndex]!;
          if (spillChangedSeen.has(changedCellIndex)) {
            continue;
          }
          spillChangedSeen.add(changedCellIndex);
          spillChangedRoots.push(changedCellIndex);
        }
      };
      const queueKernelSync = (cellIndex: number): void => {
        this.pendingKernelSync[pendingKernelSyncCount] = cellIndex;
        pendingKernelSyncCount += 1;
      };
      const evaluateWasmSpillFormula = (cellIndex: number, formula: RuntimeFormula): number => {
        this.wasm.syncFromStore(
          this.workbook.cellStore,
          this.pendingKernelSync.subarray(0, pendingKernelSyncCount),
        );
        pendingKernelSyncCount = 0;
        if (formula.compiled.volatile) {
          this.wasm.uploadVolatileNowSerial(volatileState.nowSerial);
          this.wasm.uploadVolatileRandomValues(
            this.consumeVolatileRandomValues(volatileState, formula.compiled.randCallCount),
          );
        }
        const batchIndices = Uint32Array.of(cellIndex);
        this.wasm.evalBatch(batchIndices);
        this.wasm.syncToStore(this.workbook.cellStore, batchIndices, this.strings);
        const spill = this.wasm.readSpill(cellIndex, this.strings);
        const spillMaterialization = spill
          ? this.materializeSpill(cellIndex, {
              rows: spill.rows,
              cols: spill.cols,
              values: spill.values,
            })
          : {
              changedCellIndices: this.clearOwnedSpill(cellIndex),
              ownerValue: this.workbook.cellStore.getValue(cellIndex, (id) => this.strings.get(id)),
            };
        const currentFlags =
          (this.workbook.cellStore.flags[cellIndex] ?? 0) &
          ~(CellFlags.SpillChild | CellFlags.PivotOutput);
        this.workbook.cellStore.setValue(
          cellIndex,
          spillMaterialization.ownerValue,
          spillMaterialization.ownerValue.tag === ValueTag.String
            ? this.strings.intern(spillMaterialization.ownerValue.value)
            : 0,
        );
        this.workbook.cellStore.flags[cellIndex] = currentFlags;
        queueKernelSync(cellIndex);
        for (
          let spillIndex = 0;
          spillIndex < spillMaterialization.changedCellIndices.length;
          spillIndex += 1
        ) {
          queueKernelSync(spillMaterialization.changedCellIndices[spillIndex]!);
        }
        noteSpillChanges(spillMaterialization.changedCellIndices);
        return 1;
      };
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
          if (formula.compiled.producesSpill) {
            wasmCount += flushWasmBatch(wasmBatchCount, wasmBatchHasVolatile, wasmBatchRandCount);
            wasmBatchCount = 0;
            wasmBatchHasVolatile = false;
            wasmBatchRandCount = 0;
            wasmCount += evaluateWasmSpillFormula(cellIndex, formula);
            continue;
          }
          this.wasmBatch[wasmBatchCount] = cellIndex;
          wasmBatchCount += 1;
          wasmBatchHasVolatile = wasmBatchHasVolatile || formula.compiled.volatile;
          wasmBatchRandCount += formula.compiled.randCallCount;
          continue;
        }
        wasmCount += flushWasmBatch(wasmBatchCount, wasmBatchHasVolatile, wasmBatchRandCount);
        wasmBatchCount = 0;
        wasmBatchHasVolatile = false;
        wasmBatchRandCount = 0;
        jsCount += 1;
        const spillChanges = this.evaluateUnsupportedFormula(cellIndex);
        noteSpillChanges(spillChanges);
        queueKernelSync(cellIndex);
      }

      wasmCount += flushWasmBatch(wasmBatchCount, wasmBatchHasVolatile, wasmBatchRandCount);
      if (pendingKernelSyncCount > 0) {
        this.wasm.syncFromStore(
          this.workbook.cellStore,
          this.pendingKernelSync.subarray(0, pendingKernelSyncCount),
        );
      }

      if (spillChangedRoots.length === 0) {
        break;
      }
      if (singlePassOrdered !== null) {
        const orderedChunk = singlePassOrdered;
        for (let orderedIndex = 0; orderedIndex < singlePassOrderedCount; orderedIndex += 1) {
          const cellIndex = orderedChunk[orderedIndex];
          if (cellIndex === undefined) {
            continue;
          }
          allOrdered.push(cellIndex);
        }
        singlePassOrdered = null;
        singlePassOrderedCount = 0;
      }
      allChangedRoots.push(...spillChangedRoots);
      passRoots = spillChangedRoots;
      passKernelRoots = spillChangedRoots;
    }

    this.lastMetrics.dirtyFormulaCount = totalOrderedCount;
    this.lastMetrics.jsFormulaCount = jsCount;
    this.lastMetrics.wasmFormulaCount = wasmCount;
    this.lastMetrics.rangeNodeVisits = totalRangeNodeVisits;
    this.lastMetrics.recalcMs = performance.now() - started;
    if (singlePassOrdered !== null) {
      return totalOrderedCount === 0 && allChangedRoots.length === 0
        ? this.changedUnion.subarray(0, 0)
        : this.composeChangedRootsAndOrdered(
            allChangedRoots,
            singlePassOrdered,
            singlePassOrderedCount,
          );
    }
    return totalOrderedCount === 0 && allChangedRoots.length === 0
      ? this.changedUnion.subarray(0, 0)
      : this.composeChangedRootsAndOrdered(
          allChangedRoots,
          Uint32Array.from(allOrdered),
          allOrdered.length,
        );
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
    symbolicRangeCapacity = 0,
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
    if (rangeSize > this.wasmRangeRowCounts.length) {
      this.wasmRangeRowCounts = growUint32(this.wasmRangeRowCounts, rangeSize);
    }
    if (rangeSize > this.wasmRangeColCounts.length) {
      this.wasmRangeColCounts = growUint32(this.wasmRangeColCounts, rangeSize);
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

  private createRecalcVolatileState(): RecalcVolatileState {
    return {
      nowSerial: utcDateToExcelSerial(new Date()),
      randomValues: [],
      randomCursor: 0,
    };
  }

  private ensureVolatileRandomValues(state: RecalcVolatileState, count: number): void {
    const required = state.randomCursor + count;
    while (state.randomValues.length < required) {
      state.randomValues.push(Math.random());
    }
  }

  private consumeVolatileRandomValues(state: RecalcVolatileState, count: number): Float64Array {
    if (count === 0) {
      return new Float64Array();
    }
    this.ensureVolatileRandomValues(state, count);
    const end = state.randomCursor + count;
    const slice = state.randomValues.slice(state.randomCursor, end);
    state.randomCursor = end;
    return Float64Array.from(slice);
  }

  private evaluateUnsupportedFormula(cellIndex: number): number[] {
    const formula = this.formulas.get(cellIndex);
    const sheetName = this.workbook.getSheetNameById(this.workbook.cellStore.sheetIds[cellIndex]!);
    if (!formula || !sheetName) {
      return [];
    }

    const evaluationContext = {
      sheetName,
      currentAddress: this.workbook.getAddress(cellIndex),
      resolveCell: (targetSheetName, address) => this.readCellValue(targetSheetName, address),
      resolveRange: (targetSheetName, start, end, refKind) =>
        this.readRangeValues(targetSheetName, start, end, refKind),
      resolveName: (name) => {
        const definedName = this.workbook.getDefinedName(name);
        if (!definedName) {
          return errorValue(ErrorCode.Name);
        }
        return definedNameValueToCellValue(definedName.value, this.strings);
      },
      resolveFormula: (targetSheetName: string, address: string) =>
        this.getCell(targetSheetName, address).formula,
      resolvePivotData: ({ dataField, sheetName: pivotSheetName, address, filters }) =>
        this.resolvePivotData(pivotSheetName, address, dataField, filters),
      resolveMultipleOperations: ({
        formulaSheetName,
        formulaAddress,
        rowCellSheetName,
        rowCellAddress,
        rowReplacementSheetName,
        rowReplacementAddress,
        columnCellSheetName,
        columnCellAddress,
        columnReplacementSheetName,
        columnReplacementAddress,
      }: {
        formulaSheetName: string;
        formulaAddress: string;
        rowCellSheetName: string;
        rowCellAddress: string;
        rowReplacementSheetName: string;
        rowReplacementAddress: string;
        columnCellSheetName?: string;
        columnCellAddress?: string;
        columnReplacementSheetName?: string;
        columnReplacementAddress?: string;
      }) =>
        this.resolveMultipleOperations({
          formulaSheetName,
          formulaAddress,
          rowCellSheetName,
          rowCellAddress,
          rowReplacementSheetName,
          rowReplacementAddress,
          ...(columnCellSheetName ? { columnCellSheetName } : {}),
          ...(columnCellAddress ? { columnCellAddress } : {}),
          ...(columnReplacementSheetName ? { columnReplacementSheetName } : {}),
          ...(columnReplacementAddress ? { columnReplacementAddress } : {}),
        }),
      listSheetNames: () =>
        [...this.workbook.sheetsByName.values()]
          .toSorted((left, right) => left.order - right.order)
          .map((sheet) => sheet.name),
    } as Parameters<typeof evaluatePlanResult>[1];
    const result = evaluatePlanResult(formula.compiled.jsPlan, evaluationContext);

    const materialization = isArrayValue(result)
      ? this.materializeSpill(cellIndex, result)
      : {
          changedCellIndices: this.clearOwnedSpill(cellIndex),
          ownerValue: result,
        };

    this.workbook.cellStore.flags[cellIndex] =
      (this.workbook.cellStore.flags[cellIndex] ?? 0) &
      ~(CellFlags.SpillChild | CellFlags.PivotOutput);
    this.workbook.cellStore.setValue(
      cellIndex,
      materialization.ownerValue,
      materialization.ownerValue.tag === ValueTag.String
        ? this.strings.intern(materialization.ownerValue.value)
        : 0,
    );
    return materialization.changedCellIndices;
  }

  private compileFormulaForSheet(
    currentSheetName: string,
    source: string,
  ): ReturnType<typeof compileFormula> {
    const compiled = compileFormula(source);
    if (
      compiled.symbolicNames.length === 0 &&
      compiled.symbolicTables.length === 0 &&
      compiled.symbolicSpills.length === 0
    ) {
      return compiled;
    }

    const resolved = resolveMetadataReferencesInAst(compiled.ast, {
      resolveName: (name) => this.workbook.getDefinedName(name)?.value,
      resolveStructuredReference: (tableName, columnName) =>
        this.resolveStructuredReference(tableName, columnName),
      resolveSpillReference: (sheetName, address) =>
        this.resolveSpillReference(currentSheetName, sheetName, address),
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
  }

  private resolveStructuredReference(
    tableName: string,
    columnName: string,
  ): FormulaNode | undefined {
    const table = this.workbook.getTable(tableName);
    if (!table) {
      return undefined;
    }
    const columnIndex = table.columnNames.findIndex(
      (name) => name.trim().toUpperCase() === columnName.trim().toUpperCase(),
    );
    if (columnIndex === -1) {
      return undefined;
    }
    const start = parseCellAddress(table.startAddress, table.sheetName);
    const end = parseCellAddress(table.endAddress, table.sheetName);
    const startRow = start.row + (table.headerRow ? 1 : 0);
    const endRow = end.row - (table.totalsRow ? 1 : 0);
    if (endRow < startRow) {
      return { kind: "ErrorLiteral", code: ErrorCode.Ref };
    }
    const column = start.col + columnIndex;
    return {
      kind: "RangeRef",
      refKind: "cells",
      sheetName: table.sheetName,
      start: formatAddress(startRow, column),
      end: formatAddress(endRow, column),
    };
  }

  private resolveSpillReference(
    currentSheetName: string,
    sheetName: string | undefined,
    address: string,
  ): FormulaNode | undefined {
    const targetSheetName = sheetName ?? currentSheetName;
    const spill = this.workbook.getSpill(targetSheetName, address);
    if (!spill) {
      return undefined;
    }
    const owner = parseCellAddress(address, targetSheetName);
    return {
      kind: "RangeRef",
      refKind: "cells",
      sheetName: targetSheetName,
      start: owner.text,
      end: formatAddress(owner.row + spill.rows - 1, owner.col + spill.cols - 1),
    };
  }

  private readCellValue(sheetName: string, address: string): CellValue {
    const cellIndex = this.workbook.getCellIndex(sheetName, address);
    if (cellIndex === undefined) {
      return emptyValue();
    }
    return this.workbook.cellStore.getValue(cellIndex, (id) => this.strings.get(id));
  }

  private readRangeValueMatrix(range: CellRangeRef): CellValue[][] {
    const bounds = normalizeRange(range);
    const width = bounds.endCol - bounds.startCol + 1;
    const height = bounds.endRow - bounds.startRow + 1;
    const rows = Array.from<CellValue[]>({ length: height });
    const sheet = this.workbook.getSheet(range.sheetName);

    for (let rowOffset = 0; rowOffset < height; rowOffset += 1) {
      const row = bounds.startRow + rowOffset;
      const values = Array.from<CellValue>({ length: width });
      for (let colOffset = 0; colOffset < width; colOffset += 1) {
        const col = bounds.startCol + colOffset;
        const cellIndex = sheet?.grid.get(row, col) ?? -1;
        values[colOffset] =
          cellIndex === -1
            ? emptyValue()
            : this.workbook.cellStore.getValue(cellIndex, (id) => this.strings.get(id));
      }
      rows[rowOffset] = values;
    }

    return rows;
  }

  private readRangeValues(
    sheetName: string,
    start: string,
    end: string,
    refKind: "cells" | "rows" | "cols",
  ): CellValue[] {
    if (refKind !== "cells") {
      return [];
    }
    const range = parseRangeAddress(`${start}:${end}`, sheetName);
    if (range.kind !== "cells") {
      return [];
    }
    const rows = this.readRangeValueMatrix({
      sheetName,
      startAddress: start,
      endAddress: end,
    });
    const values = Array.from<CellValue>({ length: rows.length * (rows[0]?.length ?? 0) });
    let valueIndex = 0;
    for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
      const row = rows[rowIndex]!;
      for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
        values[valueIndex] = row[colIndex]!;
        valueIndex += 1;
      }
    }
    return values;
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
    this.ensureWasmProgramScratchCapacity(
      Math.max(wasmFormulaCount, 1),
      Math.max(this.ranges.size, 1),
    );

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
      constantLengths: this.wasmConstantLengths.subarray(0, wasmFormulaCount),
    });

    const rangeCapacity = Math.max(this.ranges.size, 1);
    if (this.ranges.size === 0) {
      this.wasmRangeOffsets[0] = 0;
      this.wasmRangeLengths[0] = 0;
      this.wasmRangeRowCounts[0] = 0;
      this.wasmRangeColCounts[0] = 0;
    }
    for (let rangeIndex = 0; rangeIndex < this.ranges.size; rangeIndex += 1) {
      const descriptor = this.ranges.getDescriptor(rangeIndex);
      this.wasmRangeOffsets[rangeIndex] = descriptor.refCount > 0 ? descriptor.membersOffset : 0;
      this.wasmRangeLengths[rangeIndex] = descriptor.refCount > 0 ? descriptor.membersLength : 0;
      this.wasmRangeRowCounts[rangeIndex] =
        descriptor.refCount > 0 ? descriptor.row2 - descriptor.row1 + 1 : 0;
      this.wasmRangeColCounts[rangeIndex] =
        descriptor.refCount > 0 ? descriptor.col2 - descriptor.col1 + 1 : 0;
    }

    this.wasm.uploadRanges({
      members: this.ranges.getMemberPoolView(),
      offsets: this.wasmRangeOffsets.subarray(0, rangeCapacity),
      lengths: this.wasmRangeLengths.subarray(0, rangeCapacity),
      rowCounts: this.wasmRangeRowCounts.subarray(0, rangeCapacity),
      colCounts: this.wasmRangeColCounts.subarray(0, rangeCapacity),
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

  private appendDefinedNameReverseEdge(name: string, dependentCellIndex: number): void {
    this.appendTrackedReverseEdge(
      this.reverseDefinedNameEdges,
      normalizeDefinedName(name),
      dependentCellIndex,
    );
  }

  private removeDefinedNameReverseEdge(name: string, dependentCellIndex: number): void {
    this.removeTrackedReverseEdge(
      this.reverseDefinedNameEdges,
      normalizeDefinedName(name),
      dependentCellIndex,
    );
  }

  private appendTrackedReverseEdge(
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

  private removeTrackedReverseEdge(
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

  private forEachFormulaDependencyCell(
    cellIndex: number,
    fn: (dependencyCellIndex: number) => void,
  ): void {
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
      case "setWorkbookMetadata":
        return `workbook-meta:${op.key}`;
      case "setCalculationSettings":
        return "workbook-calc";
      case "setVolatileContext":
        return "workbook-volatile";
      case "upsertSheet":
      case "deleteSheet":
        return `sheet:${op.name}`;
      case "renameSheet":
        return `sheet:${op.oldName}`;
      case "insertRows":
      case "deleteRows":
      case "moveRows":
        return `row-structure:${op.sheetName}`;
      case "insertColumns":
      case "deleteColumns":
      case "moveColumns":
        return `column-structure:${op.sheetName}`;
      case "updateRowMetadata":
        return `row-meta:${op.sheetName}:${op.start}:${op.count}`;
      case "updateColumnMetadata":
        return `column-meta:${op.sheetName}:${op.start}:${op.count}`;
      case "setFreezePane":
      case "clearFreezePane":
        return `freeze:${op.sheetName}`;
      case "setFilter":
      case "clearFilter":
        return `filter:${op.sheetName}:${op.range.startAddress}:${op.range.endAddress}`;
      case "setSort":
      case "clearSort":
        return `sort:${op.sheetName}:${op.range.startAddress}:${op.range.endAddress}`;
      case "setCellFormat":
        return `format:${op.sheetName}!${op.address}`;
      case "upsertCellStyle":
        return `style:${op.style.id}`;
      case "upsertCellNumberFormat":
        return `number-format:${op.format.id}`;
      case "setStyleRange":
        return `style-range:${op.range.sheetName}:${op.range.startAddress}:${op.range.endAddress}`;
      case "setFormatRange":
        return `format-range:${op.range.sheetName}:${op.range.startAddress}:${op.range.endAddress}`;
      case "setCellValue":
      case "setCellFormula":
      case "clearCell":
        return `cell:${op.sheetName}!${op.address}`;
      case "upsertDefinedName":
      case "deleteDefinedName":
        return `defined-name:${normalizeDefinedName(op.name)}`;
      case "upsertTable":
        return `table:${normalizeDefinedName(op.table.name)}`;
      case "deleteTable":
        return `table:${normalizeDefinedName(op.name)}`;
      case "upsertSpillRange":
      case "deleteSpillRange":
        return `spill:${op.sheetName}!${op.address}`;
      case "upsertPivotTable":
      case "deletePivotTable":
        return `pivot:${pivotKey(op.sheetName, op.address)}`;
    }
    return assertNever(op);
  }

  private sheetDeleteBarrierForOp(op: EngineOp): OpOrder | undefined {
    switch (op.kind) {
      case "upsertWorkbook":
      case "setWorkbookMetadata":
      case "setCalculationSettings":
      case "setVolatileContext":
      case "deleteSheet":
      case "upsertDefinedName":
      case "deleteDefinedName":
      case "upsertTable":
      case "deleteTable":
        return undefined;
      case "updateRowMetadata":
      case "updateColumnMetadata":
      case "insertRows":
      case "deleteRows":
      case "moveRows":
      case "insertColumns":
      case "deleteColumns":
      case "moveColumns":
      case "setFreezePane":
      case "clearFreezePane":
      case "setFilter":
      case "clearFilter":
      case "setSort":
      case "clearSort":
      case "setCellFormat":
      case "setCellValue":
      case "setCellFormula":
      case "clearCell":
      case "upsertSpillRange":
      case "deleteSpillRange":
      case "deletePivotTable":
        return this.sheetDeleteVersions.get(op.sheetName);
      case "setStyleRange":
      case "setFormatRange":
        return this.sheetDeleteVersions.get(op.range.sheetName);
      case "upsertCellNumberFormat":
        return undefined;
      case "upsertCellStyle":
        return undefined;
      case "upsertSheet":
        return this.sheetDeleteVersions.get(op.name);
      case "renameSheet":
        return this.sheetDeleteVersions.get(op.oldName);
      case "upsertPivotTable":
        return (
          this.sheetDeleteVersions.get(op.sheetName) ??
          this.sheetDeleteVersions.get(op.source.sheetName)
        );
    }
  }

  private removeSheetRuntime(
    sheetName: string,
    explicitChangedCount: number,
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
      const nextSheet = [...this.workbook.sheetsByName.values()].toSorted(
        (left, right) => left.order - right.order,
      )[0];
      this.setSelection(nextSheet?.name ?? sheetName, "A1");
    }
    formulaChangedCount = this.rebindFormulasForSheet(
      sheetName,
      formulaChangedCount,
      this.impactedFormulaBuffer.subarray(0, impactedCount),
    );
    return { changedInputCount, formulaChangedCount, explicitChangedCount };
  }

  private rebindFormulasForSheet(
    sheetName: string,
    formulaChangedCount: number,
    candidates?: readonly number[] | U32,
  ): number {
    if (candidates) {
      for (let index = 0; index < candidates.length; index += 1) {
        const cellIndex = candidates[index]!;
        const formula = this.formulas.get(cellIndex);
        if (!formula) continue;
        const ownerSheetName = this.workbook.getSheetNameById(
          this.workbook.cellStore.sheetIds[cellIndex]!,
        );
        if (!ownerSheetName) continue;
        const touchesSheet = formula.compiled.deps.some((dep) => {
          if (!dep.includes("!")) return false;
          const [qualifiedSheet] = dep.split("!");
          return qualifiedSheet?.replace(/^'(.*)'$/, "$1") === sheetName;
        });
        if (!touchesSheet) continue;
        const compiled = this.compileFormulaForSheet(ownerSheetName, formula.source);
        const dependencies = this.materializeDependencies(ownerSheetName, compiled);
        this.setFormula(cellIndex, formula.source, compiled, dependencies);
        formulaChangedCount = this.markFormulaChanged(cellIndex, formulaChangedCount);
      }
      return formulaChangedCount;
    }

    this.formulas.forEach((formula, cellIndex) => {
      if (!formula) return;
      const ownerSheetName = this.workbook.getSheetNameById(
        this.workbook.cellStore.sheetIds[cellIndex]!,
      );
      if (!ownerSheetName) return;
      const touchesSheet = formula.compiled.deps.some((dep) => {
        if (!dep.includes("!")) return false;
        const [qualifiedSheet] = dep.split("!");
        return qualifiedSheet?.replace(/^'(.*)'$/, "$1") === sheetName;
      });
      if (!touchesSheet) return;
      const compiled = this.compileFormulaForSheet(ownerSheetName, formula.source);
      const dependencies = this.materializeDependencies(ownerSheetName, compiled);
      this.setFormula(cellIndex, formula.source, compiled, dependencies);
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

  private markVolatileFormulasChanged(count: number): number {
    this.formulas.forEach((formula, cellIndex) => {
      if (!formula.compiled.volatile) {
        return;
      }
      count = this.markFormulaChanged(cellIndex, count);
    });
    return count;
  }

  private markSpillRootsChanged(cellIndices: readonly number[], count: number): number {
    for (let index = 0; index < cellIndices.length; index += 1) {
      count = this.markInputChanged(cellIndices[index]!, count);
    }
    return count;
  }

  private markPivotRootsChanged(cellIndices: readonly number[], count: number): number {
    for (let index = 0; index < cellIndices.length; index += 1) {
      count = this.markInputChanged(cellIndices[index]!, count);
    }
    return count;
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

  private unionChangedSets(...sets: Array<readonly number[] | U32>): U32 {
    this.changedUnionEpoch += 1;
    if (this.changedUnionEpoch === 0xffff_ffff) {
      this.changedUnionEpoch = 1;
      this.changedUnionSeen.fill(0);
    }
    let changedCount = 0;
    for (let setIndex = 0; setIndex < sets.length; setIndex += 1) {
      const set = sets[setIndex]!;
      for (let index = 0; index < set.length; index += 1) {
        const cellIndex = set[index]!;
        if (this.changedUnionSeen[cellIndex] === this.changedUnionEpoch) {
          continue;
        }
        this.changedUnionSeen[cellIndex] = this.changedUnionEpoch;
        this.changedUnion[changedCount] = cellIndex;
        changedCount += 1;
      }
    }
    return this.changedUnion.subarray(0, changedCount);
  }

  private composeChangedRootsAndOrdered(
    changedRoots: readonly number[] | U32,
    ordered: U32,
    orderedCount: number,
  ): U32 {
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

  private clearOwnedSpill(cellIndex: number): number[] {
    const sheetName = this.workbook.getSheetNameById(this.workbook.cellStore.sheetIds[cellIndex]!);
    const address = this.workbook.getAddress(cellIndex);
    const spill = this.workbook.getSpill(sheetName, address);
    if (!spill) {
      return [];
    }

    const owner = parseCellAddress(address, sheetName);
    const changedCellIndices: number[] = [];
    for (let rowOffset = 0; rowOffset < spill.rows; rowOffset += 1) {
      for (let colOffset = 0; colOffset < spill.cols; colOffset += 1) {
        if (rowOffset === 0 && colOffset === 0) {
          continue;
        }
        const childAddress = formatAddress(owner.row + rowOffset, owner.col + colOffset);
        const childIndex = this.workbook.getCellIndex(sheetName, childAddress);
        if (childIndex === undefined) {
          continue;
        }
        if (this.clearSpillChildCell(childIndex)) {
          changedCellIndices.push(childIndex);
        }
      }
    }
    changedCellIndices.push(
      ...this.applyDerivedOp({ kind: "deleteSpillRange", sheetName, address }),
    );
    return changedCellIndices;
  }

  private materializeSpill(
    cellIndex: number,
    arrayValue: { values: CellValue[]; rows: number; cols: number },
  ): SpillMaterialization {
    const changedCellIndices = this.clearOwnedSpill(cellIndex);
    const sheetId = this.workbook.cellStore.sheetIds[cellIndex]!;
    const sheetName = this.workbook.getSheetNameById(sheetId);
    const address = this.workbook.getAddress(cellIndex);
    const owner = parseCellAddress(address, sheetName);

    if (owner.row + arrayValue.rows > MAX_ROWS || owner.col + arrayValue.cols > MAX_COLS) {
      return { changedCellIndices, ownerValue: errorValue(ErrorCode.Spill) };
    }

    for (let rowOffset = 0; rowOffset < arrayValue.rows; rowOffset += 1) {
      for (let colOffset = 0; colOffset < arrayValue.cols; colOffset += 1) {
        if (rowOffset === 0 && colOffset === 0) {
          continue;
        }
        const targetAddress = formatAddress(owner.row + rowOffset, owner.col + colOffset);
        const targetIndex = this.workbook.getCellIndex(sheetName, targetAddress);
        if (targetIndex === undefined) {
          continue;
        }
        const targetValue = this.workbook.cellStore.getValue(targetIndex, (id) =>
          this.strings.get(id),
        );
        if (this.formulas.get(targetIndex) || targetValue.tag !== ValueTag.Empty) {
          return { changedCellIndices, ownerValue: errorValue(ErrorCode.Blocked) };
        }
      }
    }

    for (let rowOffset = 0; rowOffset < arrayValue.rows; rowOffset += 1) {
      for (let colOffset = 0; colOffset < arrayValue.cols; colOffset += 1) {
        if (rowOffset === 0 && colOffset === 0) {
          continue;
        }
        const targetIndex = this.ensureCellTrackedByCoords(
          sheetId,
          owner.row + rowOffset,
          owner.col + colOffset,
        );
        const valueIndex = rowOffset * arrayValue.cols + colOffset;
        const value = arrayValue.values[valueIndex] ?? emptyValue();
        if (this.setSpillChildValue(targetIndex, value)) {
          changedCellIndices.push(targetIndex);
        }
      }
    }

    if (arrayValue.rows > 1 || arrayValue.cols > 1) {
      changedCellIndices.push(
        ...this.applyDerivedOp({
          kind: "upsertSpillRange",
          sheetName,
          address,
          rows: arrayValue.rows,
          cols: arrayValue.cols,
        }),
      );
    }

    return {
      changedCellIndices,
      ownerValue: arrayValue.values[0] ?? emptyValue(),
    };
  }

  private clearSpillChildCell(cellIndex: number): boolean {
    const currentFlags = this.workbook.cellStore.flags[cellIndex] ?? 0;
    const currentValue = this.workbook.cellStore.getValue(cellIndex, (id) => this.strings.get(id));
    if (currentValue.tag === ValueTag.Empty && (currentFlags & CellFlags.SpillChild) === 0) {
      return false;
    }
    this.workbook.cellStore.setValue(cellIndex, emptyValue());
    this.workbook.cellStore.flags[cellIndex] = currentFlags & ~CellFlags.SpillChild;
    return true;
  }

  private setSpillChildValue(cellIndex: number, value: CellValue): boolean {
    const currentValue = this.workbook.cellStore.getValue(cellIndex, (id) => this.strings.get(id));
    const currentFlags = this.workbook.cellStore.flags[cellIndex] ?? 0;
    const nextFlags =
      (currentFlags & ~(CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle)) |
      CellFlags.SpillChild;
    if (areCellValuesEqual(currentValue, value) && currentFlags === nextFlags) {
      return false;
    }
    this.workbook.cellStore.setValue(
      cellIndex,
      value,
      value.tag === ValueTag.String ? this.strings.intern(value.value) : 0,
    );
    this.workbook.cellStore.flags[cellIndex] = nextFlags;
    return true;
  }

  private forEachSheetCell(
    sheetId: number,
    fn: (cellIndex: number, row: number, col: number) => void,
  ): void {
    const sheet = this.workbook.getSheetById(sheetId);
    if (!sheet) {
      return;
    }
    sheet.grid.forEachCellEntry((cellIndex, row, col) => {
      fn(cellIndex, row, col);
    });
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
    this.reverseDefinedNameEdges.clear();
    this.reverseTableEdges.clear();
    this.reverseSpillEdges.clear();
    this.pivotOutputOwners.clear();
    this.ranges.reset();
    this.edgeArena.reset();
    this.entityVersions.clear();
    this.sheetDeleteVersions.clear();
    this.undoStack.length = 0;
    this.redoStack.length = 0;
    this.selection = {
      sheetName: "Sheet1",
      address: "A1",
      anchorAddress: "A1",
      range: { startAddress: "A1", endAddress: "A1" },
      editMode: "idle",
    };
    this.syncState = "local-only";
    this.lastMetrics = {
      batchId: previousBatchId,
      changedInputCount: 0,
      dirtyFormulaCount: 0,
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
      rangeNodeVisits: 0,
      recalcMs: 0,
      compileMs: 0,
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

function intersectRangeBounds(
  range: CellRangeRef,
  bounds: { startRow: number; endRow: number; startCol: number; endCol: number },
): { startRow: number; endRow: number; startCol: number; endCol: number } | undefined {
  const normalized = normalizeRange(range);
  const startRow = Math.max(bounds.startRow, normalized.startRow);
  const endRow = Math.min(bounds.endRow, normalized.endRow);
  const startCol = Math.max(bounds.startCol, normalized.startCol);
  const endCol = Math.min(bounds.endCol, normalized.endCol);
  if (startRow > endRow || startCol > endCol) {
    return undefined;
  }
  return { startRow, endRow, startCol, endCol };
}

function normalizeRange(range: CellRangeRef): {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
} {
  const start = parseCellAddress(range.startAddress, range.sheetName);
  const end = parseCellAddress(range.endAddress, range.sheetName);
  return {
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endCol: Math.max(start.col, end.col),
  };
}

export const selectors = {
  selectCellSnapshot,
  selectMetrics,
  selectSelectionState,
  selectViewportCells,
};
