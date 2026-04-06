import { formatAddress } from "@bilig/formula";
import { queries } from "@bilig/zero-sync";
import {
  ValueTag,
  formatCellDisplayValue,
  type CellSnapshot,
  type CellStyleRecord,
  type CellValue,
  type LiteralInput,
  type RecalcMetrics,
  type Viewport,
} from "@bilig/protocol";
import type { ViewportPatch } from "@bilig/worker-transport";
import type { WorkerViewportCache } from "./viewport-cache.js";

interface ZeroWorkbookSyncSourceLike {
  materialize(query: unknown): unknown;
}

interface LiveView<T> {
  readonly data: T;
  addListener(listener: (value: T) => void): () => void;
  destroy(): void;
}

interface CellSourceRow {
  sheetName: string;
  address: string;
  inputValue?: LiteralInput | null;
  formula?: string;
  styleId?: string | null;
  explicitFormatId?: string | null;
}

interface CellEvalRow {
  sheetName: string;
  address: string;
  value: CellValue;
  flags: number;
  version: number;
  styleId?: string | null;
  formatId?: string | null;
  formatCode?: string | null;
}

interface AxisMetadataRow {
  readonly startIndex: number;
  readonly size?: number;
  readonly hidden?: boolean;
}

interface ViewportProjectionState {
  nextVersion: number;
  lastCellSignatures: Map<string, string>;
  lastColumnSignatures: Map<number, string>;
  lastRowSignatures: Map<number, string>;
}

interface WorkbookRevisionState {
  headRevision: number;
  calculatedRevision: number;
}

interface ZeroViewportSubscription {
  readonly sheetName: string;
  readonly viewport: Viewport;
  readonly listener: (damage?: readonly { cell: readonly [number, number] }[]) => void;
  readonly state: ViewportProjectionState;
  readonly sourceView: LiveView<readonly unknown[]>;
  readonly evalView: LiveView<readonly unknown[]>;
  readonly rowView: LiveView<readonly unknown[]>;
  readonly columnView: LiveView<readonly unknown[]>;
  destroy(): void;
}

const DEFAULT_STYLE_ID = "style-0";
const EMPTY_METRICS: RecalcMetrics = {
  batchId: 0,
  changedInputCount: 0,
  dirtyFormulaCount: 0,
  wasmFormulaCount: 0,
  jsFormulaCount: 0,
  rangeNodeVisits: 0,
  recalcMs: 0,
  compileMs: 0,
};
const PRODUCT_COLUMN_WIDTH = 104;
const PRODUCT_ROW_HEIGHT = 22;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isLiveView<T>(value: unknown): value is LiveView<T> {
  return (
    isRecord(value) &&
    "data" in value &&
    typeof value["addListener"] === "function" &&
    typeof value["destroy"] === "function"
  );
}

function toErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function assertLiveView<T>(value: unknown, label: string): LiveView<T> {
  if (!isLiveView<T>(value)) {
    throw new Error(`Zero workbook sync source returned an invalid ${label} view`);
  }
  return value;
}

function emptyCellSnapshot(sheetName: string, address: string): CellSnapshot {
  return {
    sheetName,
    address,
    value: { tag: ValueTag.Empty },
    flags: 0,
    version: 0,
  };
}

function isLiteralInput(value: unknown): value is LiteralInput {
  return (
    value === null ||
    typeof value === "boolean" ||
    typeof value === "number" ||
    typeof value === "string"
  );
}

function isValueTag(value: unknown): value is ValueTag {
  return (
    value === ValueTag.Number ||
    value === ValueTag.Boolean ||
    value === ValueTag.String ||
    value === ValueTag.Error ||
    value === ValueTag.Empty
  );
}

function normalizeCellSourceRow(value: unknown): CellSourceRow | null {
  if (
    !isRecord(value) ||
    typeof value["sheetName"] !== "string" ||
    typeof value["address"] !== "string"
  ) {
    return null;
  }
  const row: CellSourceRow = {
    sheetName: value["sheetName"],
    address: value["address"],
  };
  if (isLiteralInput(value["inputValue"])) {
    row.inputValue = value["inputValue"];
  }
  if (typeof value["formula"] === "string") {
    row.formula = value["formula"];
  }
  if (typeof value["styleId"] === "string" || value["styleId"] === null) {
    row.styleId = value["styleId"];
  }
  if (typeof value["explicitFormatId"] === "string" || value["explicitFormatId"] === null) {
    row.explicitFormatId = value["explicitFormatId"];
  }
  return row;
}

function normalizeCellSourceRows(value: unknown): Map<string, CellSourceRow> {
  const rows = new Map<string, CellSourceRow>();
  if (!Array.isArray(value)) {
    return rows;
  }
  for (const entry of value) {
    const normalized = normalizeCellSourceRow(entry);
    if (!normalized) {
      continue;
    }
    rows.set(normalized.address, normalized);
  }
  return rows;
}

function normalizeCellValue(value: unknown): CellValue | null {
  if (!isRecord(value) || !isValueTag(value["tag"])) {
    return null;
  }
  switch (value["tag"]) {
    case ValueTag.Number:
      return typeof value["value"] === "number"
        ? { tag: ValueTag.Number, value: value["value"] }
        : null;
    case ValueTag.Boolean:
      return typeof value["value"] === "boolean"
        ? { tag: ValueTag.Boolean, value: value["value"] }
        : null;
    case ValueTag.String:
      return typeof value["value"] === "string"
        ? {
            tag: ValueTag.String,
            value: value["value"],
            stringId: typeof value["stringId"] === "number" ? value["stringId"] : 0,
          }
        : null;
    case ValueTag.Error:
      return typeof value["code"] === "number"
        ? { tag: ValueTag.Error, code: value["code"] }
        : null;
    case ValueTag.Empty:
      return { tag: ValueTag.Empty };
    default:
      return null;
  }
}

function normalizeCellEvalRow(value: unknown): CellEvalRow | null {
  if (
    !isRecord(value) ||
    typeof value["sheetName"] !== "string" ||
    typeof value["address"] !== "string" ||
    typeof value["flags"] !== "number" ||
    typeof value["version"] !== "number"
  ) {
    return null;
  }
  const normalizedValue = normalizeCellValue(value["value"]);
  if (!normalizedValue) {
    return null;
  }
  const row: CellEvalRow = {
    sheetName: value["sheetName"],
    address: value["address"],
    value: normalizedValue,
    flags: value["flags"],
    version: value["version"],
  };
  if (typeof value["styleId"] === "string" || value["styleId"] === null) {
    row.styleId = value["styleId"];
  }
  if (typeof value["formatId"] === "string" || value["formatId"] === null) {
    row.formatId = value["formatId"];
  }
  if (typeof value["formatCode"] === "string" || value["formatCode"] === null) {
    row.formatCode = value["formatCode"];
  }
  return row;
}

function normalizeCellEvalRows(value: unknown): Map<string, CellEvalRow> {
  const rows = new Map<string, CellEvalRow>();
  if (!Array.isArray(value)) {
    return rows;
  }
  for (const entry of value) {
    const normalized = normalizeCellEvalRow(entry);
    if (!normalized) {
      continue;
    }
    rows.set(normalized.address, normalized);
  }
  return rows;
}

function normalizeWorkbookRevisionState(value: unknown): WorkbookRevisionState | null {
  if (
    !isRecord(value) ||
    typeof value["headRevision"] !== "number" ||
    typeof value["calculatedRevision"] !== "number"
  ) {
    return null;
  }
  return {
    headRevision: value["headRevision"],
    calculatedRevision: value["calculatedRevision"],
  };
}

function normalizeAxisMetadataRows(value: unknown): Map<number, AxisMetadataRow> {
  const rows = new Map<number, AxisMetadataRow>();
  if (!Array.isArray(value)) {
    return rows;
  }
  for (const entry of value) {
    if (!isRecord(entry) || typeof entry["startIndex"] !== "number") {
      continue;
    }
    rows.set(entry["startIndex"], {
      startIndex: entry["startIndex"],
      ...(typeof entry["size"] === "number" ? { size: entry["size"] } : {}),
      ...(typeof entry["hidden"] === "boolean" ? { hidden: entry["hidden"] } : {}),
    });
  }
  return rows;
}

function normalizeStylesById(value: unknown): Map<string, CellStyleRecord> {
  const styles = new Map<string, CellStyleRecord>([[DEFAULT_STYLE_ID, { id: DEFAULT_STYLE_ID }]]);
  if (!Array.isArray(value)) {
    return styles;
  }
  for (const entry of value) {
    if (!isRecord(entry) || typeof entry["styleId"] !== "string" || !isRecord(entry["styleJson"])) {
      continue;
    }
    styles.set(entry["styleId"], {
      id: entry["styleId"],
      ...(entry["styleJson"] as Omit<CellStyleRecord, "id">),
    });
  }
  return styles;
}

function normalizeNumberFormatsById(value: unknown): Map<string, string> {
  const formats = new Map<string, string>();
  if (!Array.isArray(value)) {
    return formats;
  }
  for (const entry of value) {
    if (
      !isRecord(entry) ||
      typeof entry["formatId"] !== "string" ||
      typeof entry["code"] !== "string"
    ) {
      continue;
    }
    formats.set(entry["formatId"], entry["code"]);
  }
  return formats;
}

function resolveFormatCode(
  formatId: string | null | undefined,
  inlineFormatCode: string | null | undefined,
  numberFormatsById: ReadonlyMap<string, string>,
): string | undefined {
  if (typeof inlineFormatCode === "string") {
    return inlineFormatCode;
  }
  if (typeof formatId !== "string") {
    return undefined;
  }
  return numberFormatsById.get(formatId);
}

function literalInputToCellValue(value: LiteralInput | undefined | null): CellValue {
  if (value === undefined || value === null) {
    return { tag: ValueTag.Empty };
  }
  if (typeof value === "number") {
    return { tag: ValueTag.Number, value };
  }
  if (typeof value === "boolean") {
    return { tag: ValueTag.Boolean, value };
  }
  return { tag: ValueTag.String, value, stringId: 0 };
}

function applyLiteralValueFallback(
  snapshot: CellSnapshot,
  sourceRow: CellSourceRow | undefined,
): void {
  if (!sourceRow || sourceRow.formula !== undefined || sourceRow.inputValue === undefined) {
    return;
  }
  if (snapshot.value.tag !== ValueTag.Empty) {
    return;
  }
  snapshot.value = literalInputToCellValue(sourceRow.inputValue);
}

function applySourceCellOverrides(
  snapshot: CellSnapshot,
  sourceRow: CellSourceRow | undefined,
  numberFormatsById: ReadonlyMap<string, string>,
): void {
  if (!sourceRow) {
    return;
  }
  if (sourceRow.inputValue !== undefined) {
    snapshot.input = sourceRow.inputValue;
  }
  if (sourceRow.formula !== undefined) {
    snapshot.formula = sourceRow.formula;
  }
  if (sourceRow.styleId !== undefined) {
    if (sourceRow.styleId === null) {
      delete snapshot.styleId;
    } else {
      snapshot.styleId = sourceRow.styleId;
    }
  }
  if (sourceRow.explicitFormatId !== undefined) {
    if (sourceRow.explicitFormatId === null) {
      delete snapshot.numberFormatId;
      delete snapshot.format;
    } else {
      snapshot.numberFormatId = sourceRow.explicitFormatId;
      const format = resolveFormatCode(sourceRow.explicitFormatId, undefined, numberFormatsById);
      if (format !== undefined) {
        snapshot.format = format;
      }
    }
  }
}

function buildCellSnapshot(
  sheetName: string,
  address: string,
  evalRow: CellEvalRow | undefined,
  sourceRow: CellSourceRow | undefined,
  numberFormatsById: ReadonlyMap<string, string>,
): CellSnapshot {
  const formatCode = resolveFormatCode(evalRow?.formatId, evalRow?.formatCode, numberFormatsById);
  const snapshot: CellSnapshot = evalRow
    ? {
        sheetName,
        address,
        value: evalRow.value,
        flags: evalRow.flags,
        version: evalRow.version,
        ...(typeof evalRow.styleId === "string" ? { styleId: evalRow.styleId } : {}),
        ...(typeof evalRow.formatId === "string" ? { numberFormatId: evalRow.formatId } : {}),
        ...(typeof formatCode === "string" ? { format: formatCode } : {}),
      }
    : emptyCellSnapshot(sheetName, address);
  applySourceCellOverrides(snapshot, sourceRow, numberFormatsById);
  applyLiteralValueFallback(snapshot, sourceRow);
  return snapshot;
}

function snapshotValueSignature(snapshot: CellSnapshot): string {
  switch (snapshot.value.tag) {
    case ValueTag.Number:
      return `n:${snapshot.value.value}`;
    case ValueTag.Boolean:
      return `b:${snapshot.value.value ? 1 : 0}`;
    case ValueTag.String:
      return `s:${snapshot.value.stringId}:${snapshot.value.value}`;
    case ValueTag.Error:
      return `e:${snapshot.value.code}`;
    case ValueTag.Empty:
      return "empty";
  }
}

function toEditorText(snapshot: CellSnapshot): string {
  if (snapshot.formula) {
    return `=${snapshot.formula}`;
  }
  if (snapshot.input === null || snapshot.input === undefined) {
    return formatCellDisplayValue(snapshot.value, snapshot.format);
  }
  if (typeof snapshot.input === "boolean") {
    return snapshot.input ? "TRUE" : "FALSE";
  }
  return String(snapshot.input);
}

function buildPatchedCellSignature(
  snapshot: CellSnapshot,
  displayText: string,
  copyText: string,
  editorText: string,
): string {
  return [
    snapshot.version,
    snapshot.flags,
    snapshot.formula ?? "",
    snapshot.input ?? "",
    snapshot.format ?? "",
    snapshot.styleId ?? "",
    snapshot.numberFormatId ?? "",
    snapshotValueSignature(snapshot),
    displayText,
    copyText,
    editorText,
  ].join("|");
}

function createViewportProjectionState(): ViewportProjectionState {
  return {
    nextVersion: 1,
    lastCellSignatures: new Map(),
    lastColumnSignatures: new Map(),
    lastRowSignatures: new Map(),
  };
}

function buildAxisPatches(
  start: number,
  end: number,
  entries: ReadonlyMap<number, AxisMetadataRow>,
  defaultSize: number,
  previous: Map<number, string>,
  full: boolean,
): { patches: ViewportPatch["columns"]; signatures: Map<number, string> } {
  const signatures = full ? new Map<number, string>() : new Map(previous);
  const patches: ViewportPatch["columns"] = [];
  for (let index = start; index <= end; index += 1) {
    const entry = entries.get(index);
    const size = entry?.size ?? defaultSize;
    const hidden = entry?.hidden ?? false;
    const signature = `${size}:${hidden ? 1 : 0}`;
    signatures.set(index, signature);
    if (full || previous.get(index) !== signature) {
      patches.push({ index, size, hidden });
    }
  }
  return { patches, signatures };
}

function shouldDeferFormulaProjection(
  existingSnapshot: CellSnapshot | undefined,
  sourceRow: CellSourceRow | undefined,
  workbookRevisionState: WorkbookRevisionState | null,
): boolean {
  if (
    workbookRevisionState === null ||
    workbookRevisionState.calculatedRevision >= workbookRevisionState.headRevision
  ) {
    return false;
  }
  return sourceRow?.formula !== undefined || existingSnapshot?.formula !== undefined;
}

function projectViewportPatch(input: {
  state: ViewportProjectionState;
  sheetName: string;
  viewport: Viewport;
  full: boolean;
  sourceRows: ReadonlyMap<string, CellSourceRow>;
  evalRows: ReadonlyMap<string, CellEvalRow>;
  rowMetadata: ReadonlyMap<number, AxisMetadataRow>;
  columnMetadata: ReadonlyMap<number, AxisMetadataRow>;
  stylesById: ReadonlyMap<string, CellStyleRecord>;
  numberFormatsById: ReadonlyMap<string, string>;
  workbookRevisionState: WorkbookRevisionState | null;
  existingSnapshotAtAddress: (address: string) => CellSnapshot | undefined;
}): ViewportPatch {
  const { state, sheetName, viewport, full, sourceRows, evalRows, rowMetadata, columnMetadata } =
    input;
  const styles = new Map<string, CellStyleRecord>();
  const cells: ViewportPatch["cells"] = [];

  if (full) {
    state.lastCellSignatures.clear();
  }

  for (let row = viewport.rowStart; row <= viewport.rowEnd; row += 1) {
    for (let col = viewport.colStart; col <= viewport.colEnd; col += 1) {
      const address = formatAddress(row, col);
      const sourceRow = sourceRows.get(address);
      if (
        shouldDeferFormulaProjection(
          input.existingSnapshotAtAddress(address),
          sourceRow,
          input.workbookRevisionState,
        )
      ) {
        continue;
      }
      const snapshot = buildCellSnapshot(
        sheetName,
        address,
        evalRows.get(address),
        sourceRow,
        input.numberFormatsById,
      );
      const displayText = formatCellDisplayValue(snapshot.value, snapshot.format);
      const editorText = toEditorText(snapshot);
      const copyText = snapshot.formula ? editorText : displayText;
      const signature = buildPatchedCellSignature(snapshot, displayText, copyText, editorText);
      const key = `${sheetName}!${address}`;
      if (!full && state.lastCellSignatures.get(key) === signature) {
        continue;
      }
      state.lastCellSignatures.set(key, signature);
      const styleId = snapshot.styleId ?? DEFAULT_STYLE_ID;
      styles.set(styleId, input.stylesById.get(styleId) ?? { id: DEFAULT_STYLE_ID });
      cells.push({
        row,
        col,
        snapshot,
        displayText,
        copyText,
        editorText,
        formatId: 0,
        styleId,
      });
    }
  }

  const columnResult = buildAxisPatches(
    viewport.colStart,
    viewport.colEnd,
    columnMetadata,
    PRODUCT_COLUMN_WIDTH,
    state.lastColumnSignatures,
    full,
  );
  state.lastColumnSignatures = columnResult.signatures;

  const rowResult = buildAxisPatches(
    viewport.rowStart,
    viewport.rowEnd,
    rowMetadata,
    PRODUCT_ROW_HEIGHT,
    state.lastRowSignatures,
    full,
  );
  state.lastRowSignatures = rowResult.signatures;

  return {
    version: state.nextVersion++,
    full,
    viewport: { sheetName, ...viewport },
    metrics: EMPTY_METRICS,
    styles: [...styles.values()],
    cells,
    columns: columnResult.patches,
    rows: rowResult.patches,
  };
}

export class ZeroWorkbookLiveSync {
  private readonly workbookView: LiveView<unknown>;
  private readonly stylesView: LiveView<readonly unknown[]>;
  private readonly numberFormatsView: LiveView<readonly unknown[]>;
  private readonly viewportSubscriptions = new Set<ZeroViewportSubscription>();
  private readonly cleanup = new Set<() => void>();
  private stylesById = new Map<string, CellStyleRecord>([
    [DEFAULT_STYLE_ID, { id: DEFAULT_STYLE_ID }],
  ]);
  private numberFormatsById = new Map<string, string>();
  private workbookRevisionState: WorkbookRevisionState | null = null;
  private disposed = false;

  constructor(
    private readonly input: {
      zero: ZeroWorkbookSyncSourceLike;
      documentId: string;
      cache: WorkerViewportCache;
      onError: (message: string) => void;
    },
  ) {
    this.workbookView = assertLiveView<unknown>(
      input.zero.materialize(queries.workbook.get({ documentId: input.documentId })),
      "workbook",
    );
    this.stylesView = assertLiveView<readonly unknown[]>(
      input.zero.materialize(queries.cellStyle.byWorkbook({ documentId: input.documentId })),
      "cell style",
    );
    this.numberFormatsView = assertLiveView<readonly unknown[]>(
      input.zero.materialize(queries.numberFormat.byWorkbook({ documentId: input.documentId })),
      "number format",
    );

    const refreshStyles = (value: unknown) => {
      this.stylesById = normalizeStylesById(value);
      this.refreshAll(true);
    };
    const refreshNumberFormats = (value: unknown) => {
      this.numberFormatsById = normalizeNumberFormatsById(value);
      this.refreshAll(true);
    };
    const refreshWorkbookRevisionState = (value: unknown) => {
      this.workbookRevisionState = normalizeWorkbookRevisionState(value);
      this.refreshAll(true);
    };

    refreshWorkbookRevisionState(this.workbookView.data);
    refreshStyles(this.stylesView.data);
    refreshNumberFormats(this.numberFormatsView.data);

    this.cleanup.add(
      this.workbookView.addListener((value) => {
        refreshWorkbookRevisionState(value);
      }),
    );
    this.cleanup.add(
      this.stylesView.addListener((value) => {
        refreshStyles(value);
      }),
    );
    this.cleanup.add(
      this.numberFormatsView.addListener((value) => {
        refreshNumberFormats(value);
      }),
    );
  }

  subscribeViewport(
    sheetName: string,
    viewport: Viewport,
    listener: (damage?: readonly { cell: readonly [number, number] }[]) => void,
  ): () => void {
    const sourceView = assertLiveView<readonly unknown[]>(
      this.input.zero.materialize(
        queries.cellInput.tile({
          documentId: this.input.documentId,
          sheetName,
          rowStart: viewport.rowStart,
          rowEnd: viewport.rowEnd,
          colStart: viewport.colStart,
          colEnd: viewport.colEnd,
        }),
      ),
      "viewport cell source",
    );
    const evalView = assertLiveView<readonly unknown[]>(
      this.input.zero.materialize(
        queries.cellEval.tile({
          documentId: this.input.documentId,
          sheetName,
          rowStart: viewport.rowStart,
          rowEnd: viewport.rowEnd,
          colStart: viewport.colStart,
          colEnd: viewport.colEnd,
        }),
      ),
      "viewport cell eval",
    );
    const rowView = assertLiveView<readonly unknown[]>(
      this.input.zero.materialize(
        queries.sheetRow.tile({
          documentId: this.input.documentId,
          sheetName,
          rowStart: viewport.rowStart,
          rowEnd: viewport.rowEnd,
        }),
      ),
      "viewport row metadata",
    );
    const columnView = assertLiveView<readonly unknown[]>(
      this.input.zero.materialize(
        queries.sheetCol.tile({
          documentId: this.input.documentId,
          sheetName,
          colStart: viewport.colStart,
          colEnd: viewport.colEnd,
        }),
      ),
      "viewport column metadata",
    );

    const subscription: ZeroViewportSubscription = {
      sheetName,
      viewport,
      listener,
      state: createViewportProjectionState(),
      sourceView,
      evalView,
      rowView,
      columnView,
      destroy: () => {
        sourceView.destroy();
        evalView.destroy();
        rowView.destroy();
        columnView.destroy();
      },
    };

    const publish = (full: boolean) => {
      try {
        if (this.disposed || !this.viewportSubscriptions.has(subscription)) {
          return;
        }
        const patch = projectViewportPatch({
          state: subscription.state,
          sheetName,
          viewport,
          full,
          sourceRows: normalizeCellSourceRows(subscription.sourceView.data),
          evalRows: normalizeCellEvalRows(subscription.evalView.data),
          rowMetadata: normalizeAxisMetadataRows(subscription.rowView.data),
          columnMetadata: normalizeAxisMetadataRows(subscription.columnView.data),
          stylesById: this.stylesById,
          numberFormatsById: this.numberFormatsById,
          workbookRevisionState: this.workbookRevisionState,
          existingSnapshotAtAddress: (address) => this.input.cache.peekCell(sheetName, address),
        });
        if (
          !full &&
          patch.cells.length === 0 &&
          patch.columns.length === 0 &&
          patch.rows.length === 0 &&
          patch.styles.length === 0
        ) {
          return;
        }
        const damage = this.input.cache.applyViewportPatch(patch);
        listener(damage);
      } catch (error) {
        if (!this.disposed) {
          this.input.onError(toErrorMessage(error));
        }
      }
    };

    const unsubscribes = [
      sourceView.addListener(() => publish(false)),
      evalView.addListener(() => publish(false)),
      rowView.addListener(() => publish(false)),
      columnView.addListener(() => publish(false)),
    ];

    const dispose = () => {
      unsubscribes.forEach((unsubscribe) => unsubscribe());
      subscription.destroy();
      this.viewportSubscriptions.delete(subscription);
    };

    this.viewportSubscriptions.add(subscription);
    publish(true);
    return dispose;
  }

  dispose(): void {
    if (this.disposed) {
      return;
    }
    this.disposed = true;
    this.cleanup.forEach((cleanup) => cleanup());
    this.cleanup.clear();
    this.workbookView.destroy();
    this.stylesView.destroy();
    this.numberFormatsView.destroy();
    this.viewportSubscriptions.forEach((subscription) => {
      subscription.destroy();
    });
    this.viewportSubscriptions.clear();
  }

  private refreshAll(full: boolean): void {
    if (this.disposed) {
      return;
    }
    this.viewportSubscriptions.forEach((subscription) => {
      try {
        const patch = projectViewportPatch({
          state: subscription.state,
          sheetName: subscription.sheetName,
          viewport: subscription.viewport,
          full,
          sourceRows: normalizeCellSourceRows(subscription.sourceView.data),
          evalRows: normalizeCellEvalRows(subscription.evalView.data),
          rowMetadata: normalizeAxisMetadataRows(subscription.rowView.data),
          columnMetadata: normalizeAxisMetadataRows(subscription.columnView.data),
          stylesById: this.stylesById,
          numberFormatsById: this.numberFormatsById,
          workbookRevisionState: this.workbookRevisionState,
          existingSnapshotAtAddress: (address) =>
            this.input.cache.peekCell(subscription.sheetName, address),
        });
        if (
          !full &&
          patch.cells.length === 0 &&
          patch.columns.length === 0 &&
          patch.rows.length === 0 &&
          patch.styles.length === 0
        ) {
          return;
        }
        const damage = this.input.cache.applyViewportPatch(patch);
        subscription.listener(damage);
      } catch (error) {
        this.input.onError(toErrorMessage(error));
      }
    });
  }
}
