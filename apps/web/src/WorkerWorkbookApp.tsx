import { useCallback, useEffect, useMemo, useRef, useState, useSyncExternalStore } from "react";
import { useActorRef, useSelector } from "@xstate/react";
import type { CommitOp } from "@bilig/core";
import { WorkbookView, type EditMovement, type EditSelectionBehavior } from "@bilig/grid";
import { formatAddress, parseCellAddress } from "@bilig/formula";
import {
  type CellNumberFormatInput,
  type CellRangeRef,
  MAX_COLS,
  MAX_ROWS,
  ValueTag,
  formatErrorCode,
  parseCellNumberFormatCode,
  type CellSnapshot,
  type CellStyleField,
  type CellStylePatch,
  type LiteralInput,
} from "@bilig/protocol";
import { mutators, type BiligRuntimeConfig } from "@bilig/zero-sync";
import { createWorkerRuntimeMachine } from "./runtime-machine.js";
import { resolveRuntimeConfig } from "./runtime-config.js";
import type { ZeroClient } from "./runtime-session.js";
import { loadPersistedSelection, persistSelection } from "./selection-persistence.js";
import { WorkerViewportCache } from "./viewport-cache.js";
import { WorkbookToolbar, type BorderPreset } from "./workbook-toolbar.js";
import { isPresetColor, mergeRecentCustomColors, normalizeHexColor } from "./workbook-colors.js";
import { cn } from "./cn.js";

type EditingMode = "idle" | "cell" | "formula";

type ParsedEditorInput =
  | { kind: "clear" }
  | { kind: "formula"; formula: string }
  | { kind: "value"; value: LiteralInput };

type ZeroConnectionState =
  | { name: "connected" }
  | { name: "connecting"; reason?: string }
  | { name: "disconnected"; reason: string }
  | {
      name: "needs-auth";
      reason:
        | { type: "mutate"; status: 401 | 403; body?: string }
        | { type: "query"; status: 401 | 403; body?: string }
        | { type: "zero-cache"; reason: string };
    }
  | { name: "error"; reason: string }
  | { name: "closed"; reason: string };

const BORDER_CLEAR_FIELDS: readonly CellStyleField[] = [
  "borderTop",
  "borderRight",
  "borderBottom",
  "borderLeft",
] as const;

const DEFAULT_BORDER_SIDE = {
  style: "solid",
  weight: "thin",
  color: "#111827",
} as const;
function createNextSheetName(sheetNames: readonly string[]): string {
  const existing = new Set(sheetNames);
  let index = 1;
  while (existing.has(`Sheet${index}`)) {
    index += 1;
  }
  return `Sheet${index}`;
}

function normalizeSheetNameKey(value: string): string {
  return value.trim().toUpperCase();
}
function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function assert(condition: unknown, message: string): asserts condition {
  if (!condition) {
    throw new Error(message);
  }
}

function isLiteralInput(value: unknown): value is LiteralInput {
  return (
    value === null ||
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean"
  );
}

function isCellRangeRef(value: unknown): value is CellRangeRef {
  return (
    isRecord(value) &&
    typeof value["sheetName"] === "string" &&
    typeof value["startAddress"] === "string" &&
    typeof value["endAddress"] === "string"
  );
}

function isCellStyleFieldList(value: unknown): value is CellStyleField[] {
  return Array.isArray(value) && value.every((entry) => typeof entry === "string");
}

function isCellStylePatchValue(value: unknown): value is CellStylePatch {
  return isRecord(value);
}

function isCellNumberFormatInputValue(value: unknown): value is CellNumberFormatInput {
  return typeof value === "string" || isRecord(value);
}

function isCommitOps(value: unknown): value is CommitOp[] {
  return Array.isArray(value);
}

function toResolvedValue(cell: CellSnapshot): string {
  switch (cell.value.tag) {
    case ValueTag.Number:
      return String(cell.value.value);
    case ValueTag.Boolean:
      return cell.value.value ? "TRUE" : "FALSE";
    case ValueTag.String:
      return cell.value.value;
    case ValueTag.Error:
      return formatErrorCode(cell.value.code);
    case ValueTag.Empty:
      return "";
  }
  const exhaustiveValue: never = cell.value;
  return String(exhaustiveValue);
}

function toEditorValue(cell: CellSnapshot): string {
  if (cell.value.tag === ValueTag.Error) {
    return formatErrorCode(cell.value.code);
  }
  if (cell.formula) {
    return `=${cell.formula}`;
  }
  if (cell.input === null || cell.input === undefined) {
    return toResolvedValue(cell);
  }
  if (typeof cell.input === "boolean") {
    return cell.input ? "TRUE" : "FALSE";
  }
  return String(cell.input);
}

function parseEditorInput(rawValue: string): ParsedEditorInput {
  const normalized = rawValue.trim();
  if (normalized.startsWith("=")) {
    return { kind: "formula", formula: normalized.slice(1) };
  }
  if (normalized === "") {
    return { kind: "clear" };
  }
  if (normalized === "TRUE" || normalized === "FALSE") {
    return { kind: "value", value: normalized === "TRUE" };
  }
  const numeric = Number(normalized);
  if (!Number.isNaN(numeric) && /^-?\d+(\.\d+)?$/.test(normalized)) {
    return { kind: "value", value: numeric };
  }
  return { kind: "value", value: normalized };
}

function clampSelectionMovement(
  address: string,
  sheetName: string,
  movement: EditMovement,
): string {
  const parsed = parseCellAddress(address, sheetName);
  const nextRow = Math.min(MAX_ROWS - 1, Math.max(0, parsed.row + movement[1]));
  const nextCol = Math.min(MAX_COLS - 1, Math.max(0, parsed.col + movement[0]));
  return formatAddress(nextRow, nextCol);
}

function parseSelectionTarget(
  input: string,
  fallbackSheet: string,
): { sheetName: string; address: string } | null {
  const trimmed = input.trim();
  if (trimmed.length === 0) {
    return null;
  }

  const bangIndex = trimmed.lastIndexOf("!");
  const nextSheetName = bangIndex === -1 ? fallbackSheet : trimmed.slice(0, bangIndex);
  const nextAddress = bangIndex === -1 ? trimmed : trimmed.slice(bangIndex + 1);

  try {
    const parsed = parseCellAddress(nextAddress.toUpperCase(), nextSheetName || fallbackSheet);
    return {
      sheetName: nextSheetName || fallbackSheet,
      address: formatAddress(parsed.row, parsed.col),
    };
  } catch {
    return null;
  }
}

function parseSelectionRangeLabel(
  label: string,
  sheetName: string,
): { sheetName: string; startAddress: string; endAddress: string } {
  const trimmed = label.trim().toUpperCase();
  if (trimmed === "ALL") {
    return {
      sheetName,
      startAddress: "A1",
      endAddress: formatAddress(MAX_ROWS - 1, MAX_COLS - 1),
    };
  }

  const rowSelection = /^(\d+):(\d+)$/.exec(trimmed);
  if (rowSelection) {
    const startRow = Math.min(Number(rowSelection[1]) - 1, Number(rowSelection[2]) - 1);
    const endRow = Math.max(Number(rowSelection[1]) - 1, Number(rowSelection[2]) - 1);
    return {
      sheetName,
      startAddress: formatAddress(startRow, 0),
      endAddress: formatAddress(endRow, MAX_COLS - 1),
    };
  }

  const columnSelection = /^([A-Z]+):([A-Z]+)$/.exec(trimmed);
  if (columnSelection) {
    const startColumn = parseCellAddress(`${columnSelection[1]}1`, sheetName).col;
    const endColumn = parseCellAddress(`${columnSelection[2]}1`, sheetName).col;
    return {
      sheetName,
      startAddress: formatAddress(0, Math.min(startColumn, endColumn)),
      endAddress: formatAddress(MAX_ROWS - 1, Math.max(startColumn, endColumn)),
    };
  }

  const [startAddress = label, endAddress = startAddress] = trimmed.includes(":")
    ? trimmed.split(":")
    : [trimmed, trimmed];
  return { sheetName, startAddress, endAddress };
}

function getNormalizedRangeBounds(range: CellRangeRef): {
  sheetName: string;
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
} {
  const start = parseCellAddress(range.startAddress, range.sheetName);
  const end = parseCellAddress(range.endAddress, range.sheetName);
  return {
    sheetName: range.sheetName,
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endCol: Math.max(start.col, end.col),
  };
}

function cloneCellSnapshotForAddress(
  snapshot: CellSnapshot,
  sheetName: string,
  address: string,
): CellSnapshot {
  return {
    ...snapshot,
    sheetName,
    address,
    version: snapshot.version + 1,
  };
}

function applyOptimisticMoveRange(
  cache: WorkerViewportCache,
  source: CellRangeRef,
  target: CellRangeRef,
): void {
  const sourceBounds = getNormalizedRangeBounds(source);
  const targetBounds = getNormalizedRangeBounds(target);
  const height = sourceBounds.endRow - sourceBounds.startRow + 1;
  const width = sourceBounds.endCol - sourceBounds.startCol + 1;
  const sourceMatrix: CellSnapshot[][] = [];

  for (let rowOffset = 0; rowOffset < height; rowOffset += 1) {
    const row: CellSnapshot[] = [];
    for (let colOffset = 0; colOffset < width; colOffset += 1) {
      row.push(
        cache.getCell(
          source.sheetName,
          formatAddress(sourceBounds.startRow + rowOffset, sourceBounds.startCol + colOffset),
        ),
      );
    }
    sourceMatrix.push(row);
  }

  for (let rowOffset = 0; rowOffset < height; rowOffset += 1) {
    for (let colOffset = 0; colOffset < width; colOffset += 1) {
      const sourceAddress = formatAddress(
        sourceBounds.startRow + rowOffset,
        sourceBounds.startCol + colOffset,
      );
      cache.setCellSnapshot(
        cloneCellSnapshotForAddress(
          emptyCellSnapshot(source.sheetName, sourceAddress),
          source.sheetName,
          sourceAddress,
        ),
      );
    }
  }

  for (let rowOffset = 0; rowOffset < height; rowOffset += 1) {
    for (let colOffset = 0; colOffset < width; colOffset += 1) {
      const sourceSnapshot = sourceMatrix[rowOffset]![colOffset]!;
      const targetAddress = formatAddress(
        targetBounds.startRow + rowOffset,
        targetBounds.startCol + colOffset,
      );
      cache.setCellSnapshot(
        cloneCellSnapshotForAddress(
          { ...sourceSnapshot, address: targetAddress, sheetName: target.sheetName },
          target.sheetName,
          targetAddress,
        ),
      );
    }
  }
}

function createRangeRef(
  sheetName: string,
  startRow: number,
  startCol: number,
  endRow: number,
  endCol: number,
): CellRangeRef {
  return {
    sheetName,
    startAddress: formatAddress(startRow, startCol),
    endAddress: formatAddress(endRow, endCol),
  };
}

function formatConnectionStateLabel(state: ZeroConnectionState["name"]): string {
  switch (state) {
    case "connected":
      return "Live";
    case "connecting":
      return "Connecting";
    case "disconnected":
      return "Disconnected";
    case "needs-auth":
      return "Needs auth";
    case "error":
      return "Error";
    case "closed":
      return "Closed";
    default:
      return state;
  }
}

function toErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }
  return JSON.stringify(error);
}

function isMutationErrorResult(value: unknown): value is { type: "error"; error: unknown } {
  return (
    typeof value === "object" &&
    value !== null &&
    "type" in value &&
    value.type === "error" &&
    "error" in value
  );
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

function isTextEntryTarget(target: EventTarget | null): boolean {
  return (
    target instanceof HTMLInputElement ||
    target instanceof HTMLTextAreaElement ||
    target instanceof HTMLSelectElement ||
    (target instanceof HTMLElement && target.isContentEditable)
  );
}

const workerRuntimeMachine = createWorkerRuntimeMachine();

export function WorkerWorkbookApp(props: {
  config: BiligRuntimeConfig;
  connectionState: ZeroConnectionState;
  zero: ZeroClient;
}) {
  const runtimeConfig = useMemo(() => resolveRuntimeConfig(props.config), [props.config]);
  const runtimeKey = [
    runtimeConfig.documentId,
    runtimeConfig.persistState ? "persist" : "memory",
  ].join("|");

  return (
    <WorkerWorkbookAppInner
      key={runtimeKey}
      runtimeConfig={runtimeConfig}
      connectionState={props.connectionState}
      zero={props.zero}
    />
  );
}

function WorkerWorkbookAppInner({
  runtimeConfig,
  connectionState,
  zero,
}: {
  runtimeConfig: ReturnType<typeof resolveRuntimeConfig>;
  connectionState: ZeroConnectionState;
  zero: ZeroClient;
}) {
  const documentId = runtimeConfig.documentId;
  const replicaId = useMemo(() => `browser:${Math.random().toString(36).slice(2)}`, []);
  const initialSelection = useMemo(() => loadPersistedSelection(documentId), [documentId]);
  const runtimeActorRef = useActorRef(workerRuntimeMachine, {
    input: {
      documentId,
      replicaId,
      persistState: runtimeConfig.persistState,
      zero,
      initialSelection,
    },
  });
  const runtimeController = useSelector(runtimeActorRef, (snapshot) => snapshot.context.controller);
  const workerHandle = useSelector(runtimeActorRef, (snapshot) => snapshot.context.handle);
  const bridgeState = useSelector(runtimeActorRef, (snapshot) => snapshot.context.bridgeState);
  const selection = useSelector(runtimeActorRef, (snapshot) => snapshot.context.selection);
  const runtimeSelectedCell = useSelector(
    runtimeActorRef,
    (snapshot) => snapshot.context.selectedCell,
  );
  const runtimeError = useSelector(runtimeActorRef, (snapshot) => snapshot.context.error);
  const loading = useSelector(runtimeActorRef, (snapshot) =>
    snapshot.matches({ active: "booting" }),
  );
  const runtimeReady = !loading && Boolean(workerHandle);
  const workbookReady = runtimeReady;
  const selectedCell =
    runtimeSelectedCell ?? emptyCellSnapshot(selection.sheetName, selection.address);
  const [selectionLabel, setSelectionLabel] = useState("A1");
  const [recentFillColors, setRecentFillColors] = useState<readonly string[]>([]);
  const [recentTextColors, setRecentTextColors] = useState<readonly string[]>([]);
  const [zeroHealthReady, setZeroHealthReady] = useState(false);
  const [editorValue, setEditorValue] = useState("");
  const [editorSelectionBehavior, setEditorSelectionBehavior] =
    useState<EditSelectionBehavior>("select-all");
  const [editingMode, setEditingMode] = useState<EditingMode>("idle");
  const selectionRef = useRef(selection);
  const workerHandleRef = useRef(workerHandle);
  const editorValueRef = useRef(editorValue);
  const editingModeRef = useRef(editingMode);
  const editorTargetRef = useRef(selection);
  const zeroRef = useRef<ZeroClient>(zero);

  useEffect(() => {
    selectionRef.current = selection;
  }, [selection]);

  useEffect(() => {
    persistSelection(documentId, selection);
  }, [documentId, selection]);

  useEffect(() => {
    workerHandleRef.current = workerHandle;
  }, [workerHandle]);

  useEffect(() => {
    editorValueRef.current = editorValue;
  }, [editorValue]);

  useEffect(() => {
    editingModeRef.current = editingMode;
  }, [editingMode]);

  useEffect(() => {
    zeroRef.current = zero;
  }, [zero]);

  useEffect(() => {
    if (!runtimeConfig.persistState || !runtimeReady) {
      setZeroHealthReady(false);
      return;
    }
    if (connectionState.name === "error" || connectionState.name === "closed") {
      setZeroHealthReady(false);
      return;
    }

    let cancelled = false;
    const probe = async (): Promise<void> => {
      try {
        const response = await fetch("/zero/keepalive", { cache: "no-store" });
        if (response.ok) {
          if (!cancelled) {
            setZeroHealthReady(true);
          }
          return;
        }
      } catch {}
      if (!cancelled) {
        window.setTimeout(() => {
          void probe();
        }, 250);
      }
    };

    setZeroHealthReady(false);
    void probe();
    return () => {
      cancelled = true;
    };
  }, [connectionState.name, runtimeConfig.persistState, runtimeReady]);

  const writesAllowed =
    connectionState.name === "connected" || connectionState.name === "connecting";

  const columnWidths = useSyncExternalStore(
    useCallback(
      (listener: () => void) => workerHandle?.cache.subscribe(listener) ?? (() => {}),
      [workerHandle],
    ),
    () => workerHandle?.cache.getColumnWidths(selection.sheetName),
    () => workerHandle?.cache.getColumnWidths(selection.sheetName),
  );

  const reportRuntimeError = useCallback(
    (error: unknown) => {
      runtimeActorRef.send({
        type: "session.error",
        message: error instanceof Error ? error.message : String(error),
      });
    },
    [runtimeActorRef],
  );

  const runZeroMutation = useCallback(
    (mutation: Parameters<ZeroClient["mutate"]>[0]) => {
      const result = zeroRef.current.mutate(mutation);
      const observerResult =
        (result as { server?: Promise<unknown> }).server ??
        (result as { client?: Promise<unknown> }).client;
      if (!observerResult) {
        return;
      }
      void observerResult
        .then((serverResult) => {
          if (isMutationErrorResult(serverResult)) {
            const error = serverResult.error;
            reportRuntimeError(error instanceof Error ? error : new Error(toErrorMessage(error)));
          }
          return undefined;
        })
        .catch(reportRuntimeError);
    },
    [reportRuntimeError],
  );

  const invokeMutation = useCallback(
    (method: string, ...args: unknown[]): Promise<void> => {
      if (!writesAllowed) {
        return Promise.reject(
          new Error(`Writes are unavailable while Zero is ${connectionState.name}`),
        );
      }

      switch (method) {
        case "setCellValue": {
          const [sheetName, address, value] = args;
          assert(
            typeof sheetName === "string" && typeof address === "string" && isLiteralInput(value),
            "Invalid setCellValue args",
          );
          runZeroMutation(
            mutators.workbook.setCellValue({ documentId, sheetName, address, value }),
          );
          return Promise.resolve();
        }
        case "setCellFormula": {
          const [sheetName, address, formula] = args;
          assert(
            typeof sheetName === "string" &&
              typeof address === "string" &&
              typeof formula === "string",
            "Invalid setCellFormula args",
          );
          runZeroMutation(
            mutators.workbook.setCellFormula({
              documentId,
              sheetName,
              address,
              formula,
            }),
          );
          return Promise.resolve();
        }
        case "clearCell": {
          const [sheetName, address] = args;
          assert(
            typeof sheetName === "string" && typeof address === "string",
            "Invalid clearCell args",
          );
          runZeroMutation(mutators.workbook.clearCell({ documentId, sheetName, address }));
          return Promise.resolve();
        }
        case "clearRange": {
          const [range] = args;
          assert(isCellRangeRef(range), "Invalid clearRange args");
          runZeroMutation(mutators.workbook.clearRange({ documentId, range }));
          return Promise.resolve();
        }
        case "renderCommit": {
          const [ops] = args;
          assert(isCommitOps(ops), "Invalid renderCommit args");
          runZeroMutation(mutators.workbook.renderCommit({ documentId, ops }));
          return Promise.resolve();
        }
        case "fillRange": {
          const [source, target] = args;
          assert(isCellRangeRef(source) && isCellRangeRef(target), "Invalid fillRange args");
          runZeroMutation(mutators.workbook.fillRange({ documentId, source, target }));
          return Promise.resolve();
        }
        case "copyRange": {
          const [source, target] = args;
          assert(isCellRangeRef(source) && isCellRangeRef(target), "Invalid copyRange args");
          runZeroMutation(mutators.workbook.copyRange({ documentId, source, target }));
          return Promise.resolve();
        }
        case "moveRange": {
          const [source, target] = args;
          assert(isCellRangeRef(source) && isCellRangeRef(target), "Invalid moveRange args");
          runZeroMutation(mutators.workbook.moveRange({ documentId, source, target }));
          return Promise.resolve();
        }
        case "updateColumnWidth": {
          const [sheetName, columnIndex, width] = args;
          assert(
            typeof sheetName === "string" &&
              typeof columnIndex === "number" &&
              typeof width === "number",
            "Invalid updateColumnWidth args",
          );
          runZeroMutation(
            mutators.workbook.updateColumnWidth({
              documentId,
              sheetName,
              columnIndex,
              width,
            }),
          );
          return Promise.resolve();
        }
        case "setRangeStyle": {
          const [range, patch] = args;
          assert(
            isCellRangeRef(range) && isCellStylePatchValue(patch),
            "Invalid setRangeStyle args",
          );
          runZeroMutation(mutators.workbook.setRangeStyle({ documentId, range, patch }));
          return Promise.resolve();
        }
        case "clearRangeStyle": {
          const [range, fields] = args;
          assert(
            isCellRangeRef(range) && (fields === undefined || isCellStyleFieldList(fields)),
            "Invalid clearRangeStyle args",
          );
          runZeroMutation(mutators.workbook.clearRangeStyle({ documentId, range, fields }));
          return Promise.resolve();
        }
        case "setRangeNumberFormat": {
          const [range, format] = args;
          assert(
            isCellRangeRef(range) && isCellNumberFormatInputValue(format),
            "Invalid setRangeNumberFormat args",
          );
          runZeroMutation(
            mutators.workbook.setRangeNumberFormat({
              documentId,
              range,
              format,
            }),
          );
          return Promise.resolve();
        }
        case "clearRangeNumberFormat": {
          const [range] = args;
          assert(isCellRangeRef(range), "Invalid clearRangeNumberFormat args");
          runZeroMutation(mutators.workbook.clearRangeNumberFormat({ documentId, range }));
          return Promise.resolve();
        }
        default:
          return Promise.reject(new Error(`Unsupported workbook mutation: ${method}`));
      }
    },
    [connectionState.name, documentId, runZeroMutation, writesAllowed],
  );

  const getLiveSelectedCell = useCallback(
    (nextSelection = selectionRef.current) => {
      const active = workerHandleRef.current;
      if (!active) {
        return selectedCell;
      }
      return active.cache.getCell(nextSelection.sheetName, nextSelection.address);
    },
    [selectedCell],
  );

  const beginEditing = useCallback(
    (
      seed?: string,
      selectionBehavior: EditSelectionBehavior = "select-all",
      mode: Exclude<EditingMode, "idle"> = "cell",
    ) => {
      if (!writesAllowed) {
        return;
      }
      const nextEditorValue = seed ?? toEditorValue(getLiveSelectedCell());
      editorValueRef.current = nextEditorValue;
      setEditorValue(nextEditorValue);
      setEditorSelectionBehavior(selectionBehavior);
      editorTargetRef.current = selectionRef.current;
      editingModeRef.current = mode;
      setEditingMode(mode);
    },
    [getLiveSelectedCell, writesAllowed],
  );

  const applyParsedInput = useCallback(
    async (sheetName: string, address: string, parsed: ParsedEditorInput) => {
      const cache = workerHandleRef.current?.cache;
      if (cache) {
        const current = cache.peekCell(sheetName, address) ?? emptyCellSnapshot(sheetName, address);
        const next: CellSnapshot = {
          ...current,
          sheetName,
          address,
          version: current.version + 1,
        };

        if (parsed.kind === "formula") {
          next.formula = parsed.formula;
          next.input = parsed.formula;
          next.value = { tag: ValueTag.Empty };
        } else if (parsed.kind === "clear") {
          delete next.formula;
          delete next.input;
          next.value = { tag: ValueTag.Empty };
        } else if (parsed.value === null) {
          delete next.formula;
          next.input = null;
          next.value = { tag: ValueTag.Empty };
        } else if (typeof parsed.value === "number") {
          delete next.formula;
          next.input = parsed.value;
          next.value = { tag: ValueTag.Number, value: parsed.value };
        } else if (typeof parsed.value === "boolean") {
          delete next.formula;
          next.input = parsed.value;
          next.value = { tag: ValueTag.Boolean, value: parsed.value };
        } else {
          delete next.formula;
          next.input = parsed.value;
          next.value = { tag: ValueTag.String, value: parsed.value, stringId: 0 };
        }

        cache.setCellSnapshot(next);
      }

      if (parsed.kind === "formula") {
        await invokeMutation("setCellFormula", sheetName, address, parsed.formula);
        return;
      }
      if (parsed.kind === "clear") {
        await invokeMutation("clearCell", sheetName, address);
        return;
      }
      await invokeMutation("setCellValue", sheetName, address, parsed.value);
    },
    [invokeMutation],
  );

  const commitEditor = useCallback(
    (movement?: EditMovement) => {
      if (!writesAllowed) {
        return;
      }
      const targetSelection =
        editingModeRef.current === "idle" ? selectionRef.current : editorTargetRef.current;
      const nextValue =
        editingModeRef.current === "idle"
          ? toEditorValue(getLiveSelectedCell(targetSelection))
          : editorValueRef.current;
      const parsed = parseEditorInput(nextValue);
      editorTargetRef.current = targetSelection;
      editingModeRef.current = "idle";
      setEditingMode("idle");
      setEditorSelectionBehavior("select-all");
      if (movement) {
        const nextAddress = clampSelectionMovement(
          targetSelection.address,
          targetSelection.sheetName,
          movement,
        );
        const nextSelection = { sheetName: targetSelection.sheetName, address: nextAddress };
        selectionRef.current = nextSelection;
        runtimeActorRef.send({ type: "selection.changed", selection: nextSelection });
      }
      editorTargetRef.current = selectionRef.current;
      void applyParsedInput(targetSelection.sheetName, targetSelection.address, parsed).catch(
        reportRuntimeError,
      );
    },
    [applyParsedInput, getLiveSelectedCell, reportRuntimeError, runtimeActorRef, writesAllowed],
  );

  const cancelEditor = useCallback(() => {
    const nextEditorValue = toEditorValue(getLiveSelectedCell());
    editorValueRef.current = nextEditorValue;
    setEditorValue(nextEditorValue);
    setEditorSelectionBehavior("select-all");
    editorTargetRef.current = selectionRef.current;
    editingModeRef.current = "idle";
    setEditingMode("idle");
  }, [getLiveSelectedCell]);

  const clearSelectedRange = useCallback(() => {
    if (!writesAllowed) {
      return;
    }
    const targetRange = parseSelectionRangeLabel(selectionLabel, selection.sheetName);
    editorValueRef.current = "";
    setEditorValue("");
    editorTargetRef.current = selectionRef.current;
    editingModeRef.current = "idle";
    setEditingMode("idle");
    void invokeMutation("clearRange", targetRange).catch((error: unknown) => {
      reportRuntimeError(error);
    });
  }, [invokeMutation, reportRuntimeError, selection.sheetName, selectionLabel, writesAllowed]);

  const clearSelectedCell = useCallback(() => {
    if (!writesAllowed) {
      return;
    }
    clearSelectedRange();
  }, [clearSelectedRange, writesAllowed]);

  const toggleBooleanCell = useCallback(
    (sheetName: string, address: string, nextValue: boolean) => {
      if (!writesAllowed) {
        return;
      }
      void applyParsedInput(sheetName, address, { kind: "value", value: nextValue }).catch(
        reportRuntimeError,
      );
    },
    [applyParsedInput, reportRuntimeError, writesAllowed],
  );

  const pasteIntoSelection = useCallback(
    (sheetName: string, startAddr: string, values: readonly (readonly string[])[]) => {
      const start = parseCellAddress(startAddr, sheetName);
      const ops: {
        kind: "upsertCell" | "deleteCell";
        sheetName: string;
        addr: string;
        formula?: string;
        value?: LiteralInput;
      }[] = [];
      values.forEach((rowValues, rowOffset) => {
        rowValues.forEach((cellValue, colOffset) => {
          const address = formatAddress(start.row + rowOffset, start.col + colOffset);
          const parsed = parseEditorInput(cellValue);
          if (parsed.kind === "formula") {
            ops.push({
              kind: "upsertCell",
              sheetName,
              addr: address,
              formula: parsed.formula,
            });
            return;
          }
          if (parsed.kind === "clear") {
            ops.push({ kind: "deleteCell", sheetName, addr: address });
            return;
          }
          ops.push({
            kind: "upsertCell",
            sheetName,
            addr: address,
            value: parsed.value,
          });
        });
      });
      if (ops.length === 0) {
        return;
      }
      void invokeMutation("renderCommit", ops).catch(reportRuntimeError);
      setEditorSelectionBehavior("select-all");
      editorTargetRef.current = selectionRef.current;
      editingModeRef.current = "idle";
      setEditingMode("idle");
    },
    [invokeMutation, reportRuntimeError],
  );

  const fillSelectionRange = useCallback(
    (
      sourceStartAddr: string,
      sourceEndAddr: string,
      targetStartAddr: string,
      targetEndAddr: string,
    ) => {
      const targetSelection = selectionRef.current;
      const source = {
        sheetName: targetSelection.sheetName,
        startAddress: sourceStartAddr,
        endAddress: sourceEndAddr,
      };
      const target = {
        sheetName: targetSelection.sheetName,
        startAddress: targetStartAddr,
        endAddress: targetEndAddr,
      };
      void invokeMutation("fillRange", source, target)
        .then(() => {
          editorTargetRef.current = selectionRef.current;
          editingModeRef.current = "idle";
          setEditingMode("idle");
          return undefined;
        })
        .catch(reportRuntimeError);
    },
    [invokeMutation, reportRuntimeError],
  );

  const copySelectionRange = useCallback(
    (
      sourceStartAddr: string,
      sourceEndAddr: string,
      targetStartAddr: string,
      targetEndAddr: string,
    ) => {
      const targetSelection = selectionRef.current;
      const source = {
        sheetName: targetSelection.sheetName,
        startAddress: sourceStartAddr,
        endAddress: sourceEndAddr,
      };
      const target = {
        sheetName: targetSelection.sheetName,
        startAddress: targetStartAddr,
        endAddress: targetEndAddr,
      };
      void invokeMutation("copyRange", source, target)
        .then(() => {
          editorTargetRef.current = selectionRef.current;
          editingModeRef.current = "idle";
          setEditingMode("idle");
          return undefined;
        })
        .catch(reportRuntimeError);
    },
    [invokeMutation, reportRuntimeError],
  );

  const moveSelectionRange = useCallback(
    (
      sourceStartAddr: string,
      sourceEndAddr: string,
      targetStartAddr: string,
      targetEndAddr: string,
    ) => {
      const targetSelection = selectionRef.current;
      const source = {
        sheetName: targetSelection.sheetName,
        startAddress: sourceStartAddr,
        endAddress: sourceEndAddr,
      };
      const target = {
        sheetName: targetSelection.sheetName,
        startAddress: targetStartAddr,
        endAddress: targetEndAddr,
      };
      const cache = workerHandleRef.current?.cache;
      if (cache) {
        applyOptimisticMoveRange(cache, source, target);
      }
      void invokeMutation("moveRange", source, target)
        .then(() => {
          editorTargetRef.current = selectionRef.current;
          editingModeRef.current = "idle";
          setEditingMode("idle");
          return undefined;
        })
        .catch(reportRuntimeError);
    },
    [invokeMutation, reportRuntimeError],
  );

  const selectAddress = useCallback(
    (sheetName: string, address: string) => {
      if (editingModeRef.current !== "idle") {
        editorTargetRef.current = { sheetName, address };
        editingModeRef.current = "idle";
        setEditingMode("idle");
      }
      const nextSelection = { sheetName, address };
      selectionRef.current = nextSelection;
      editorTargetRef.current = nextSelection;
      runtimeActorRef.send({ type: "selection.changed", selection: nextSelection });
    },
    [runtimeActorRef],
  );

  const handleEditorChange = useCallback((next: string) => {
    editorValueRef.current = next;
    setEditorValue(next);
    setEditingMode((current) => {
      const nextMode = current === "idle" ? "cell" : current;
      editingModeRef.current = nextMode;
      return nextMode;
    });
  }, []);

  const isEditing = editingMode !== "idle";
  const isEditingCell = editingMode === "cell";
  const visibleEditorValue = isEditing ? editorValue : toEditorValue(selectedCell);
  const resolvedValue = toResolvedValue(selectedCell);
  const sheetNames = useMemo(
    () => [...(bridgeState?.sheetNames ?? [selection.sheetName])],
    [bridgeState?.sheetNames, selection.sheetName],
  );
  const selectedStyle = workerHandle?.cache.getCellStyle(selectedCell.styleId);
  const selectionRange = parseSelectionRangeLabel(selectionLabel, selection.sheetName);
  const currentNumberFormat = parseCellNumberFormatCode(selectedCell.format);
  const selectedFontSize = String(selectedStyle?.font?.size ?? 11);
  const isBoldActive = selectedStyle?.font?.bold === true;
  const isItalicActive = selectedStyle?.font?.italic === true;
  const isUnderlineActive = selectedStyle?.font?.underline === true;
  const horizontalAlignment = selectedStyle?.alignment?.horizontal ?? null;
  const isWrapActive = selectedStyle?.alignment?.wrap === true;
  const currentFillColor = normalizeHexColor(selectedStyle?.fill?.backgroundColor ?? "#ffffff");
  const currentTextColor = normalizeHexColor(selectedStyle?.font?.color ?? "#111827");
  const visibleRecentFillColors = useMemo(
    () =>
      isPresetColor(currentFillColor)
        ? recentFillColors
        : mergeRecentCustomColors(recentFillColors, currentFillColor),
    [currentFillColor, recentFillColors],
  );
  const visibleRecentTextColors = useMemo(
    () =>
      isPresetColor(currentTextColor)
        ? recentTextColors
        : mergeRecentCustomColors(recentTextColors, currentTextColor),
    [currentTextColor, recentTextColors],
  );
  const statusModeLabel = formatConnectionStateLabel(connectionState.name);
  const statusModeValue = connectionState.name === "connected" ? "Live" : statusModeLabel;

  const subscribeViewport = useCallback(
    (
      sheetName: string,
      viewport: Parameters<WorkerViewportCache["subscribeViewport"]>[1],
      listener: Parameters<WorkerViewportCache["subscribeViewport"]>[2],
    ) => {
      if (!runtimeController) {
        return () => {};
      }
      return runtimeController.subscribeViewport(sheetName, viewport, listener);
    },
    [runtimeController],
  );

  const statusSyncValue = !runtimeReady
    ? "Loading"
    : !runtimeConfig.persistState || zeroHealthReady
      ? "Ready"
      : connectionState.name === "connecting"
        ? "Syncing"
        : "Unavailable";
  const statusChipClass =
    "inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 text-[12px] font-medium text-[var(--wb-text-muted)] shadow-[var(--wb-shadow-sm)]";

  const selectionStatus = useMemo(
    () => (
      <span className={statusChipClass} data-testid="status-selection">
        {selection.sheetName}!{selectionLabel}
      </span>
    ),
    [selection.sheetName, selectionLabel, statusChipClass],
  );

  const headerStatus = useMemo(
    () => (
      <>
        <span
          aria-label={`Workbook status: ${statusModeValue}, ${statusSyncValue}`}
          className="inline-flex h-8 w-8 items-center justify-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] shadow-[var(--wb-shadow-sm)]"
          data-testid="status-mode"
          role="status"
          title={statusSyncValue}
        >
          <span
            aria-hidden="true"
            className={cn(
              "block h-2.5 w-2.5 rounded-full",
              statusSyncValue === "Ready"
                ? "bg-[#1f7a43]"
                : statusSyncValue === "Syncing" || statusSyncValue === "Loading"
                  ? "bg-[#b26a00]"
                  : "bg-[#b42318]",
            )}
          />
          <span className="sr-only">{statusModeValue}</span>
        </span>
        <span className="sr-only" data-testid="status-sync">
          {statusSyncValue}
        </span>
      </>
    ),
    [statusModeValue, statusSyncValue],
  );

  const applyRangeStyle = useCallback(
    async (patch: CellStylePatch) => {
      workerHandleRef.current?.cache.applyOptimisticRangeStyle(selectionRange, patch);
      await invokeMutation("setRangeStyle", selectionRange, patch);
    },
    [invokeMutation, selectionRange],
  );

  const clearRangeStyleFields = useCallback(
    async (fields?: CellStyleField[]) => {
      workerHandleRef.current?.cache.clearOptimisticRangeStyle(selectionRange, fields);
      await invokeMutation("clearRangeStyle", selectionRange, fields);
    },
    [invokeMutation, selectionRange],
  );

  const applyFillColor = useCallback(
    async (color: string, source: "preset" | "custom") => {
      const normalized = normalizeHexColor(color);
      await applyRangeStyle({ fill: { backgroundColor: normalized } });
      if (source === "custom") {
        setRecentFillColors((current) => mergeRecentCustomColors(current, normalized));
      }
    },
    [applyRangeStyle],
  );

  const resetFillColor = useCallback(async () => {
    await applyRangeStyle({ fill: { backgroundColor: null } });
  }, [applyRangeStyle]);

  const applyTextColor = useCallback(
    async (color: string, source: "preset" | "custom") => {
      const normalized = normalizeHexColor(color);
      await applyRangeStyle({ font: { color: normalized } });
      if (source === "custom") {
        setRecentTextColors((current) => mergeRecentCustomColors(current, normalized));
      }
    },
    [applyRangeStyle],
  );

  const resetTextColor = useCallback(async () => {
    await applyRangeStyle({ font: { color: null } });
  }, [applyRangeStyle]);

  const applyBorderPreset = useCallback(
    async (preset: BorderPreset) => {
      const { sheetName, startRow, endRow, startCol, endCol } =
        getNormalizedRangeBounds(selectionRange);
      const applyBorders = async (
        range: CellRangeRef,
        borders: NonNullable<CellStylePatch["borders"]>,
      ) => {
        await invokeMutation("setRangeStyle", range, { borders });
      };
      const applyRowBorder = async (rowStart: number, rowEnd: number, side: "top" | "bottom") => {
        if (rowStart > rowEnd) {
          return;
        }
        await applyBorders(createRangeRef(sheetName, rowStart, startCol, rowEnd, endCol), {
          [side]: DEFAULT_BORDER_SIDE,
        });
      };
      const applyColumnBorder = async (
        colStart: number,
        colEnd: number,
        side: "left" | "right",
      ) => {
        if (colStart > colEnd) {
          return;
        }
        await applyBorders(createRangeRef(sheetName, startRow, colStart, endRow, colEnd), {
          [side]: DEFAULT_BORDER_SIDE,
        });
      };

      await invokeMutation("clearRangeStyle", selectionRange, [...BORDER_CLEAR_FIELDS]);

      switch (preset) {
        case "clear":
          return;
        case "all":
          await applyRowBorder(startRow, endRow, "top");
          await applyColumnBorder(startCol, endCol, "left");
          await applyRowBorder(endRow, endRow, "bottom");
          await applyColumnBorder(endCol, endCol, "right");
          return;
        case "outer":
          await applyRowBorder(startRow, startRow, "top");
          await applyRowBorder(endRow, endRow, "bottom");
          await applyColumnBorder(startCol, startCol, "left");
          await applyColumnBorder(endCol, endCol, "right");
          return;
        case "left":
          await applyColumnBorder(startCol, startCol, "left");
          return;
        case "top":
          await applyRowBorder(startRow, startRow, "top");
          return;
        case "right":
          await applyColumnBorder(endCol, endCol, "right");
          return;
        case "bottom":
          await applyRowBorder(endRow, endRow, "bottom");
          return;
        default: {
          const exhaustive: never = preset;
          return exhaustive;
        }
      }
    },
    [invokeMutation, selectionRange],
  );

  const createSheet = useCallback(() => {
    const nextSheetName = createNextSheetName(sheetNames);
    workerHandle?.cache.setKnownSheets([...sheetNames, nextSheetName]);
    void invokeMutation("renderCommit", [
      {
        kind: "upsertSheet",
        name: nextSheetName,
        order: sheetNames.length,
      } satisfies CommitOp,
    ])
      .then(() => selectAddress(nextSheetName, "A1"))
      .catch(reportRuntimeError);
  }, [invokeMutation, reportRuntimeError, selectAddress, sheetNames, workerHandle]);

  const renameSheet = useCallback(
    (currentName: string, nextName: string) => {
      const trimmedName = nextName.trim();
      if (trimmedName.length === 0 || trimmedName === currentName) {
        return;
      }
      const currentKey = normalizeSheetNameKey(currentName);
      const nextKey = normalizeSheetNameKey(trimmedName);
      if (
        sheetNames.some(
          (name) =>
            normalizeSheetNameKey(name) === nextKey && normalizeSheetNameKey(name) !== currentKey,
        )
      ) {
        return;
      }

      workerHandle?.cache.setKnownSheets(
        sheetNames.map((name) => (name === currentName ? trimmedName : name)),
      );
      void invokeMutation("renderCommit", [
        {
          kind: "renameSheet",
          oldName: currentName,
          newName: trimmedName,
        } satisfies CommitOp,
      ])
        .then(() => {
          if (selectionRef.current.sheetName === currentName) {
            selectAddress(trimmedName, selectionRef.current.address);
          }
          return undefined;
        })
        .catch(reportRuntimeError);
    },
    [invokeMutation, reportRuntimeError, selectAddress, sheetNames, workerHandle],
  );

  const setNumberFormatPreset = useCallback(
    async (preset: string) => {
      switch (preset) {
        case "general":
          await invokeMutation("clearRangeNumberFormat", selectionRange);
          return;
        case "number":
          await invokeMutation("setRangeNumberFormat", selectionRange, {
            kind: "number",
            decimals: 2,
            useGrouping: true,
          });
          return;
        case "currency":
          await invokeMutation("setRangeNumberFormat", selectionRange, {
            kind: "currency",
            currency: "USD",
            decimals: 2,
            useGrouping: true,
            negativeStyle: "minus",
            zeroStyle: "zero",
          });
          return;
        case "accounting":
          await invokeMutation("setRangeNumberFormat", selectionRange, {
            kind: "accounting",
            currency: "USD",
            decimals: 2,
            useGrouping: true,
            negativeStyle: "parentheses",
            zeroStyle: "dash",
          });
          return;
        case "percent":
          await invokeMutation("setRangeNumberFormat", selectionRange, {
            kind: "percent",
            decimals: 2,
          });
          return;
        case "date":
          await invokeMutation("setRangeNumberFormat", selectionRange, {
            kind: "date",
            dateStyle: "short",
          });
          return;
        case "text":
          await invokeMutation("setRangeNumberFormat", selectionRange, "text");
          return;
      }
    },
    [invokeMutation, selectionRange],
  );

  useEffect(() => {
    const handleWindowShortcut = (event: KeyboardEvent) => {
      if (event.defaultPrevented || isTextEntryTarget(event.target) || event.altKey) {
        return;
      }

      const hasPrimaryModifier = event.metaKey || event.ctrlKey;
      if (!hasPrimaryModifier) {
        return;
      }

      const normalizedKey = event.key.toLowerCase();
      if (normalizedKey === "s") {
        event.preventDefault();
        return;
      }

      if (!writesAllowed) {
        return;
      }

      if (!event.shiftKey && normalizedKey === "b") {
        event.preventDefault();
        void applyRangeStyle({ font: { bold: !isBoldActive } });
        return;
      }

      if (!event.shiftKey && normalizedKey === "i") {
        event.preventDefault();
        void applyRangeStyle({ font: { italic: !isItalicActive } });
        return;
      }

      if (!event.shiftKey && normalizedKey === "u") {
        event.preventDefault();
        void applyRangeStyle({ font: { underline: !isUnderlineActive } });
        return;
      }

      if (event.shiftKey && event.code === "Digit1") {
        event.preventDefault();
        void setNumberFormatPreset("number");
        return;
      }

      if (event.shiftKey && event.code === "Digit4") {
        event.preventDefault();
        void setNumberFormatPreset("currency");
        return;
      }

      if (event.shiftKey && event.code === "Digit5") {
        event.preventDefault();
        void setNumberFormatPreset("percent");
        return;
      }

      if (event.shiftKey && event.code === "Digit7") {
        event.preventDefault();
        void applyBorderPreset("outer");
        return;
      }

      if (event.shiftKey && normalizedKey === "l") {
        event.preventDefault();
        void applyRangeStyle({ alignment: { horizontal: "left" } });
        return;
      }

      if (event.shiftKey && normalizedKey === "e") {
        event.preventDefault();
        void applyRangeStyle({ alignment: { horizontal: "center" } });
        return;
      }

      if (event.shiftKey && normalizedKey === "r") {
        event.preventDefault();
        void applyRangeStyle({ alignment: { horizontal: "right" } });
        return;
      }

      if (!event.shiftKey && event.code === "Backslash") {
        event.preventDefault();
        void clearRangeStyleFields();
      }
    };

    window.addEventListener("keydown", handleWindowShortcut, true);
    return () => {
      window.removeEventListener("keydown", handleWindowShortcut, true);
    };
  }, [
    applyBorderPreset,
    applyRangeStyle,
    clearRangeStyleFields,
    isBoldActive,
    isItalicActive,
    isUnderlineActive,
    setNumberFormatPreset,
    writesAllowed,
  ]);

  const ribbon = useMemo(
    () => (
      <WorkbookToolbar
        currentFillColor={currentFillColor}
        currentNumberFormatKind={currentNumberFormat.kind}
        currentTextColor={currentTextColor}
        horizontalAlignment={horizontalAlignment}
        isBoldActive={isBoldActive}
        isItalicActive={isItalicActive}
        isUnderlineActive={isUnderlineActive}
        isWrapActive={isWrapActive}
        onApplyBorderPreset={applyBorderPreset}
        onClearStyle={() => {
          void clearRangeStyleFields();
        }}
        onFillColorReset={() => {
          void resetFillColor();
        }}
        onFillColorSelect={(color, source) => {
          void applyFillColor(color, source);
        }}
        onFontSizeChange={(value) => {
          void applyRangeStyle({ font: { size: value ? Number(value) : null } });
        }}
        onHorizontalAlignmentChange={(alignment) => {
          void applyRangeStyle({
            alignment: {
              horizontal: horizontalAlignment === alignment ? null : alignment,
            },
          });
        }}
        onNumberFormatChange={(value) => {
          void setNumberFormatPreset(value);
        }}
        onTextColorReset={() => {
          void resetTextColor();
        }}
        onTextColorSelect={(color, source) => {
          void applyTextColor(color, source);
        }}
        onToggleBold={() => {
          void applyRangeStyle({ font: { bold: !isBoldActive } });
        }}
        onToggleItalic={() => {
          void applyRangeStyle({ font: { italic: !isItalicActive } });
        }}
        onToggleUnderline={() => {
          void applyRangeStyle({ font: { underline: !isUnderlineActive } });
        }}
        onToggleWrap={() => {
          void applyRangeStyle({
            alignment: { wrap: !isWrapActive },
          });
        }}
        recentFillColors={visibleRecentFillColors}
        recentTextColors={visibleRecentTextColors}
        selectedFontSize={selectedFontSize}
        writesAllowed={writesAllowed}
      />
    ),
    [
      applyBorderPreset,
      applyFillColor,
      applyRangeStyle,
      applyTextColor,
      clearRangeStyleFields,
      currentFillColor,
      currentNumberFormat.kind,
      currentTextColor,
      horizontalAlignment,
      isBoldActive,
      isItalicActive,
      isUnderlineActive,
      isWrapActive,
      resetFillColor,
      resetTextColor,
      selectedFontSize,
      setNumberFormatPreset,
      visibleRecentFillColors,
      visibleRecentTextColors,
      writesAllowed,
    ],
  );

  return (
    <div className="flex h-screen flex-col overflow-hidden bg-[var(--wb-app-bg)] text-[var(--wb-text)]">
      {runtimeError ? (
        <div
          className="border-b border-[#f1b5b5] bg-[#fff7f7] px-3 py-2 text-sm text-[#991b1b]"
          data-testid="worker-error"
        >
          {runtimeError}
        </div>
      ) : null}
      {runtimeReady && !writesAllowed ? (
        <div className="border-b border-[var(--wb-accent-ring)] bg-[var(--wb-accent-soft)] px-3 py-2 text-sm text-[var(--wb-accent)]">
          Zero is {statusModeLabel.toLowerCase()}. Editing is disabled until the connection
          recovers.
        </div>
      ) : null}
      <div className="relative flex min-h-0 flex-1">
        <div className="min-h-0 min-w-0 flex-1">
          {workbookReady && workerHandle ? (
            <WorkbookView
              ribbon={ribbon}
              editorValue={visibleEditorValue}
              editorSelectionBehavior={editorSelectionBehavior}
              engine={workerHandle.cache}
              isEditing={Boolean(writesAllowed && isEditing)}
              isEditingCell={Boolean(writesAllowed && isEditingCell)}
              onAddressCommit={(input) => {
                const nextTarget = parseSelectionTarget(input, selection.sheetName);
                if (nextTarget) {
                  selectAddress(nextTarget.sheetName, nextTarget.address);
                }
              }}
              onAutofitColumn={(columnIndex: number, fallbackWidth: number) => {
                workerHandle?.cache.setColumnWidth(selection.sheetName, columnIndex, fallbackWidth);
                return invokeMutation(
                  "updateColumnWidth",
                  selection.sheetName,
                  columnIndex,
                  fallbackWidth,
                )
                  .then(() => undefined)
                  .catch(reportRuntimeError);
              }}
              onBeginEdit={beginEditing}
              onBeginFormulaEdit={(seed?: string) => beginEditing(seed, "select-all", "formula")}
              onCancelEdit={cancelEditor}
              onClearCell={clearSelectedCell}
              onColumnWidthChange={(columnIndex: number, newSize: number) => {
                workerHandle?.cache.setColumnWidth(selection.sheetName, columnIndex, newSize);
                void invokeMutation(
                  "updateColumnWidth",
                  selection.sheetName,
                  columnIndex,
                  newSize,
                ).catch(reportRuntimeError);
              }}
              onCommitEdit={commitEditor}
              onCopyRange={copySelectionRange}
              onCreateSheet={writesAllowed ? createSheet : undefined}
              onEditorChange={handleEditorChange}
              onFillRange={fillSelectionRange}
              onMoveRange={moveSelectionRange}
              onPaste={pasteIntoSelection}
              onToggleBooleanCell={toggleBooleanCell}
              onRenameSheet={writesAllowed ? renameSheet : undefined}
              onSelectionLabelChange={setSelectionLabel}
              onSelect={(addr) => selectAddress(selection.sheetName, addr)}
              onSelectSheet={(sheetName) => selectAddress(sheetName, "A1")}
              resolvedValue={resolvedValue}
              selectedAddr={selection.address}
              selectedCellSnapshot={selectedCell}
              selectionStatus={selectionStatus}
              sheetName={selection.sheetName}
              sheetNames={sheetNames}
              headerStatus={headerStatus}
              subscribeViewport={subscribeViewport}
              columnWidths={columnWidths}
            />
          ) : null}
        </div>
      </div>
    </div>
  );
}
