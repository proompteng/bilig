import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { WorkbookView, type EditMovement, type EditSelectionBehavior } from "@bilig/grid";
import { formatAddress, parseCellAddress } from "@bilig/formula";
import {
  MAX_COLS,
  MAX_ROWS,
  ValueTag,
  formatErrorCode,
  type CellValue,
  type CellSnapshot,
  type LiteralInput
} from "@bilig/protocol";
import { createWorkerEngineClient, type MessagePortLike, type WorkerEngineClient } from "@bilig/worker-transport";
import { WorkerViewportCache } from "./viewport-cache.js";
import type { WorkbookWorkerBootstrapOptions, WorkbookWorkerStateSnapshot } from "./worker-runtime.js";

type EditingMode = "idle" | "cell" | "formula";

type ParsedEditorInput =
  | { kind: "clear" }
  | { kind: "formula"; formula: string }
  | { kind: "value"; value: LiteralInput };

interface WorkerHandle {
  worker: Worker;
  client: WorkerEngineClient;
  cache: WorkerViewportCache;
}

interface RuntimeConfig {
  documentId: string;
  baseUrl: string | null;
  persistState: boolean;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isCellSnapshot(value: unknown): value is CellSnapshot {
  return isRecord(value)
    && typeof value["sheetName"] === "string"
    && typeof value["address"] === "string"
    && typeof value["flags"] === "number"
    && typeof value["version"] === "number"
    && isRecord(value["value"])
    && typeof value["value"]["tag"] === "number";
}

function isRuntimeStateSnapshot(value: unknown): value is WorkbookWorkerStateSnapshot {
  return isRecord(value)
    && typeof value["workbookName"] === "string"
    && Array.isArray(value["sheetNames"])
    && isRecord(value["metrics"])
    && typeof value["syncState"] === "string";
}

function createWorkerPort(worker: Worker): MessagePortLike {
  type PortListener = Parameters<NonNullable<MessagePortLike["addEventListener"]>>[1];
  const listenerMap = new Map<PortListener, EventListener>();
  return {
    postMessage(message: unknown) {
      worker.postMessage(message);
    },
    addEventListener(type: "message", listener: PortListener) {
      const wrapped: EventListener = (event) => {
        if (event instanceof MessageEvent) {
          listener(event);
        }
      };
      listenerMap.set(listener, wrapped);
      worker.addEventListener(type, wrapped);
    },
    removeEventListener(type: "message", listener: PortListener) {
      const wrapped = listenerMap.get(listener);
      if (!wrapped) {
        return;
      }
      listenerMap.delete(listener);
      worker.removeEventListener(type, wrapped);
    }
  };
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

function clampSelectionMovement(address: string, sheetName: string, movement: EditMovement): string {
  const parsed = parseCellAddress(address, sheetName);
  const nextRow = Math.min(MAX_ROWS - 1, Math.max(0, parsed.row + movement[1]));
  const nextCol = Math.min(MAX_COLS - 1, Math.max(0, parsed.col + movement[0]));
  return formatAddress(nextRow, nextCol);
}

function parseSelectionTarget(input: string, fallbackSheet: string): { sheetName: string; address: string } | null {
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
      address: formatAddress(parsed.row, parsed.col)
    };
  } catch {
    return null;
  }
}

function formatSyncStateLabel(state: WorkbookWorkerStateSnapshot["syncState"]): string {
  switch (state) {
    case "live":
      return "Live";
    case "syncing":
      return "Syncing";
    case "local-only":
      return "Local";
    case "behind":
      return "Behind";
    case "reconnecting":
      return "Reconnecting";
  }
  const exhaustiveState: never = state;
  return exhaustiveState;
}

function emptyCellSnapshot(sheetName: string, address: string): CellSnapshot {
  return {
    sheetName,
    address,
    value: { tag: ValueTag.Empty },
    flags: 0,
    version: 0
  };
}

function toOptimisticCellValue(value: LiteralInput, currentValue: CellValue): CellValue {
  if (value === null) {
    return { tag: ValueTag.Empty };
  }
  if (typeof value === "number") {
    return { tag: ValueTag.Number, value };
  }
  if (typeof value === "boolean") {
    return { tag: ValueTag.Boolean, value };
  }
  return {
    tag: ValueTag.String,
    value,
    stringId:
      currentValue.tag === ValueTag.String && currentValue.value === value
        ? currentValue.stringId
        : 0
  };
}

function createSessionDocumentId(defaultDocumentId: string): string {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return `${defaultDocumentId}:${crypto.randomUUID()}`;
  }
  return `${defaultDocumentId}:${Math.random().toString(36).slice(2)}`;
}

function resolveRuntimeConfig(searchParams: URLSearchParams, defaultDocumentId: string, defaultLocalServerUrl: string): RuntimeConfig {
  const explicit = searchParams.get("document");
  if (explicit) {
    return {
      documentId: explicit,
      baseUrl: searchParams.get("server") ?? defaultLocalServerUrl,
      persistState: true
    };
  }
  return {
    documentId: createSessionDocumentId(defaultDocumentId),
    baseUrl: searchParams.get("server") ?? defaultLocalServerUrl,
    persistState: false
  };
}

export function WorkerWorkbookApp() {
  const runtimeConfig = useMemo(() => {
    const defaultDocumentId = import.meta.env["VITE_BILIG_DOCUMENT_ID"] ?? "bilig-demo";
    const defaultLocalServerUrl = import.meta.env["VITE_BILIG_LOCAL_SERVER_URL"] ?? "http://127.0.0.1:4381";
    const searchParams = new URLSearchParams(window.location.search);
    return resolveRuntimeConfig(searchParams, defaultDocumentId, defaultLocalServerUrl);
  }, []);
  const replicaId = useMemo(() => `browser:${Math.random().toString(36).slice(2)}`, []);
  const [workerHandle, setWorkerHandle] = useState<WorkerHandle | null>(null);
  const [runtimeState, setRuntimeState] = useState<WorkbookWorkerStateSnapshot | null>(null);
  const [selection, setSelection] = useState<{ sheetName: string; address: string }>({ sheetName: "Sheet1", address: "A1" });
  const [selectedCell, setSelectedCell] = useState<CellSnapshot>(() => emptyCellSnapshot("Sheet1", "A1"));
  const [selectionLabel, setSelectionLabel] = useState("A1");
  const [editorValue, setEditorValue] = useState("");
  const [editorSelectionBehavior, setEditorSelectionBehavior] = useState<EditSelectionBehavior>("select-all");
  const [editingMode, setEditingMode] = useState<EditingMode>("idle");
  const [runtimeError, setRuntimeError] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);
  const [, setCacheVersion] = useState(0);
  const selectionRef = useRef(selection);
  const workerHandleRef = useRef<WorkerHandle | null>(null);

  useEffect(() => {
    selectionRef.current = selection;
  }, [selection]);

  const refreshRuntimeState = useCallback(async (handle?: WorkerHandle) => {
    const active = handle ?? workerHandleRef.current;
    if (!active) {
      return;
    }
    const response = await active.client.invoke("getRuntimeState");
    if (!isRuntimeStateSnapshot(response)) {
      throw new Error("Worker returned an invalid runtime state payload");
    }
    const nextState = response;
    setRuntimeState(nextState);
  }, []);

  const refreshSelectedCell = useCallback(async (handle?: WorkerHandle, nextSelection?: { sheetName: string; address: string }) => {
    const active = handle ?? workerHandleRef.current;
    const target = nextSelection ?? selectionRef.current;
    if (!active) {
      return;
    }
    const cached = active.cache.peekCell(target.sheetName, target.address);
    if (cached) {
      setSelectedCell(cached);
    }
    const response = await active.client.invoke("getCell", target.sheetName, target.address);
    if (!isCellSnapshot(response)) {
      throw new Error("Worker returned an invalid cell snapshot");
    }
    const snapshot = response;
    if (
      selectionRef.current.sheetName === target.sheetName
      && selectionRef.current.address === target.address
    ) {
      setSelectedCell(snapshot);
    }
  }, []);

  useEffect(() => {
    let disposed = false;
    let unsubscribeEvents: () => void = () => {};
    let unsubscribeCache: () => void = () => {};
    let interval = 0;

    const worker = new Worker(new URL("./workbook.worker.ts", import.meta.url), { type: "module" });
    const client = createWorkerEngineClient({ port: createWorkerPort(worker) });
    const cache = new WorkerViewportCache(client);
    const handle: WorkerHandle = { worker, client, cache };

    setLoading(true);
    setRuntimeError(null);
    void (async () => {
      try {
        const response = await client.invoke("bootstrap", {
          documentId: runtimeConfig.documentId,
          replicaId,
          baseUrl: runtimeConfig.baseUrl,
          persistState: runtimeConfig.persistState
        } satisfies WorkbookWorkerBootstrapOptions);
        if (!isRuntimeStateSnapshot(response)) {
          throw new Error("Worker returned an invalid bootstrap payload");
        }
        const bootstrap = response;
        if (disposed) {
          return;
        }
        workerHandleRef.current = handle;
        setWorkerHandle(handle);
        setRuntimeState(bootstrap);
        const firstSheet = bootstrap.sheetNames[0] ?? "Sheet1";
        setSelection({ sheetName: firstSheet, address: "A1" });
        selectionRef.current = { sheetName: firstSheet, address: "A1" };
        await refreshSelectedCell(handle, selectionRef.current);
        unsubscribeEvents = client.subscribe(() => {
          void refreshRuntimeState(handle).catch((error: unknown) => {
            if (!disposed) {
              setRuntimeError(error instanceof Error ? error.message : String(error));
            }
          });
        });
        unsubscribeCache = cache.subscribe(() => {
          if (disposed) {
            return;
          }
          setCacheVersion((current) => current + 1);
          const next = cache.peekCell(selectionRef.current.sheetName, selectionRef.current.address);
          if (next) {
            setSelectedCell(next);
          }
        });
        interval = window.setInterval(() => {
          void refreshRuntimeState(handle).catch((error: unknown) => {
            if (!disposed) {
              setRuntimeError(error instanceof Error ? error.message : String(error));
            }
          });
        }, 250);
      } catch (error) {
        if (!disposed) {
          setRuntimeError(error instanceof Error ? error.message : String(error));
        }
      } finally {
        if (!disposed) {
          setLoading(false);
        }
      }
    })();

    return () => {
      disposed = true;
      unsubscribeEvents();
      unsubscribeCache();
      if (interval) {
        window.clearInterval(interval);
      }
      client.dispose();
      worker.terminate();
      workerHandleRef.current = null;
    };
  }, [refreshRuntimeState, refreshSelectedCell, replicaId, runtimeConfig.baseUrl, runtimeConfig.documentId, runtimeConfig.persistState]);

  useEffect(() => {
    if (!runtimeState || runtimeState.sheetNames.length === 0) {
      return;
    }
    if (!runtimeState.sheetNames.includes(selection.sheetName)) {
      const nextSelection = { sheetName: runtimeState.sheetNames[0]!, address: "A1" };
      setSelection(nextSelection);
      selectionRef.current = nextSelection;
      void refreshSelectedCell(undefined, nextSelection).catch((error: unknown) => {
        setRuntimeError(error instanceof Error ? error.message : String(error));
      });
    }
  }, [refreshSelectedCell, runtimeState, selection.sheetName]);

  useEffect(() => {
    if (!workerHandle) {
      return;
    }
    let cancelled = false;
    void refreshSelectedCell().catch((error: unknown) => {
      if (!cancelled) {
        setRuntimeError(error instanceof Error ? error.message : String(error));
      }
    });
    return () => {
      cancelled = true;
    };
  }, [refreshSelectedCell, selection.address, selection.sheetName, workerHandle]);

  const invokeMutation = useCallback(async (method: string, ...args: unknown[]): Promise<unknown> => {
    const active = workerHandleRef.current;
    if (!active) {
      throw new Error("Worker runtime is not ready");
    }
    return await active.client.invoke(method, ...args);
  }, []);

  const applyOptimisticCellEdit = useCallback((sheetName: string, address: string, parsed: ParsedEditorInput) => {
    const active = workerHandleRef.current;
    if (!active) {
      return;
    }
    const current = active.cache.getCell(sheetName, address);
    const nextVersion = current.version + 1;
    if (parsed.kind === "clear") {
      active.cache.setCellSnapshot({
        sheetName,
        address,
        value: { tag: ValueTag.Empty },
        flags: current.flags,
        version: nextVersion,
        ...(current.format ? { format: current.format } : {})
      });
      return;
    }
    if (parsed.kind === "formula") {
      active.cache.setCellSnapshot({
        sheetName,
        address,
        formula: parsed.formula,
        value: current.value,
        flags: current.flags,
        version: nextVersion,
        ...(current.format ? { format: current.format } : {})
      });
      return;
    }
    active.cache.setCellSnapshot({
      sheetName,
      address,
      input: parsed.value,
      value: toOptimisticCellValue(parsed.value, current.value),
      flags: current.flags,
      version: nextVersion,
      ...(current.format ? { format: current.format } : {})
    });
  }, []);

  const beginEditing = useCallback((seed?: string, selectionBehavior: EditSelectionBehavior = "select-all", mode: Exclude<EditingMode, "idle"> = "cell") => {
    setEditorValue(seed ?? toEditorValue(selectedCell));
    setEditorSelectionBehavior(selectionBehavior);
    setEditingMode(mode);
  }, [selectedCell]);

  const applyParsedInput = useCallback(async (sheetName: string, address: string, parsed: ParsedEditorInput) => {
    if (parsed.kind === "formula") {
      await invokeMutation("setCellFormula", sheetName, address, parsed.formula);
      return;
    }
    if (parsed.kind === "clear") {
      await invokeMutation("clearCell", sheetName, address);
      return;
    }
    await invokeMutation("setCellValue", sheetName, address, parsed.value);
  }, [invokeMutation]);

  const commitEditor = useCallback((movement?: EditMovement) => {
    const nextValue = editingMode === "idle" ? toEditorValue(selectedCell) : editorValue;
    const parsed = parseEditorInput(nextValue);
    applyOptimisticCellEdit(selection.sheetName, selection.address, parsed);
    setEditingMode("idle");
    setEditorSelectionBehavior("select-all");
    if (movement) {
      const nextAddress = clampSelectionMovement(selection.address, selection.sheetName, movement);
      setSelection({ sheetName: selection.sheetName, address: nextAddress });
      selectionRef.current = { sheetName: selection.sheetName, address: nextAddress };
    }
    void applyParsedInput(selection.sheetName, selection.address, parsed).catch((error: unknown) => {
      setRuntimeError(error instanceof Error ? error.message : String(error));
    });
  }, [applyOptimisticCellEdit, applyParsedInput, editorValue, editingMode, selectedCell, selection.address, selection.sheetName]);

  const cancelEditor = useCallback(() => {
    setEditorValue(toEditorValue(selectedCell));
    setEditorSelectionBehavior("select-all");
    setEditingMode("idle");
  }, [selectedCell]);

  const clearSelectedCell = useCallback(() => {
    applyOptimisticCellEdit(selection.sheetName, selection.address, { kind: "clear" });
    setEditorValue("");
    setEditingMode("idle");
    void invokeMutation("clearCell", selection.sheetName, selection.address).catch((error: unknown) => {
      setRuntimeError(error instanceof Error ? error.message : String(error));
    });
  }, [applyOptimisticCellEdit, invokeMutation, selection.address, selection.sheetName]);

  const pasteIntoSelection = useCallback((startAddr: string, values: readonly (readonly string[])[]) => {
    const start = parseCellAddress(startAddr, selection.sheetName);
    const ops: { kind: "upsertCell" | "deleteCell"; sheetName: string; addr: string; formula?: string; value?: LiteralInput }[] = [];
    values.forEach((rowValues, rowOffset) => {
      rowValues.forEach((cellValue, colOffset) => {
        const address = formatAddress(start.row + rowOffset, start.col + colOffset);
        const parsed = parseEditorInput(cellValue);
        if (parsed.kind === "formula") {
          ops.push({ kind: "upsertCell", sheetName: selection.sheetName, addr: address, formula: parsed.formula });
          return;
        }
        if (parsed.kind === "clear") {
          ops.push({ kind: "deleteCell", sheetName: selection.sheetName, addr: address });
          return;
        }
        ops.push({ kind: "upsertCell", sheetName: selection.sheetName, addr: address, value: parsed.value });
      });
    });
    if (ops.length === 0) {
      return;
    }
    void invokeMutation("renderCommit", ops).catch((error: unknown) => {
      setRuntimeError(error instanceof Error ? error.message : String(error));
    });
    setEditorSelectionBehavior("select-all");
    setEditingMode("idle");
  }, [invokeMutation, selection.sheetName]);

  const fillSelectionRange = useCallback((sourceStartAddr: string, sourceEndAddr: string, targetStartAddr: string, targetEndAddr: string) => {
    void invokeMutation(
      "fillRange",
      {
        sheetName: selection.sheetName,
        startAddress: sourceStartAddr,
        endAddress: sourceEndAddr
      },
      {
        sheetName: selection.sheetName,
        startAddress: targetStartAddr,
        endAddress: targetEndAddr
      }
    ).then(() => {
      setEditingMode("idle");
      return undefined;
    }).catch((error: unknown) => {
      setRuntimeError(error instanceof Error ? error.message : String(error));
    });
  }, [invokeMutation, selection.sheetName]);

  const copySelectionRange = useCallback((sourceStartAddr: string, sourceEndAddr: string, targetStartAddr: string, targetEndAddr: string) => {
    void invokeMutation(
      "copyRange",
      {
        sheetName: selection.sheetName,
        startAddress: sourceStartAddr,
        endAddress: sourceEndAddr
      },
      {
        sheetName: selection.sheetName,
        startAddress: targetStartAddr,
        endAddress: targetEndAddr
      }
    ).then(() => {
      setEditingMode("idle");
      return undefined;
    }).catch((error: unknown) => {
      setRuntimeError(error instanceof Error ? error.message : String(error));
    });
  }, [invokeMutation, selection.sheetName]);

  const selectAddress = useCallback((sheetName: string, address: string) => {
    setSelection({ sheetName, address });
    selectionRef.current = { sheetName, address };
    if (editingMode === "formula") {
      setEditingMode("idle");
    }
  }, [editingMode]);

  const isEditing = editingMode !== "idle";
  const isEditingCell = editingMode === "cell";
  const visibleEditorValue = isEditing ? editorValue : toEditorValue(selectedCell);
  const resolvedValue = toResolvedValue(selectedCell);
  const sheetNames = runtimeState?.sheetNames ?? [];
  const columnWidths = workerHandle ? workerHandle.cache.getColumnWidths(selection.sheetName) : undefined;

  const subscribeViewport = useCallback((sheetName: string, viewport: Parameters<WorkerViewportCache["subscribeViewport"]>[1], listener: Parameters<WorkerViewportCache["subscribeViewport"]>[2]) => {
    if (!workerHandle) {
      return () => {};
    }
    return workerHandle.cache.subscribeViewport(sheetName, viewport, listener);
  }, [workerHandle]);

  const statusBar = (
    <>
      <span data-testid="status-mode">{formatSyncStateLabel(runtimeState?.syncState ?? "local-only")}</span>
      <span data-testid="status-selection">{selection.sheetName}!{selectionLabel}</span>
      <span data-testid="status-sync">{isEditing ? "Editing" : "Ready"}</span>
    </>
  );

  return (
    <div className="app-shell app-shell-product">
      {runtimeError ? (
        <div className="error-banner" data-testid="worker-error">
          {runtimeError}
        </div>
      ) : null}
      {loading || !workerHandle || !runtimeState ? (
        <div className="loading-banner" data-testid="worker-loading">
          Starting worker runtime...
        </div>
      ) : (
        <WorkbookView
          editorValue={visibleEditorValue}
          editorSelectionBehavior={editorSelectionBehavior}
          engine={workerHandle.cache}
          isEditing={isEditing}
          isEditingCell={isEditingCell}
          onAddressCommit={(input) => {
            const nextTarget = parseSelectionTarget(input, selection.sheetName);
            if (nextTarget) {
              selectAddress(nextTarget.sheetName, nextTarget.address);
            }
          }}
          onAutofitColumn={(columnIndex: number, fallbackWidth: number) => {
            workerHandle?.cache.setColumnWidth(selection.sheetName, columnIndex, fallbackWidth);
            return invokeMutation("autofitColumn", selection.sheetName, columnIndex).then((width) => {
              if (typeof width === "number") {
                workerHandle?.cache.setColumnWidth(selection.sheetName, columnIndex, width);
              }
              return undefined;
            });
          }}
          onBeginEdit={beginEditing}
          onBeginFormulaEdit={(seed?: string) => beginEditing(seed, "select-all", "formula")}
          onCancelEdit={cancelEditor}
          onClearCell={clearSelectedCell}
          onColumnWidthChange={(columnIndex: number, newSize: number) => {
            workerHandle?.cache.setColumnWidth(selection.sheetName, columnIndex, newSize);
            void invokeMutation("updateColumnWidth", selection.sheetName, columnIndex, newSize).catch((error: unknown) => {
              setRuntimeError(error instanceof Error ? error.message : String(error));
            });
          }}
          onCommitEdit={commitEditor}
          onCopyRange={copySelectionRange}
          onEditorChange={(next) => {
            setEditorValue(next);
            setEditingMode((current) => current === "idle" ? "cell" : current);
          }}
          onFillRange={fillSelectionRange}
          onPaste={pasteIntoSelection}
          onSelectionLabelChange={setSelectionLabel}
          onSelect={(addr) => selectAddress(selection.sheetName, addr)}
          onSelectSheet={(sheetName) => selectAddress(sheetName, "A1")}
          resolvedValue={resolvedValue}
          selectedAddr={selection.address}
          sheetName={selection.sheetName}
          sheetNames={sheetNames}
          statusBar={statusBar}
          subscribeViewport={subscribeViewport}
          columnWidths={columnWidths}
          variant="product"
          workbookName={runtimeState.workbookName}
        />
      )}
    </div>
  );
}
