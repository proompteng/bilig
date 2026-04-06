import { useCallback, useEffect, useMemo, useRef, useState, useSyncExternalStore } from "react";
import { useActorRef, useSelector } from "@xstate/react";
import type { CommitOp } from "@bilig/core";
import type { EditMovement, EditSelectionBehavior } from "@bilig/grid";
import { formatAddress, parseCellAddress } from "@bilig/formula";
import type { LiteralInput } from "@bilig/protocol";
import { createWorkerRuntimeMachine } from "./runtime-machine.js";
import { resolveRuntimeConfig } from "./runtime-config.js";
import type { ZeroClient } from "./runtime-session.js";
import { loadPersistedSelection, persistSelection } from "./selection-persistence.js";
import { ProjectedViewportStore } from "./projected-viewport-store.js";
import {
  type EditingMode,
  type ParsedEditorInput,
  type ZeroConnectionState,
  canAttemptRemoteSync,
  clampSelectionMovement,
  createNextSheetName,
  emptyCellSnapshot,
  normalizeSheetNameKey,
  parseEditorInput,
  parseSelectionRangeLabel,
  toEditorValue,
  toResolvedValue,
} from "./worker-workbook-app-model.js";
import { useWorkbookSync } from "./use-workbook-sync.js";
import { useWorkbookToolbar } from "./use-workbook-toolbar.js";
import { useWorkbookPresence } from "./use-workbook-presence.js";
import { WorkbookPresenceBar } from "./WorkbookPresenceBar.js";

const workerRuntimeMachine = createWorkerRuntimeMachine();

export function useWorkerWorkbookAppState(input: {
  runtimeConfig: ReturnType<typeof resolveRuntimeConfig>;
  connectionState: ZeroConnectionState;
  zero: ZeroClient;
}) {
  const { runtimeConfig, connectionState, zero } = input;
  const documentId = runtimeConfig.documentId;
  const replicaId = useMemo(() => `browser:${Math.random().toString(36).slice(2)}`, []);
  const initialSelection = useMemo(() => loadPersistedSelection(documentId), [documentId]);
  const runtimeActorRef = useActorRef(workerRuntimeMachine, {
    input: {
      documentId,
      replicaId,
      persistState: runtimeConfig.persistState,
      connectionStateName: connectionState.name,
      zero,
      initialSelection,
    },
  });
  const runtimeController = useSelector(runtimeActorRef, (snapshot) => snapshot.context.controller);
  const workerHandle = useSelector(runtimeActorRef, (snapshot) => snapshot.context.handle);
  const runtimeState = useSelector(runtimeActorRef, (snapshot) => snapshot.context.runtimeState);
  const selection = useSelector(runtimeActorRef, (snapshot) => snapshot.context.selection);
  const runtimeError = useSelector(runtimeActorRef, (snapshot) => snapshot.context.error);
  const runtimeReady = Boolean(workerHandle);
  const workbookReady = runtimeReady;
  const emptySelectedCell = useMemo(
    () => emptyCellSnapshot(selection.sheetName, selection.address),
    [selection.address, selection.sheetName],
  );
  const [selectionLabel, setSelectionLabel] = useState("A1");
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
  const connectionStateRef = useRef(connectionState.name);

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
    connectionStateRef.current = connectionState.name;
  }, [connectionState.name]);

  useEffect(() => {
    runtimeActorRef.send({
      type: "connection.changed",
      connectionStateName: connectionState.name,
    });
  }, [connectionState.name, runtimeActorRef]);

  useEffect(() => {
    if (!runtimeReady) {
      setZeroHealthReady(false);
      return;
    }
    if (
      connectionState.name === "disconnected" ||
      connectionState.name === "needs-auth" ||
      connectionState.name === "error" ||
      connectionState.name === "closed"
    ) {
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
  }, [connectionState.name, runtimeReady]);

  const writesAllowed = runtimeReady;
  const remoteSyncAvailable = canAttemptRemoteSync(connectionState.name);

  const columnWidths = useSyncExternalStore(
    useCallback(
      (listener: () => void) => workerHandle?.viewportStore.subscribe(listener) ?? (() => {}),
      [workerHandle],
    ),
    () => workerHandle?.viewportStore.getColumnWidths(selection.sheetName),
    () => workerHandle?.viewportStore.getColumnWidths(selection.sheetName),
  );

  const selectedCell = useSyncExternalStore(
    useCallback(
      (listener: () => void) => workerHandle?.viewportStore.subscribe(listener) ?? (() => {}),
      [workerHandle],
    ),
    () =>
      workerHandle?.viewportStore.peekCell(selection.sheetName, selection.address) ??
      emptySelectedCell,
    () => emptySelectedCell,
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
  const { invokeMutation, invokeColumnWidthMutation } = useWorkbookSync({
    documentId,
    connectionStateName: connectionState.name,
    connectionStateRef,
    runtimeController,
    workerHandleRef,
    zeroRef,
    reportRuntimeError,
  });

  const getLiveSelectedCell = useCallback(
    (nextSelection = selectionRef.current) => {
      const active = workerHandleRef.current;
      if (!active) {
        return selectedCell;
      }
      return active.viewportStore.getCell(nextSelection.sheetName, nextSelection.address);
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
      if (
        editingModeRef.current === "idle" &&
        selectionRef.current.sheetName === sheetName &&
        selectionRef.current.address === address
      ) {
        return;
      }
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
    () => [...(runtimeState?.sheetNames ?? [selection.sheetName])],
    [runtimeState?.sheetNames, selection.sheetName],
  );
  const collaborators = useWorkbookPresence({
    documentId,
    sessionId: `${documentId}:${replicaId}`,
    selection,
    sheetNames,
    zero,
    enabled: runtimeReady && remoteSyncAvailable,
  });
  const selectedStyle = workerHandle?.viewportStore.getCellStyle(selectedCell.styleId);
  const selectionRange = parseSelectionRangeLabel(selectionLabel, selection.sheetName);

  const subscribeViewport = useCallback(
    (
      sheetName: string,
      viewport: Parameters<ProjectedViewportStore["subscribeViewport"]>[1],
      listener: Parameters<ProjectedViewportStore["subscribeViewport"]>[2],
    ) => {
      if (!runtimeController) {
        return () => {};
      }
      return runtimeController.subscribeViewport(sheetName, viewport, listener);
    },
    [runtimeController],
  );

  const {
    headerStatus: toolbarHeaderStatus,
    ribbon,
    selectionStatus,
    statusModeLabel,
  } = useWorkbookToolbar({
    connectionStateName: connectionState.name,
    runtimeReady,
    remoteSyncAvailable,
    zeroHealthReady,
    invokeMutation,
    selectionRange,
    selection,
    selectionLabel,
    selectedCell,
    selectedStyle,
    writesAllowed,
  });

  const createSheet = useCallback(() => {
    const nextSheetName = createNextSheetName(sheetNames);
    void invokeMutation("renderCommit", [
      {
        kind: "upsertSheet",
        name: nextSheetName,
        order: sheetNames.length,
      } satisfies CommitOp,
    ])
      .then(() => selectAddress(nextSheetName, "A1"))
      .catch(reportRuntimeError);
  }, [invokeMutation, reportRuntimeError, selectAddress, sheetNames]);

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
    [invokeMutation, reportRuntimeError, selectAddress, sheetNames],
  );

  const headerStatus = useMemo(() => {
    if (collaborators.length === 0) {
      return toolbarHeaderStatus;
    }
    return (
      <>
        {toolbarHeaderStatus}
        <WorkbookPresenceBar
          collaborators={collaborators}
          onJump={(sheetName, address) => {
            selectAddress(sheetName, address);
          }}
        />
      </>
    );
  }, [collaborators, selectAddress, toolbarHeaderStatus]);

  return {
    beginEditing,
    cancelEditor,
    clearSelectedCell,
    columnWidths,
    commitEditor,
    copySelectionRange,
    createSheet,
    editorSelectionBehavior,
    fillSelectionRange,
    handleEditorChange,
    headerStatus,
    invokeColumnWidthMutation,
    isEditing,
    isEditingCell,
    moveSelectionRange,
    pasteIntoSelection,
    remoteSyncAvailable,
    renameSheet,
    reportRuntimeError,
    resolvedValue,
    ribbon,
    runtimeError,
    runtimeReady,
    selectAddress,
    selectedCell,
    selection,
    selectionStatus,
    setSelectionLabel,
    sheetNames,
    statusModeLabel,
    subscribeViewport,
    toggleBooleanCell,
    visibleEditorValue,
    workbookReady,
    workerHandle,
    writesAllowed,
  };
}
