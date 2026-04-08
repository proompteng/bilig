import { useCallback, useEffect, useRef, type MutableRefObject } from "react";
import { flushSync } from "react-dom";
import type { WorkerHandle, WorkerRuntimeSessionController } from "./runtime-session.js";
import {
  buildZeroWorkbookMutation,
  isCellNumberFormatInputValue,
  isCellRangeRef,
  isCellStyleFieldList,
  isCellStylePatchValue,
  isCommitOps,
  isLiteralInput,
  isPendingWorkbookMutation,
  isPendingWorkbookMutationList,
  type PendingWorkbookMutation,
  type PendingWorkbookMutationInput,
  type WorkbookMutationMethod,
} from "./workbook-sync.js";
import {
  assert,
  canAttemptRemoteSync,
  isMutationErrorResult,
  toErrorMessage,
  type ZeroConnectionState,
} from "./worker-workbook-app-model.js";

interface ZeroMutationSource {
  mutate(mutation: unknown): unknown;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function observeZeroMutationResult(result: unknown): Promise<unknown> | null {
  if (!isRecord(result)) {
    return null;
  }
  const observer = result["server"] ?? result["client"];
  return observer instanceof Promise ? observer : null;
}

export function useWorkbookSync(input: {
  documentId: string;
  connectionStateName: ZeroConnectionState["name"];
  connectionStateRef: MutableRefObject<ZeroConnectionState["name"]>;
  runtimeController: WorkerRuntimeSessionController | null;
  workerHandleRef: MutableRefObject<WorkerHandle | null>;
  zeroRef: MutableRefObject<ZeroMutationSource>;
  reportRuntimeError: (error: unknown) => void;
}) {
  const {
    documentId,
    connectionStateName,
    connectionStateRef,
    runtimeController,
    workerHandleRef,
    zeroRef,
    reportRuntimeError,
  } = input;
  const localMutationQueueRef = useRef<Promise<void>>(Promise.resolve());
  const syncQueueRef = useRef<Promise<void>>(Promise.resolve());

  const runSerializedSyncTask = useCallback(
    async (task: () => Promise<unknown>): Promise<unknown> => {
      const previousTask = syncQueueRef.current;
      let releaseQueue = () => {};
      syncQueueRef.current = new Promise<void>((resolve) => {
        releaseQueue = resolve;
      });
      await previousTask.catch(() => {});
      try {
        return await task();
      } finally {
        releaseQueue();
      }
    },
    [],
  );

  const runSerializedLocalMutationTask = useCallback(
    async (task: () => Promise<unknown>): Promise<unknown> => {
      const previousTask = localMutationQueueRef.current;
      let releaseQueue = () => {};
      localMutationQueueRef.current = new Promise<void>((resolve) => {
        releaseQueue = resolve;
      });
      await previousTask.catch(() => {});
      try {
        return await task();
      } finally {
        releaseQueue();
      }
    },
    [],
  );

  const listPendingMutations = useCallback(async (): Promise<
    readonly PendingWorkbookMutation[]
  > => {
    if (!runtimeController) {
      throw new Error("Workbook runtime is not ready");
    }
    const value = await runtimeController.invoke("listPendingMutations");
    assert(
      isPendingWorkbookMutationList(value),
      "Worker returned an invalid pending workbook mutation list",
    );
    return value;
  }, [runtimeController]);

  const enqueuePendingMutation = useCallback(
    async (mutation: PendingWorkbookMutationInput): Promise<PendingWorkbookMutation> => {
      if (!runtimeController) {
        throw new Error("Workbook runtime is not ready");
      }
      const value = await runtimeController.invoke("enqueuePendingMutation", mutation);
      assert(isPendingWorkbookMutation(value), "Worker returned an invalid pending mutation");
      return value;
    },
    [runtimeController],
  );

  const runZeroMutation = useCallback(
    async (
      mutation: PendingWorkbookMutation,
    ): Promise<{ ok: true } | { ok: false; retryable: boolean; error: Error }> => {
      try {
        const result = zeroRef.current.mutate(buildZeroWorkbookMutation(documentId, mutation));
        const observerResult = observeZeroMutationResult(result);
        if (!observerResult) {
          return { ok: true };
        }
        const remoteResult = await observerResult;
        if (!isMutationErrorResult(remoteResult)) {
          return { ok: true };
        }
        const details =
          remoteResult.error.type === "app" && remoteResult.error.details !== undefined
            ? ` (${JSON.stringify(remoteResult.error.details)})`
            : "";
        return {
          ok: false,
          retryable: remoteResult.error.type === "zero",
          error: new Error(`${remoteResult.error.message}${details}`),
        };
      } catch (error) {
        return {
          ok: false,
          retryable: true,
          error: error instanceof Error ? error : new Error(toErrorMessage(error)),
        };
      }
    },
    [documentId, zeroRef],
  );

  const drainPendingMutationsLocked = useCallback(async (): Promise<void> => {
    if (!runtimeController || !canAttemptRemoteSync(connectionStateRef.current)) {
      return;
    }

    const drainBatch = async (
      pendingMutations: readonly PendingWorkbookMutation[],
      index = 0,
    ): Promise<void> => {
      const mutation = pendingMutations[index];
      if (!mutation || !canAttemptRemoteSync(connectionStateRef.current)) {
        return;
      }

      if (mutation.status === "submitted") {
        await drainBatch(pendingMutations, index + 1);
        return;
      }

      const remoteResult = await runZeroMutation(mutation);
      if (!remoteResult.ok) {
        if (!remoteResult.retryable) {
          throw remoteResult.error;
        }
        return;
      }

      await runtimeController.invoke("markPendingMutationSubmitted", mutation.id);
      await drainBatch(pendingMutations, index + 1);
    };

    await drainBatch(await listPendingMutations());
  }, [connectionStateRef, listPendingMutations, runZeroMutation, runtimeController]);

  const drainPendingMutations = useCallback(async (): Promise<void> => {
    try {
      await runSerializedSyncTask(drainPendingMutationsLocked);
    } catch (error) {
      reportRuntimeError(error);
    }
  }, [drainPendingMutationsLocked, reportRuntimeError, runSerializedSyncTask]);

  const invokeMutation = useCallback(
    async (method: WorkbookMutationMethod, ...args: unknown[]): Promise<void> => {
      if (!runtimeController) {
        throw new Error("Workbook runtime is not ready");
      }

      let mutation: PendingWorkbookMutationInput;
      switch (method) {
        case "setCellValue": {
          const [sheetName, address, value] = args;
          assert(
            typeof sheetName === "string" && typeof address === "string" && isLiteralInput(value),
            "Invalid setCellValue args",
          );
          mutation = { method, args: [sheetName, address, value] };
          break;
        }
        case "setCellFormula": {
          const [sheetName, address, formula] = args;
          assert(
            typeof sheetName === "string" &&
              typeof address === "string" &&
              typeof formula === "string",
            "Invalid setCellFormula args",
          );
          mutation = { method, args: [sheetName, address, formula] };
          break;
        }
        case "clearCell": {
          const [sheetName, address] = args;
          assert(
            typeof sheetName === "string" && typeof address === "string",
            "Invalid clearCell args",
          );
          mutation = { method, args: [sheetName, address] };
          break;
        }
        case "clearRange": {
          const [range] = args;
          assert(isCellRangeRef(range), "Invalid clearRange args");
          mutation = { method, args: [range] };
          break;
        }
        case "renderCommit": {
          const [ops] = args;
          assert(isCommitOps(ops), "Invalid renderCommit args");
          mutation = { method, args: [ops] };
          break;
        }
        case "fillRange":
        case "copyRange":
        case "moveRange": {
          const [source, target] = args;
          assert(isCellRangeRef(source) && isCellRangeRef(target), `Invalid ${method} args`);
          mutation = { method, args: [source, target] };
          break;
        }
        case "updateColumnWidth": {
          const [sheetName, columnIndex, width] = args;
          assert(
            typeof sheetName === "string" &&
              typeof columnIndex === "number" &&
              typeof width === "number",
            "Invalid updateColumnWidth args",
          );
          mutation = { method, args: [sheetName, columnIndex, width] };
          break;
        }
        case "setRangeStyle": {
          const [range, patch] = args;
          assert(
            isCellRangeRef(range) && isCellStylePatchValue(patch),
            "Invalid setRangeStyle args",
          );
          mutation = { method, args: [range, patch] };
          break;
        }
        case "clearRangeStyle": {
          const [range, fields] = args;
          assert(
            isCellRangeRef(range) && (fields === undefined || isCellStyleFieldList(fields)),
            "Invalid clearRangeStyle args",
          );
          mutation = { method, args: [range, fields] };
          break;
        }
        case "setRangeNumberFormat": {
          const [range, format] = args;
          assert(
            isCellRangeRef(range) && isCellNumberFormatInputValue(format),
            "Invalid setRangeNumberFormat args",
          );
          mutation = { method, args: [range, format] };
          break;
        }
        case "clearRangeNumberFormat": {
          const [range] = args;
          assert(isCellRangeRef(range), "Invalid clearRangeNumberFormat args");
          mutation = { method, args: [range] };
          break;
        }
        default:
          throw new Error("Unsupported workbook mutation");
      }

      await runSerializedLocalMutationTask(() => enqueuePendingMutation(mutation));
      await runSerializedSyncTask(async () => {
        if (canAttemptRemoteSync(connectionStateRef.current)) {
          await drainPendingMutationsLocked();
        }
      });
    },
    [
      connectionStateRef,
      drainPendingMutationsLocked,
      enqueuePendingMutation,
      runSerializedLocalMutationTask,
      runSerializedSyncTask,
      runtimeController,
    ],
  );

  const invokeColumnWidthMutation = useCallback(
    async (
      sheetName: string,
      columnIndex: number,
      width: number,
      options?: { flush?: boolean },
    ): Promise<void> => {
      const viewportStore = workerHandleRef.current?.viewportStore;
      const previousWidth = viewportStore?.getColumnWidths(sheetName)[columnIndex];
      if (viewportStore) {
        const applyOptimisticWidth = () => {
          viewportStore.setColumnWidth(sheetName, columnIndex, width);
        };
        if (options?.flush) {
          flushSync(applyOptimisticWidth);
        } else {
          applyOptimisticWidth();
        }
      }
      try {
        await invokeMutation("updateColumnWidth", sheetName, columnIndex, width);
      } catch (error) {
        if (viewportStore && viewportStore.getColumnWidths(sheetName)[columnIndex] === width) {
          viewportStore.rollbackColumnWidth(sheetName, columnIndex, previousWidth);
        }
        throw error;
      }
    },
    [invokeMutation, workerHandleRef],
  );

  useEffect(() => {
    if (!runtimeController || !canAttemptRemoteSync(connectionStateName)) {
      return;
    }
    void drainPendingMutations();
  }, [connectionStateName, drainPendingMutations, runtimeController]);

  return {
    invokeMutation,
    invokeColumnWidthMutation,
  };
}
