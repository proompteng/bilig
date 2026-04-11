import { useCallback, useEffect, useMemo, useRef, useState } from "react";

type WorkbookLocalPersistenceMode = "persistent" | "ephemeral" | "follower";

export interface WorkbookPersistenceTransferRequest {
  readonly requestId: string;
  readonly requesterTabId: string;
  readonly requestedAtUnixMs: number;
}

type WorkbookPersistenceHandoffMessage =
  | {
      readonly type: "takeover-request";
      readonly fromTabId: string;
      readonly requestId: string;
      readonly requestedAtUnixMs: number;
    }
  | {
      readonly type: "takeover-released";
      readonly fromTabId: string;
      readonly requestId: string;
    };

function createChannelName(documentId: string): string {
  return `bilig:workbook-local-persistence-handoff:${documentId}`;
}

function createRequestId(tabId: string): string {
  return `${tabId}:${Date.now().toString(36)}:${Math.random().toString(36).slice(2)}`;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isWorkbookPersistenceHandoffMessage(
  value: unknown,
): value is WorkbookPersistenceHandoffMessage {
  if (
    !isRecord(value) ||
    typeof value["type"] !== "string" ||
    typeof value["fromTabId"] !== "string"
  ) {
    return false;
  }
  if (value["type"] === "takeover-request") {
    return typeof value["requestId"] === "string" && typeof value["requestedAtUnixMs"] === "number";
  }
  if (value["type"] === "takeover-released") {
    return typeof value["requestId"] === "string";
  }
  return false;
}

export function useWorkbookLocalPersistenceHandoff(input: {
  documentId: string;
  localPersistenceMode: WorkbookLocalPersistenceMode;
  retryRuntime: (persistState: boolean) => void;
}) {
  const { documentId, localPersistenceMode, retryRuntime } = input;
  const tabId = useMemo(() => `tab:${Math.random().toString(36).slice(2)}`, []);
  const channelRef = useRef<BroadcastChannel | null>(null);
  const localPersistenceModeRef = useRef(localPersistenceMode);
  const retryRuntimeRef = useRef(retryRuntime);
  const requestedTakeoverRequestIdRef = useRef<string | null>(null);
  const releaseRequestIdRef = useRef<string | null>(null);
  const [pendingTransferRequest, setPendingTransferRequest] =
    useState<WorkbookPersistenceTransferRequest | null>(null);
  const [transferRequested, setTransferRequested] = useState(false);

  const postMessage = useCallback((message: WorkbookPersistenceHandoffMessage) => {
    channelRef.current?.postMessage(message);
  }, []);

  useEffect(() => {
    localPersistenceModeRef.current = localPersistenceMode;
    retryRuntimeRef.current = retryRuntime;
  }, [localPersistenceMode, retryRuntime]);

  useEffect(() => {
    if (typeof BroadcastChannel !== "function") {
      return;
    }
    const channel = new BroadcastChannel(createChannelName(documentId));
    channelRef.current = channel;
    const handleMessage = (event: MessageEvent<unknown>) => {
      const message = event.data;
      if (!isWorkbookPersistenceHandoffMessage(message) || message.fromTabId === tabId) {
        return;
      }
      if (message.type === "takeover-request") {
        if (localPersistenceModeRef.current !== "persistent") {
          return;
        }
        setPendingTransferRequest((current) => {
          if (current && current.requestedAtUnixMs > message.requestedAtUnixMs) {
            return current;
          }
          return {
            requestId: message.requestId,
            requesterTabId: message.fromTabId,
            requestedAtUnixMs: message.requestedAtUnixMs,
          };
        });
        return;
      }
      if (localPersistenceModeRef.current !== "follower") {
        return;
      }
      if (
        requestedTakeoverRequestIdRef.current &&
        requestedTakeoverRequestIdRef.current !== message.requestId
      ) {
        return;
      }
      requestedTakeoverRequestIdRef.current = null;
      setTransferRequested(false);
      retryRuntimeRef.current(true);
    };
    channel.addEventListener("message", handleMessage);
    return () => {
      channelRef.current = null;
      channel.removeEventListener("message", handleMessage);
      channel.close();
    };
  }, [documentId, tabId]);

  useEffect(() => {
    if (localPersistenceMode === "persistent") {
      requestedTakeoverRequestIdRef.current = null;
      setTransferRequested(false);
    }
    if (localPersistenceMode === "ephemeral" || localPersistenceMode === "follower") {
      setPendingTransferRequest(null);
    }
  }, [localPersistenceMode]);

  useEffect(() => {
    if (!releaseRequestIdRef.current || localPersistenceMode === "persistent") {
      return;
    }
    postMessage({
      type: "takeover-released",
      fromTabId: tabId,
      requestId: releaseRequestIdRef.current,
    });
    releaseRequestIdRef.current = null;
  }, [localPersistenceMode, postMessage, tabId]);

  const requestPersistenceTransfer = useCallback(() => {
    const requestId = createRequestId(tabId);
    requestedTakeoverRequestIdRef.current = requestId;
    setTransferRequested(true);
    postMessage({
      type: "takeover-request",
      fromTabId: tabId,
      requestId,
      requestedAtUnixMs: Date.now(),
    });
    retryRuntimeRef.current(true);
  }, [postMessage, tabId]);

  const approvePersistenceTransfer = useCallback(() => {
    if (!pendingTransferRequest) {
      return;
    }
    releaseRequestIdRef.current = pendingTransferRequest.requestId;
    setPendingTransferRequest(null);
    retryRuntimeRef.current(false);
  }, [pendingTransferRequest]);

  const dismissPersistenceTransferRequest = useCallback(() => {
    setPendingTransferRequest(null);
  }, []);

  return {
    pendingTransferRequest,
    transferRequested,
    requestPersistenceTransfer,
    approvePersistenceTransfer,
    dismissPersistenceTransferRequest,
  };
}
