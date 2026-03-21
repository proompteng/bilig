import { useEffect, useState } from "react";
import type { SpreadsheetEngine } from "@bilig/core";
import type { SyncState } from "@bilig/protocol";
import { createWebSocketSyncClient } from "./createWebSocketSyncClient.js";

export interface RemoteSpreadsheetSyncOptions {
  enabled: boolean;
  engine: SpreadsheetEngine;
  documentId: string;
  replicaId: string;
  baseUrl: string | null;
}

export function useRemoteSpreadsheetSync(options: RemoteSpreadsheetSyncOptions): SyncState {
  const [state, setState] = useState<SyncState>("local-only");

  useEffect(() => {
    if (!options.enabled || !options.baseUrl) {
      setState("local-only");
      return;
    }

    let disposed = false;
    setState("syncing");
    void options.engine.connectSyncClient(createWebSocketSyncClient({
      documentId: options.documentId,
      replicaId: options.replicaId,
      baseUrl: options.baseUrl
    })).catch((error: unknown) => {
      if (!disposed) {
        void error;
        setState("local-only");
      }
    });

    const interval = window.setInterval(() => {
      if (!disposed) {
        setState(options.engine.getSyncState());
      }
    }, 100);

    return () => {
      disposed = true;
      window.clearInterval(interval);
      void options.engine.disconnectSyncClient();
      setState("local-only");
    };
  }, [options.baseUrl, options.documentId, options.enabled, options.engine, options.replicaId]);

  useEffect(() => {
    if (!options.enabled) {
      return;
    }
    setState(options.engine.getSyncState());
  }, [options.enabled, options.engine]);

  return state;
}
