import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { mutators, queries } from "@bilig/zero-sync";
import type { WorkerRuntimeSelection } from "./runtime-session.js";
import {
  normalizeWorkbookPresenceRows,
  selectActiveWorkbookCollaborators,
  WORKBOOK_PRESENCE_HEARTBEAT_MS,
  WORKBOOK_PRESENCE_STALE_TICK_MS,
  type WorkbookCollaboratorPresence,
} from "./workbook-presence-model.js";

interface ZeroLiveView<T> {
  readonly data: T;
  addListener(listener: (value: T) => void): () => void;
  destroy(): void;
}

interface ZeroPresenceSource {
  materialize(query: unknown): unknown;
  mutate(mutation: unknown): unknown;
}

type UpdatePresenceArgs = Parameters<typeof mutators.workbook.updatePresence>[0] & {
  presenceClientId?: string;
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isZeroLiveView<T>(value: unknown): value is ZeroLiveView<T> {
  return (
    isRecord(value) &&
    "data" in value &&
    typeof value["addListener"] === "function" &&
    typeof value["destroy"] === "function"
  );
}

function observeZeroMutationResult(result: unknown): void {
  if (!isRecord(result)) {
    return;
  }
  const observer = result["server"] ?? result["client"];
  if (!(observer instanceof Promise)) {
    return;
  }
  void observer.catch(() => undefined);
}

export function useWorkbookPresence(input: {
  readonly documentId: string;
  readonly currentUserId: string;
  readonly currentPresenceClientId: string;
  readonly sessionId: string;
  readonly selection: WorkerRuntimeSelection;
  readonly sheetNames: readonly string[];
  readonly zero: ZeroPresenceSource;
  readonly enabled: boolean;
}): readonly WorkbookCollaboratorPresence[] {
  const {
    currentPresenceClientId,
    currentUserId,
    documentId,
    enabled,
    selection,
    sessionId,
    sheetNames,
    zero,
  } = input;
  const [presenceRows, setPresenceRows] = useState(
    [] as readonly ReturnType<typeof normalizeWorkbookPresenceRows>[number][],
  );
  const [now, setNow] = useState(() => Date.now());
  const latestSelectionRef = useRef(selection);

  latestSelectionRef.current = selection;

  const publishPresence = useCallback(() => {
    if (!enabled) {
      return;
    }
    const presenceArgs: UpdatePresenceArgs = {
      documentId,
      sessionId,
      presenceClientId: currentPresenceClientId,
      sheetName: latestSelectionRef.current.sheetName,
      address: latestSelectionRef.current.address,
      selection: latestSelectionRef.current,
    };
    observeZeroMutationResult(
      zero.mutate(mutators.workbook.updatePresence(presenceArgs)),
    );
  }, [currentPresenceClientId, documentId, enabled, sessionId, zero]);

  useEffect(() => {
    if (!enabled) {
      setPresenceRows([]);
      return;
    }
    const view = zero.materialize(queries.presenceCoarse.byWorkbook({ documentId }));
    if (!isZeroLiveView<unknown>(view)) {
      throw new Error("Zero workbook presence query returned an invalid live view");
    }
    const publishRows = (value: unknown) => {
      setPresenceRows(normalizeWorkbookPresenceRows(value));
      setNow(Date.now());
    };
    publishRows(view.data);
    const cleanup = view.addListener((value) => {
      publishRows(value);
    });
    return () => {
      cleanup();
      view.destroy();
    };
  }, [documentId, enabled, zero]);

  useEffect(() => {
    if (!enabled) {
      return;
    }
    const intervalId = window.setInterval(() => {
      publishPresence();
    }, WORKBOOK_PRESENCE_HEARTBEAT_MS);
    return () => {
      window.clearInterval(intervalId);
    };
  }, [enabled, publishPresence]);

  useEffect(() => {
    if (!enabled) {
      return;
    }
    publishPresence();
  }, [enabled, publishPresence, selection.address, selection.sheetName]);

  useEffect(() => {
    if (!enabled) {
      return;
    }
    const intervalId = window.setInterval(() => {
      setNow(Date.now());
    }, WORKBOOK_PRESENCE_STALE_TICK_MS);
    return () => {
      window.clearInterval(intervalId);
    };
  }, [enabled]);

  return useMemo(
    () =>
      selectActiveWorkbookCollaborators({
        rows: presenceRows,
        currentPresenceClientId,
        currentUserId,
        currentSessionId: sessionId,
        knownSheetNames: sheetNames,
        now,
      }),
    [currentPresenceClientId, currentUserId, now, presenceRows, sessionId, sheetNames],
  );
}
