import { useEffect, useState } from "react";
import { queries } from "@bilig/zero-sync";
import {
  WorkbookAgentThreadSummarySchema,
  WorkbookAgentWorkflowRunSchema,
  decodeUnknownSync,
  type WorkbookAgentThreadSummary,
  type WorkbookAgentWorkflowRun,
} from "@bilig/contracts";
import { Schema } from "effect";

interface ZeroLiveView<T> {
  readonly data: T;
  addListener(listener: (value: T) => void): () => void;
  destroy(): void;
}

export interface ZeroWorkbookAgentSource {
  materialize(query: unknown): unknown;
}

const WorkbookAgentThreadSummaryListSchema = Schema.Array(WorkbookAgentThreadSummarySchema);
const WorkbookAgentWorkflowRunListSchema = Schema.Array(WorkbookAgentWorkflowRunSchema);

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

function decodeThreadSummaries(value: unknown): readonly WorkbookAgentThreadSummary[] {
  if (!Array.isArray(value)) {
    return [];
  }
  return decodeUnknownSync(
    WorkbookAgentThreadSummaryListSchema,
    value.map((entry) =>
      isRecord(entry)
        ? {
            ...entry,
            latestEntryText: entry["latestEntryText"] ?? null,
          }
        : entry,
    ),
  );
}

function decodeWorkflowRuns(value: unknown): readonly WorkbookAgentWorkflowRun[] {
  if (!Array.isArray(value)) {
    return [];
  }
  return decodeUnknownSync(
    WorkbookAgentWorkflowRunListSchema,
    value.map((entry) =>
      isRecord(entry)
        ? {
            ...entry,
            completedAtUnixMs: entry["completedAtUnixMs"] ?? null,
            errorMessage: entry["errorMessage"] ?? null,
            steps: entry["steps"] ?? [],
            artifact: entry["artifact"] ?? null,
          }
        : entry,
    ),
  );
}

export function useWorkbookAgentThreadSummaries(input: {
  readonly documentId: string;
  readonly zero: ZeroWorkbookAgentSource;
  readonly enabled: boolean;
}): readonly WorkbookAgentThreadSummary[] {
  const { documentId, enabled, zero } = input;
  const [threadSummaries, setThreadSummaries] = useState<readonly WorkbookAgentThreadSummary[]>([]);

  useEffect(() => {
    if (!enabled) {
      setThreadSummaries([]);
      return;
    }
    const view = zero.materialize(queries.workbookChatThread.byWorkbook({ documentId }));
    if (!isZeroLiveView<unknown>(view)) {
      throw new Error("Zero workbook agent thread query returned an invalid live view");
    }
    const publish = (value: unknown) => {
      setThreadSummaries(decodeThreadSummaries(value));
    };
    publish(view.data);
    const cleanup = view.addListener((value) => {
      publish(value);
    });
    return () => {
      cleanup();
      view.destroy();
    };
  }, [documentId, enabled, zero]);

  return threadSummaries;
}

export function useWorkbookAgentWorkflowRuns(input: {
  readonly documentId: string;
  readonly threadId: string | null;
  readonly zero: ZeroWorkbookAgentSource;
  readonly enabled: boolean;
}): readonly WorkbookAgentWorkflowRun[] {
  const { documentId, enabled, threadId, zero } = input;
  const [workflowRuns, setWorkflowRuns] = useState<readonly WorkbookAgentWorkflowRun[]>([]);

  useEffect(() => {
    if (!enabled || !threadId) {
      setWorkflowRuns([]);
      return;
    }
    const view = zero.materialize(queries.workbookWorkflowRun.byThread({ documentId, threadId }));
    if (!isZeroLiveView<unknown>(view)) {
      throw new Error("Zero workbook agent workflow query returned an invalid live view");
    }
    const publish = (value: unknown) => {
      setWorkflowRuns(decodeWorkflowRuns(value));
    };
    publish(view.data);
    const cleanup = view.addListener((value) => {
      publish(value);
    });
    return () => {
      cleanup();
      view.destroy();
    };
  }, [documentId, enabled, threadId, zero]);

  return workflowRuns;
}
