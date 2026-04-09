import { describe, expect, it, vi } from "vitest";
import { createWorkbookPerfSession } from "../perf/workbook-perf.js";

function createPerformanceRecorder() {
  const marks: string[] = [];
  const measures: { name: string; start?: string; end?: string }[] = [];
  return {
    performance: {
      mark(name: string) {
        marks.push(name);
      },
      measure(
        name: string,
        startOrOptions?: string | { start?: string; end?: string },
        endMark?: string,
      ) {
        if (typeof startOrOptions === "string") {
          measures.push({ name, start: startOrOptions, end: endMark });
          return;
        }
        measures.push({
          name,
          start: startOrOptions?.start,
          end: startOrOptions?.end,
        });
      },
    },
    marks,
    measures,
  };
}

describe("workbook perf session", () => {
  it("marks startup events once and records measures from the session start", () => {
    const recorder = createPerformanceRecorder();
    const session = createWorkbookPerfSession({
      documentId: "perf-doc",
      performance: recorder.performance,
      scope: "perf-doc:session",
    });

    session.markShellMounted();
    session.markShellMounted();
    session.noteBootstrapResult({
      restoredFromPersistence: true,
      requiresAuthoritativeHydrate: false,
    });
    session.markFirstSelectionVisible();
    session.markFirstSelectionVisible();

    expect(recorder.marks).toEqual([
      "perf-doc:session:start",
      "perf-doc:session:shell-mounted",
      "perf-doc:session:local-restore-ready",
      "perf-doc:session:first-selection-visible",
    ]);
    expect(recorder.measures).toEqual([
      {
        name: "perf-doc:session:time-to-shell-mounted",
        start: "perf-doc:session:start",
        end: "perf-doc:session:shell-mounted",
      },
      {
        name: "perf-doc:session:time-to-local-restore-ready",
        start: "perf-doc:session:start",
        end: "perf-doc:session:local-restore-ready",
      },
      {
        name: "perf-doc:session:time-to-first-selection-visible",
        start: "perf-doc:session:start",
        end: "perf-doc:session:first-selection-visible",
      },
    ]);
  });

  it("records the hydrate-required path when authoritative bootstrap is needed", () => {
    const recorder = createPerformanceRecorder();
    const session = createWorkbookPerfSession({
      documentId: "perf-doc",
      performance: recorder.performance,
      scope: "perf-doc:hydrate",
    });

    session.noteBootstrapResult({
      restoredFromPersistence: false,
      requiresAuthoritativeHydrate: true,
    });

    expect(recorder.marks).toEqual([
      "perf-doc:hydrate:start",
      "perf-doc:hydrate:authoritative-hydrate-required",
    ]);
    expect(recorder.measures).toEqual([
      {
        name: "perf-doc:hydrate:time-to-authoritative-hydrate-required",
        start: "perf-doc:hydrate:start",
        end: "perf-doc:hydrate:authoritative-hydrate-required",
      },
    ]);
  });

  it("stays non-throwing when the performance api rejects a mark", () => {
    const session = createWorkbookPerfSession({
      documentId: "perf-doc",
      performance: {
        mark: vi.fn(() => {
          throw new Error("mark failed");
        }),
        measure: vi.fn(),
      },
      scope: "perf-doc:failing",
    });

    expect(() => {
      session.markShellMounted();
      session.noteBootstrapResult({
        restoredFromPersistence: false,
        requiresAuthoritativeHydrate: true,
      });
      session.markFirstSelectionVisible();
    }).not.toThrow();
  });
});
