import { describe, expect, it } from "vitest";
import type { CodexThreadItem } from "@bilig/agent-api";
import { mapThreadItemToEntry } from "./workbook-agent-session-model.js";

describe("workbook-agent-session-model", () => {
  it("maps reasoning items into plan timeline entries", () => {
    const entry = mapThreadItemToEntry(
      {
        type: "reasoning",
        id: "reasoning-1",
        text: "Inspecting the workbook structure before changing formulas.",
      } satisfies CodexThreadItem,
      "turn-1",
    );

    expect(entry).toEqual({
      id: "reasoning-1",
      kind: "plan",
      turnId: "turn-1",
      text: "Inspecting the workbook structure before changing formulas.",
      phase: "reasoning",
      toolName: null,
      toolStatus: null,
      argumentsText: null,
      outputText: null,
      success: null,
      citations: [],
    });
  });

  it("keeps unknown non-reasoning items as system entries", () => {
    const entry = mapThreadItemToEntry(
      {
        type: "status",
        id: "status-1",
      } satisfies CodexThreadItem,
      "turn-1",
    );

    expect(entry.kind).toBe("system");
    expect(entry.text).toBe("Codex emitted status.");
  });
});
