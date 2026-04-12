import { describe, expect, it } from "vitest";
import type { CodexThreadItem } from "@bilig/agent-api";
import { mapThreadItemToEntry } from "./workbook-agent-session-model.js";

describe("workbook-agent-session-model", () => {
  it("maps reasoning items with summary arrays into reasoning timeline entries", () => {
    const entry = mapThreadItemToEntry(
      {
        type: "reasoning",
        id: "reasoning-1",
        summary: [
          {
            type: "summary_text",
            text: "Inspecting the workbook structure before changing formulas.",
          },
        ],
      } satisfies CodexThreadItem,
      "turn-1",
    );

    expect(entry).toEqual({
      id: "reasoning-1",
      kind: "reasoning",
      turnId: "turn-1",
      text: "Inspecting the workbook structure before changing formulas.",
      phase: null,
      toolName: null,
      toolStatus: null,
      argumentsText: null,
      outputText: null,
      success: null,
      citations: [],
    });
  });

  it("keeps empty reasoning items as reasoning entries so deltas can hydrate them later", () => {
    const entry = mapThreadItemToEntry(
      {
        type: "reasoning",
        id: "reasoning-1",
      } satisfies CodexThreadItem,
      "turn-1",
    );

    expect(entry).toEqual({
      id: "reasoning-1",
      kind: "reasoning",
      turnId: "turn-1",
      text: "",
      phase: null,
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

  it("normalizes legacy workbook tool names in timeline entries", () => {
    const entry = mapThreadItemToEntry(
      {
        type: "dynamicToolCall",
        id: "tool-1",
        tool: "bilig_read_workbook",
        status: "completed",
        arguments: { sheetName: "Sheet1" },
        contentItems: [{ type: "inputText", text: '{"ok":true}' }],
        success: true,
      } satisfies CodexThreadItem,
      "turn-1",
    );

    expect(entry).toEqual(
      expect.objectContaining({
        kind: "tool",
        toolName: "read_workbook",
        toolStatus: "completed",
      }),
    );
  });
});
