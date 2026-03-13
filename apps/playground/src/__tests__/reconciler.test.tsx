import React from "react";
import { describe, expect, it, vi } from "vitest";
import { SpreadsheetEngine } from "@bilig/core";
import { Cell, Sheet, Workbook, createWorkbookRendererRoot } from "../reconciler/index.js";

describe("playground reconciler", () => {
  it("commits workbook DSL into the engine", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "reconciler-test" });
    await engine.ready();
    const root = createWorkbookRendererRoot(engine);

    await root.render(
      <Workbook name="test">
        <Sheet name="Sheet1">
          <Cell addr="A1" value={10} />
          <Cell addr="B1" formula="A1*2" />
        </Sheet>
      </Workbook>
    );

    expect(engine.getCell("Sheet1", "B1").value).toEqual({ tag: 1, value: 20 });
    await root.unmount();
  });

  it("batches rerenders into a single engine commit per render", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "reconciler-batch-test" });
    await engine.ready();
    const root = createWorkbookRendererRoot(engine);
    const batches: number[] = [];
    const unsubscribe = engine.subscribe((event) => {
      batches.push(event.metrics.batchId);
    });

    await root.render(
      <Workbook name="batch-test">
        <Sheet name="Sheet1">
          <Cell addr="A1" value={10} />
          <Cell addr="B1" formula="A1*2" />
        </Sheet>
      </Workbook>
    );

    await root.render(
      <Workbook name="batch-test">
        <Sheet name="Sheet1">
          <Cell addr="A1" value={21} />
          <Cell addr="B1" formula="A1*2" />
        </Sheet>
      </Workbook>
    );

    expect(engine.getCell("Sheet1", "B1").value).toEqual({ tag: 1, value: 42 });
    expect(batches).toHaveLength(2);

    unsubscribe();
    await root.unmount();
  });

  it("does not emit a new engine batch for an identical rerender", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "reconciler-idempotent-test" });
    await engine.ready();
    const root = createWorkbookRendererRoot(engine);
    const batches: number[] = [];
    const unsubscribe = engine.subscribe((event) => {
      batches.push(event.metrics.batchId);
    });

    await root.render(
      <Workbook name="same-tree">
        <Sheet name="Sheet1">
          <Cell addr="A1" value={10} />
          <Cell addr="B1" formula="A1*2" />
        </Sheet>
      </Workbook>
    );
    await root.render(
      <Workbook name="same-tree">
        <Sheet name="Sheet1">
          <Cell addr="A1" value={10} />
          <Cell addr="B1" formula="A1*2" />
        </Sheet>
      </Workbook>
    );

    expect(engine.getCell("Sheet1", "B1").value).toEqual({ tag: 1, value: 20 });
    expect(batches).toHaveLength(1);

    unsubscribe();
    await root.unmount();
  });

  it("keeps commit behavior stable under StrictMode", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "reconciler-strict-mode-test" });
    await engine.ready();
    const root = createWorkbookRendererRoot(engine);
    const batches: number[] = [];
    const unsubscribe = engine.subscribe((event) => {
      batches.push(event.metrics.batchId);
    });

    await root.render(
      <React.StrictMode>
        <Workbook name="strict">
          <Sheet name="Sheet1">
            <Cell addr="A1" value={10} />
            <Cell addr="B1" formula="A1*2" />
          </Sheet>
        </Workbook>
      </React.StrictMode>
    );

    expect(engine.getCell("Sheet1", "B1").value).toEqual({ tag: 1, value: 20 });
    expect(batches).toHaveLength(1);

    unsubscribe();
    await root.unmount();
  });

  it("rejects invalid workbook trees without mutating the engine", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "reconciler-invalid-test" });
    await engine.ready();
    const root = createWorkbookRendererRoot(engine);
    const consoleError = vi.spyOn(console, "error").mockImplementation(() => {});

    try {
      await expect(
        root.render(
          <Workbook name="invalid">
            <Cell addr="A1" value={10} />
          </Workbook>
        )
      ).rejects.toThrow("Only <Sheet> nodes can exist under <Workbook>.");

      expect(engine.exportSnapshot()).toEqual({
        version: 1,
        workbook: { name: "reconciler-invalid-test" },
        sheets: []
      });
    } finally {
      consoleError.mockRestore();
    }
  });

  it("rejects duplicate sheet names without mutating the engine", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "reconciler-duplicate-sheet-test" });
    await engine.ready();
    const root = createWorkbookRendererRoot(engine);
    const consoleError = vi.spyOn(console, "error").mockImplementation(() => {});

    try {
      await expect(
        root.render(
          <Workbook name="duplicate-sheets">
            <Sheet name="Sheet1">
              <Cell addr="A1" value={10} />
            </Sheet>
            <Sheet name="Sheet1">
              <Cell addr="B1" value={11} />
            </Sheet>
          </Workbook>
        )
      ).rejects.toThrow("Duplicate sheet name 'Sheet1'.");

      expect(engine.exportSnapshot()).toEqual({
        version: 1,
        workbook: { name: "reconciler-duplicate-sheet-test" },
        sheets: []
      });
    } finally {
      consoleError.mockRestore();
    }
  });

  it("clears workbook state on unmount", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "reconciler-unmount-test" });
    await engine.ready();
    const root = createWorkbookRendererRoot(engine);

    await root.render(
      <Workbook name="unmount-test">
        <Sheet name="Sheet1">
          <Cell addr="A1" value={10} />
        </Sheet>
      </Workbook>
    );

    await root.unmount();

    expect(engine.getCell("Sheet1", "A1").value).toEqual({ tag: 0 });
  });
});
