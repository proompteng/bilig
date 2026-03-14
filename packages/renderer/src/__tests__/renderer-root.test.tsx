import React from "react";
import { describe, expect, it, vi } from "vitest";
import { SpreadsheetEngine } from "@bilig/core";
import { Cell, Sheet, Workbook } from "../components.js";
import { createWorkbookRendererRoot } from "../renderer-root.js";

describe("createWorkbookRendererRoot", () => {
  it("commits workbook DSL trees into the engine and clears on unmount", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "renderer-root-test" });
    await engine.ready();
    const root = createWorkbookRendererRoot(engine);

    await root.render(
      <Workbook name="test">
        <Sheet name="Sheet1">
          <Cell addr="A1" value={10} format="currency-usd" />
          <Cell addr="B1" formula="A1*2" />
        </Sheet>
      </Workbook>
    );

    expect(engine.getCell("Sheet1", "A1").format).toBe("currency-usd");
    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: 1, value: 20 });

    await root.unmount();

    expect(engine.getCellValue("Sheet1", "A1")).toEqual({ tag: 0 });
  });

  it("keeps rerenders idempotent and stable under StrictMode", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "renderer-root-strict" });
    await engine.ready();
    const root = createWorkbookRendererRoot(engine);
    const batches: number[] = [];
    const unsubscribe = engine.subscribe((event) => {
      batches.push(event.metrics.batchId);
    });

    const tree = (
      <React.StrictMode>
        <Workbook name="strict">
          <Sheet name="Sheet1">
            <Cell addr="A1" value={10} />
            <Cell addr="B1" formula="A1*2" />
          </Sheet>
        </Workbook>
      </React.StrictMode>
    );

    await root.render(tree);
    await root.render(tree);

    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: 1, value: 20 });
    expect(batches).toHaveLength(1);

    unsubscribe();
    await root.unmount();
  });

  it("rejects invalid workbook DSL trees without mutating the engine", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "renderer-root-invalid" });
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
        workbook: { name: "renderer-root-invalid" },
        sheets: []
      });
    } finally {
      consoleError.mockRestore();
    }
  });

  it("supports wrapper nodes and applies renames without leaking stale sheets or cells", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "renderer-root-updates" });
    await engine.ready();
    const root = createWorkbookRendererRoot(engine);

    await root.render(
      <>
        <Workbook name="book">
          <React.Fragment>
            <Sheet name="Sheet1">
              <Cell addr="A1" value={10} />
            </Sheet>
          </React.Fragment>
        </Workbook>
      </>
    );

    await root.render(
      <Workbook name="book-renamed">
        <Sheet name="Renamed">
          <Cell addr="B2" value={21} />
        </Sheet>
      </Workbook>
    );

    expect(engine.exportSnapshot()).toEqual({
      version: 1,
      workbook: { name: "book-renamed" },
      sheets: [
        {
          name: "Renamed",
          order: 0,
          cells: [{ address: "B2", value: 21 }]
        }
      ]
    });
    expect(engine.getCellValue("Sheet1", "A1")).toEqual({ tag: 0 });
  });

  it("rejects text nodes inside sheets without mutating committed state", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "renderer-root-text" });
    await engine.ready();
    const root = createWorkbookRendererRoot(engine);

    await root.render(
      <Workbook name="valid">
        <Sheet name="Sheet1">
          <Cell addr="A1" value={10} />
        </Sheet>
      </Workbook>
    );

    await expect(
      root.render(
        <Workbook name="invalid">
          <Sheet name="Sheet1">
            {"bad"}
          </Sheet>
        </Workbook>
      )
    ).rejects.toThrow("Workbook DSL does not support text nodes.");

    expect(engine.exportSnapshot()).toEqual({
      version: 1,
      workbook: { name: "valid" },
      sheets: [
        {
          name: "Sheet1",
          order: 0,
          cells: [{ address: "A1", value: 10 }]
        }
      ]
    });
  });

  it("rejects invalid root, sheet, and cell contracts before commit", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "renderer-root-contracts" });
    await engine.ready();
    const root = createWorkbookRendererRoot(engine);

    await expect(
      root.render(
        <Sheet name="Sheet1">
          <Cell addr="A1" value={1} />
        </Sheet>
      )
    ).rejects.toThrow("Root descriptor must be a Workbook.");

    await expect(
      root.render(
        <Workbook name="missing-sheet-name">
          <Sheet name={""}>
            <Cell addr="A1" value={1} />
          </Sheet>
        </Workbook>
      )
    ).rejects.toThrow("<Sheet> requires a name prop.");

    await expect(
      root.render(
        <Workbook name="missing-addr">
          <Sheet name="Sheet1">
            <Cell addr={""} value={1} />
          </Sheet>
        </Workbook>
      )
    ).rejects.toThrow("<Cell> requires an addr prop.");

    await expect(
      root.render(
        <Workbook name="conflicting-cell">
          <Sheet name="Sheet1">
            <Cell addr="A1" value={1} formula="B1" />
          </Sheet>
        </Workbook>
      )
    ).rejects.toThrow("<Cell> cannot specify both value and formula.");

    expect(engine.exportSnapshot()).toEqual({
      version: 1,
      workbook: { name: "renderer-root-contracts" },
      sheets: []
    });
  });

  it("allows clearing the rendered workbook by rendering null", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "renderer-root-clear" });
    await engine.ready();
    const root = createWorkbookRendererRoot(engine);

    await root.render(
      <Workbook name="clearable">
        <Sheet name="Sheet1">
          <Cell addr="A1" value={10} />
        </Sheet>
      </Workbook>
    );

    await root.render(null);

    expect(engine.exportSnapshot()).toEqual({
      version: 1,
      workbook: { name: "clearable" },
      sheets: []
    });
  });
});
