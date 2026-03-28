import React from "react";
import { describe, expect, it } from "vitest";
import { SpreadsheetEngine } from "@bilig/core";
import type { WorkbookContainer } from "../host-config.js";
import { createFiberRoot, updateFiberRoot } from "../compat.js";
import { Cell, Sheet, Workbook } from "../components.js";
import { createWorkbookRendererRoot } from "../renderer-root.js";

function createContainer(engine: SpreadsheetEngine): WorkbookContainer {
  return {
    engine,
    root: null,
    pendingOps: [],
    shouldSyncSheetOrders: false,
    lastError: null,
  };
}

describe("renderer root", () => {
  it("renders workbook DSL components even when function names are minified", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "renderer-minified-components" });
    await engine.ready();
    const root = createWorkbookRendererRoot(engine);

    const MinifiedWorkbook = Object.assign(
      function a(props: React.ComponentProps<typeof Workbook>) {
        return React.createElement(Workbook, props);
      },
      { __biligRendererKind: "Workbook" as const },
    );
    const MinifiedSheet = Object.assign(
      function b(props: React.ComponentProps<typeof Sheet>) {
        return React.createElement(Sheet, props);
      },
      { __biligRendererKind: "Sheet" as const },
    );
    const MinifiedCell = Object.assign(
      function c(props: React.ComponentProps<typeof Cell>) {
        return React.createElement(Cell, props);
      },
      { __biligRendererKind: "Cell" as const },
    );

    await root.render(
      React.createElement(
        MinifiedWorkbook,
        { name: "Book" },
        React.createElement(
          MinifiedSheet,
          { name: "Sheet1" },
          React.createElement(MinifiedCell, { addr: "A1", value: 10 }),
          React.createElement(MinifiedCell, { addr: "B1", formula: "A1*2" }),
        ),
      ),
    );

    expect(engine.getCellValue("Sheet1", "A1")).toEqual({ tag: 1, value: 10 });
    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: 1, value: 20 });

    await root.unmount();
  });

  it("renders, updates, clears, and unmounts workbook trees through the public root", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "renderer-root" });
    await engine.ready();
    const root = createWorkbookRendererRoot(engine);

    await root.render(
      React.createElement(
        Workbook,
        { name: "Book" },
        React.createElement(
          React.Fragment,
          null,
          React.createElement(
            Sheet,
            { name: "Sheet1" },
            React.createElement(Cell, { addr: "A1", value: 2 }),
            React.createElement(Cell, { addr: "B1", formula: "A1*3" }),
          ),
        ),
      ),
    );

    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: 1, value: 6 });

    await root.render(
      React.createElement(
        Workbook,
        { name: "Book" },
        React.createElement(
          React.Fragment,
          null,
          React.createElement(
            Sheet,
            { name: "Sheet1" },
            React.createElement(Cell, { addr: "A1", value: 5 }),
            React.createElement(Cell, { addr: "C1", formula: "A1+1", format: "currency-usd" }),
          ),
          React.createElement(
            Sheet,
            { name: "Sheet2" },
            React.createElement(Cell, { addr: "A1", value: "ready" }),
          ),
        ),
      ),
    );

    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: 0 });
    expect(engine.getCellValue("Sheet1", "C1")).toEqual({ tag: 1, value: 6 });
    expect(engine.getCell("Sheet1", "C1").format).toBe("currency-usd");
    expect(engine.getCellValue("Sheet2", "A1")).toMatchObject({ tag: 3, value: "ready" });

    await root.render(null);
    expect(engine.exportSnapshot().sheets).toEqual([]);

    await root.unmount();
    expect(engine.exportSnapshot().sheets).toEqual([]);
  });

  it("rejects invalid workbook DSL trees before mutating engine state", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "renderer-invalid" });
    await engine.ready();
    const root = createWorkbookRendererRoot(engine);

    await expect(root.render(React.createElement(Sheet, { name: "Sheet1" }))).rejects.toThrow(
      "Root descriptor must be a Workbook.",
    );
    await expect(
      root.render(
        React.createElement(
          Workbook,
          { name: "Book" },
          React.createElement(Sheet, { name: "Sheet1" }, "bad text child"),
        ),
      ),
    ).rejects.toThrow("Workbook DSL does not support text nodes.");
    await expect(
      root.render(
        React.createElement(
          Workbook,
          { name: "Book" },
          React.createElement(
            Sheet,
            { name: "Sheet1" },
            React.createElement(Cell, { addr: "A1", value: 1, formula: "B1" }),
          ),
        ),
      ),
    ).rejects.toThrow("<Cell> cannot specify both value and formula.");
    await expect(
      root.render(
        React.createElement(
          Workbook,
          { name: "Book" },
          React.createElement(
            Sheet,
            { name: "Sheet1" },
            React.createElement(Sheet, { name: "Nested" }),
          ),
        ),
      ),
    ).rejects.toThrow("Only <Cell> can be nested inside <Sheet>.");
    await expect(
      root.render(
        React.createElement(
          Workbook,
          { name: "Book" },
          React.createElement(Sheet, { name: "Sheet1" }, React.createElement(Cell, { value: 1 })),
        ),
      ),
    ).rejects.toThrow("<Cell> requires an addr prop.");
    await expect(
      root.render(
        React.createElement(
          Workbook,
          { name: "Book" },
          React.createElement(React.StrictMode, null, React.createElement(Sheet, null)),
        ),
      ),
    ).rejects.toThrow("<Sheet> requires a name prop.");

    expect(engine.exportSnapshot().sheets).toEqual([]);
  });

  it("treats false renders as unmounts through the public root", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "renderer-false-unmount" });
    await engine.ready();
    const root = createWorkbookRendererRoot(engine);

    await root.render(
      React.createElement(
        Workbook,
        { name: "Book" },
        React.createElement(
          React.StrictMode,
          null,
          React.createElement(
            Sheet,
            { name: "Sheet1" },
            React.createElement(Cell, { addr: "A1", value: 4 }),
          ),
        ),
      ),
    );
    expect(engine.getCellValue("Sheet1", "A1")).toEqual({ tag: 1, value: 4 });

    await root.render(false);
    expect(engine.exportSnapshot().sheets).toEqual([]);
  });

  it("guards compat updates against invalid roots", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "renderer-compat" });
    await engine.ready();
    const container = createContainer(engine);

    expect(createFiberRoot(container)).toBeTruthy();
    expect(() => updateFiberRoot(null, null, () => {})).toThrow("Invalid fiber root");
  });
});
