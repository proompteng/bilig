import React from "react";
import { describe, expect, it } from "vitest";
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
});
