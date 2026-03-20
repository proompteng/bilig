// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { describe, expect, it } from "vitest";
import { SpreadsheetEngine } from "@bilig/core";
import { useCell, useSelection } from "@bilig/grid";
import { ValueTag } from "@bilig/protocol";

describe("playground cell subscriptions", () => {
  it("rerenders a watched cell when its own value and format change", async () => {
    (globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
    const engine = new SpreadsheetEngine({ workbookName: "subscription-test" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 1);

    const Probe = React.memo(function Probe() {
      const snapshot = useCell(engine, "Sheet1", "A1");
      const value = snapshot.value.tag === ValueTag.Number ? String(snapshot.value.value) : snapshot.value.tag;
      return <div data-testid="watched-cell">{`${value}|${snapshot.format ?? ""}`}</div>;
    });

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<Probe />);
    });

    expect(host.textContent).toBe("1|");

    await act(async () => {
      engine.setCellValue("Sheet1", "A1", 9);
    });

    expect(host.textContent).toBe("9|");

    await act(async () => {
      engine.setCellFormat("Sheet1", "A1", "currency");
    });

    expect(host.textContent).toBe("9|currency");

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("rerenders a watched empty cell when it materializes later", async () => {
    (globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
    const engine = new SpreadsheetEngine({ workbookName: "materialize-test" });
    await engine.ready();
    engine.createSheet("Sheet1");

    const renders: string[] = [];

    function Probe() {
      const snapshot = useCell(engine, "Sheet1", "D4");
      renders.push(String(snapshot.value.tag));
      return null;
    }

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<Probe />);
    });

    await act(async () => {
      engine.setCellValue("Sheet1", "D4", 42);
    });

    expect(renders).toHaveLength(2);

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("rerenders selection subscribers when the engine selection changes", async () => {
    (globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
    const engine = new SpreadsheetEngine({ workbookName: "selection-test" });
    await engine.ready();

    const seen: string[] = [];

    function Probe() {
      const selection = useSelection(engine);
      seen.push(`${selection.sheetName}!${selection.address ?? "null"}`);
      return null;
    }

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<Probe />);
    });

    await act(async () => {
      engine.setSelection("Sheet9", "C7");
    });

    expect(seen).toEqual(["Sheet1!A1", "Sheet9!C7"]);

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });
});
