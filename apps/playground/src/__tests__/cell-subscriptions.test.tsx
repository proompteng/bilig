// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { describe, expect, it } from "vitest";
import { SpreadsheetEngine } from "@bilig/core";
import { useCell, useSelection } from "@bilig/grid";

describe("playground cell subscriptions", () => {
  it("rerenders only the subscribed cell when an unrelated cell changes", async () => {
    (globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
    const engine = new SpreadsheetEngine({ workbookName: "subscription-test" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 1);
    engine.setCellValue("Sheet1", "B1", 2);

    const renders = new Map<string, number>();

    function Probe({ addr }: { addr: string }) {
      useCell(engine, "Sheet1", addr);
      renders.set(addr, (renders.get(addr) ?? 0) + 1);
      return null;
    }

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <>
          <Probe addr="A1" />
          <Probe addr="B1" />
        </>
      );
    });

    expect(renders.get("A1")).toBe(1);
    expect(renders.get("B1")).toBe(1);

    await act(async () => {
      engine.setCellValue("Sheet1", "A1", 9);
    });

    expect(renders.get("A1")).toBe(2);
    expect(renders.get("B1")).toBe(1);

    await act(async () => {
      engine.setCellValue("Sheet1", "C1", 4);
    });

    expect(renders.get("A1")).toBe(2);
    expect(renders.get("B1")).toBe(1);

    await act(async () => {
      engine.setCellFormat("Sheet1", "A1", "currency");
    });

    expect(renders.get("A1")).toBe(3);
    expect(renders.get("B1")).toBe(1);

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
