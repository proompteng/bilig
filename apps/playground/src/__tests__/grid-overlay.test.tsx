// @vitest-environment jsdom
import React, { act, useState } from "react";
import { createRoot } from "react-dom/client";
import { describe, expect, it } from "vitest";
import { SpreadsheetEngine } from "@bilig/core";
import { SheetGridView } from "@bilig/grid";

class ResizeObserverStub {
  private readonly callback: ResizeObserverCallback;

  constructor(callback: ResizeObserverCallback) {
    this.callback = callback;
  }

  observe(target: Element) {
    this.callback(
      [
        {
          contentRect: {
            width: 960,
            height: 640
          }
        } as ResizeObserverEntry
      ],
      this as unknown as ResizeObserver
    );
    return target;
  }

  disconnect() {}

  unobserve() {}
}

describe("grid overlay editing", () => {
  it("opens and closes an in-grid editor overlay", async () => {
    (globalThis as typeof globalThis & {
      IS_REACT_ACT_ENVIRONMENT?: boolean;
      ResizeObserver?: typeof ResizeObserver;
    }).IS_REACT_ACT_ENVIRONMENT = true;
    globalThis.ResizeObserver = ResizeObserverStub as typeof ResizeObserver;
    Element.prototype.scrollTo = () => {};

    const engine = new SpreadsheetEngine({ workbookName: "overlay-test" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 10);
    engine.setCellValue("Sheet1", "B1", 20);

    function Harness() {
      const [selectedAddr, setSelectedAddr] = useState("A1");
      const [editing, setEditing] = useState(false);
      const [editorValue, setEditorValue] = useState("10");
      return (
        <SheetGridView
          editorValue={editorValue}
          engine={engine}
          isEditingCell={editing}
          onBeginEdit={() => setEditing(true)}
          onCancelEdit={() => setEditing(false)}
          onCommitEdit={() => setEditing(false)}
          onEditorChange={setEditorValue}
          onSelect={setSelectedAddr}
          resolvedValue={editorValue}
          selectedAddr={selectedAddr}
          sheetName="Sheet1"
        />
      );
    }

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<Harness />);
    });

    const cell = host.querySelector<HTMLButtonElement>("[data-addr='A1']");
    expect(cell).not.toBeNull();

    await act(async () => {
      cell?.dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));
    });

    expect(host.querySelector("[data-testid='cell-editor-overlay']")).not.toBeNull();

    const input = host.querySelector<HTMLInputElement>("[aria-label='Sheet1!A1 editor']");
    expect(input).not.toBeNull();

    await act(async () => {
      input?.dispatchEvent(new KeyboardEvent("keydown", { bubbles: true, key: "Escape" }));
    });

    expect(host.querySelector("[data-testid='cell-editor-overlay']")).toBeNull();

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });
});
