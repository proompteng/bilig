/* oxlint-disable @typescript-eslint/no-unsafe-type-assertion */
import { describe, expect, it } from "vitest";
import type { TypedView, Zero } from "@rocicorp/zero";
import { TileSubscriptionManager } from "./tile-subscriptions.js";

function createTypedView<T>(initial: T) {
  let data = initial;
  const listeners = new Set<(value: T) => void>();
  return {
    get data() {
      return data;
    },
    set data(next: T) {
      data = next;
    },
    addListener(listener: (value: T) => void) {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
    destroy() {},
    emit(next: T) {
      data = next;
      for (const listener of listeners) {
        listener(next);
      }
    },
  } as TypedView<T> & { emit(next: T): void };
}

describe("TileSubscriptionManager", () => {
  it("keeps viewport tile data in maps and reuses cached snapshots until a tile changes", () => {
    const sourceView = createTypedView([
      {
        workbookId: "bilig-demo",
        sheetName: "Sheet1",
        address: "A1",
        inputValue: "hello",
      },
    ]);
    const evalView = createTypedView([
      {
        workbookId: "bilig-demo",
        sheetName: "Sheet1",
        address: "A1",
        value: { tag: 0 },
        flags: 0,
        version: 1,
      },
    ]);
    const rowView = createTypedView([
      {
        workbookId: "bilig-demo",
        sheetName: "Sheet1",
        startIndex: 0,
        count: 1,
        size: 22,
      },
    ]);
    const columnView = createTypedView([
      {
        workbookId: "bilig-demo",
        sheetName: "Sheet1",
        startIndex: 0,
        count: 1,
        size: 104,
      },
    ]);
    const styleView = createTypedView([]);
    const formatView = createTypedView([]);
    const views = [sourceView, evalView, rowView, columnView, styleView, formatView] as const;

    let index = 0;
    const zero = {
      materialize() {
        const view = views[index];
        if (!view) {
          throw new Error("No more views available");
        }
        index += 1;
        return view;
      },
    } as unknown as Zero;

    let notifications = 0;
    const manager = new TileSubscriptionManager(zero, "bilig-demo", () => {});
    const attachment = manager.subscribeViewport(
      "Sheet1",
      {
        rowStart: 0,
        rowEnd: 0,
        colStart: 0,
        colEnd: 0,
      },
      () => {
        notifications += 1;
      },
    );

    const initial = attachment.getData();
    expect(initial.sourceCells.get("A1")?.inputValue).toBe("hello");
    expect(initial.cellEval.get("A1")?.version).toBe(1);
    expect(initial.rowMetadata.size).toBe(1);
    expect(initial.columnMetadata.size).toBe(1);
    expect(attachment.getData()).toBe(initial);

    sourceView.emit([
      {
        workbookId: "bilig-demo",
        sheetName: "Sheet1",
        address: "A1",
        inputValue: "updated",
      },
    ]);

    const next = attachment.getData();
    expect(next).toBe(initial);
    expect(next.sourceCells.get("A1")?.inputValue).toBe("updated");
    expect(attachment.getSourceCell("A1")?.inputValue).toBe("updated");
    expect(notifications).toBe(1);

    attachment.dispose();
    manager.dispose();
  });
});
