// @vitest-environment jsdom
import { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";
import { WorkbookSheetTabs } from "../../../../packages/grid/src/WorkbookSheetTabs.js";

afterEach(() => {
  document.body.innerHTML = "";
});

describe("WorkbookSheetTabs", () => {
  it("renders a tablist and lets the user add and switch sheets", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;

    const onCreateSheet = vi.fn();
    const onSelectSheet = vi.fn();
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <WorkbookSheetTabs
          onCreateSheet={onCreateSheet}
          onSelectSheet={onSelectSheet}
          sheetName="Sheet1"
          sheetNames={["Sheet1", "Sheet2", "Sheet3"]}
        />,
      );
    });

    expect(host.querySelector("[role='tablist']")?.getAttribute("aria-label")).toBe("Sheets");
    expect(
      host
        .querySelector("[data-testid='workbook-sheet-tab-Sheet1']")
        ?.getAttribute("aria-selected"),
    ).toBe("true");

    await act(async () => {
      host
        .querySelector("[data-testid='workbook-sheet-tab-Sheet2']")
        ?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(onSelectSheet).toHaveBeenCalledWith("Sheet2");

    await act(async () => {
      host
        .querySelector("[data-testid='workbook-sheet-add']")
        ?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(onCreateSheet).toHaveBeenCalledTimes(1);

    await act(async () => {
      root.unmount();
    });
  });

  it("supports keyboard sheet switching and rename entry", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;

    const onRenameSheet = vi.fn();
    const onSelectSheet = vi.fn();
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <WorkbookSheetTabs
          onRenameSheet={onRenameSheet}
          onSelectSheet={onSelectSheet}
          sheetName="Sheet2"
          sheetNames={["Sheet1", "Sheet2", "Sheet3"]}
        />,
      );
    });

    const activeTab = host.querySelector("[data-testid='workbook-sheet-tab-Sheet2']");
    expect(activeTab instanceof HTMLButtonElement).toBe(true);

    await act(async () => {
      activeTab?.dispatchEvent(
        new KeyboardEvent("keydown", {
          key: "ArrowRight",
          bubbles: true,
        }),
      );
    });

    expect(onSelectSheet).toHaveBeenCalledWith("Sheet3");

    await act(async () => {
      activeTab?.dispatchEvent(
        new KeyboardEvent("keydown", {
          key: "F2",
          bubbles: true,
        }),
      );
    });

    const renameInput = host.querySelector("[data-testid='workbook-sheet-rename-input']");
    expect(renameInput instanceof HTMLInputElement).toBe(true);

    await act(async () => {
      if (!(renameInput instanceof HTMLInputElement)) {
        throw new Error("rename input not found");
      }
      const descriptor = Object.getOwnPropertyDescriptor(
        window.HTMLInputElement.prototype,
        "value",
      );
      descriptor?.set?.call(renameInput, "Revenue");
      renameInput.dispatchEvent(new Event("input", { bubbles: true }));
    });

    await act(async () => {
      if (!(renameInput instanceof HTMLInputElement)) {
        throw new Error("rename input not found");
      }
      renameInput.dispatchEvent(
        new KeyboardEvent("keydown", {
          key: "Enter",
          bubbles: true,
        }),
      );
    });

    expect(onRenameSheet).toHaveBeenCalledWith("Sheet2", "Revenue");

    await act(async () => {
      root.unmount();
    });
  });

  it("opens a context menu on right click and deletes the requested sheet", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;

    const onDeleteSheet = vi.fn();
    const onRenameSheet = vi.fn();
    const onSelectSheet = vi.fn();
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <WorkbookSheetTabs
          onDeleteSheet={onDeleteSheet}
          onRenameSheet={onRenameSheet}
          onSelectSheet={onSelectSheet}
          sheetName="Sheet1"
          sheetNames={["Sheet1", "Sheet2", "Sheet3"]}
        />,
      );
    });

    const targetTab = host.querySelector("[data-testid='workbook-sheet-tab-Sheet2']");
    expect(targetTab).not.toBeNull();

    await act(async () => {
      targetTab?.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          button: 2,
          clientX: 48,
          clientY: 24,
        }),
      );
    });

    expect(onSelectSheet).not.toHaveBeenCalled();

    const deleteItem = document.body.querySelector("[data-testid='workbook-sheet-menu-delete']");
    expect(deleteItem).not.toBeNull();

    await act(async () => {
      deleteItem?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(onDeleteSheet).toHaveBeenCalledWith("Sheet2");
    expect(onRenameSheet).not.toHaveBeenCalled();

    await act(async () => {
      root.unmount();
    });
  });
});
