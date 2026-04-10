// @vitest-environment jsdom
import { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";
import type {
  HeaderSelection,
  VisibleRegionState,
} from "../../../../packages/grid/src/gridPointer.js";
import type { GridSelection } from "../../../../packages/grid/src/gridTypes.js";
import { WorkbookGridContextMenu } from "../../../../packages/grid/src/WorkbookGridContextMenu.js";
import { useWorkbookGridContextMenu } from "../../../../packages/grid/src/useWorkbookGridContextMenu.js";
import type { WorkbookGridContextMenuTarget } from "../../../../packages/grid/src/workbookGridContextMenuTarget.js";

afterEach(() => {
  document.body.innerHTML = "";
});

function ContextMenuHarness(props: {
  headerSelection: HeaderSelection | null;
  hiddenColumns?: Readonly<Record<number, true>> | undefined;
  hiddenRows?: Readonly<Record<number, true>> | undefined;
  isEditingCell?: boolean;
  onCommitEdit?: (() => void) | undefined;
  onHideColumn?: ((columnIndex: number, hidden: boolean) => void) | undefined;
  onHideRow?: ((rowIndex: number, hidden: boolean) => void) | undefined;
  openTarget?: WorkbookGridContextMenuTarget | undefined;
  onSelect?: ((addr: string) => void) | undefined;
  setGridSelection?: ((selection: GridSelection) => void) | undefined;
}) {
  const menu = useWorkbookGridContextMenu({
    focusGrid() {},
    isEditingCell: props.isEditingCell ?? false,
    onCommitEdit: props.onCommitEdit ?? (() => {}),
    onSelect: props.onSelect ?? (() => {}),
    hiddenColumnsByIndex: props.hiddenColumns,
    hiddenRowsByIndex: props.hiddenRows,
    onSetColumnHidden: props.onHideColumn,
    onSetRowHidden: props.onHideRow,
    resolveHeaderSelectionAtPointer() {
      return props.headerSelection;
    },
    selectedCell: [4, 2],
    setGridSelection: props.setGridSelection ?? (() => {}),
    visibleRegion: {
      range: { x: 0, y: 0, width: 20, height: 20 },
      tx: 0,
      ty: 0,
    } satisfies VisibleRegionState,
  });

  return (
    <div data-testid="host" onContextMenuCapture={menu.handleHostContextMenuCapture}>
      <button
        data-testid="open-context-menu"
        type="button"
        onClick={() => {
          if (props.openTarget) {
            menu.openContextMenuForTarget(props.openTarget);
          }
        }}
      >
        Open
      </button>
      {menu.contextMenuState ? (
        <WorkbookGridContextMenu
          menuRef={menu.menuRef}
          onClose={menu.closeContextMenu}
          onToggleTargetHidden={menu.toggleTargetHidden}
          state={menu.contextMenuState}
        />
      ) : null}
    </div>
  );
}

describe("workbook grid context menu", () => {
  it("opens a row context menu and hides the targeted row", async () => {
    const onHideRow = vi.fn();
    const onSelect = vi.fn();
    const setGridSelection = vi.fn();
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <ContextMenuHarness
          headerSelection={{ kind: "row", index: 7 }}
          onHideRow={onHideRow}
          onSelect={onSelect}
          setGridSelection={setGridSelection}
        />,
      );
    });

    const target = host.querySelector("[data-testid='host']");
    await act(async () => {
      target?.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          button: 2,
          clientX: 120,
          clientY: 90,
        }),
      );
    });

    expect(host.querySelector("[data-testid='grid-context-menu']")).not.toBeNull();
    expect(onSelect).toHaveBeenCalledWith("E8");

    const hideButton = host.querySelector("[data-testid='grid-context-action-hide-row']");
    await act(async () => {
      hideButton?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(onHideRow).toHaveBeenCalledWith(7, true);
    expect(setGridSelection).toHaveBeenCalledTimes(1);
    expect(host.querySelector("[data-testid='grid-context-menu']")).toBeNull();

    await act(async () => {
      root.unmount();
    });
  });

  it("shows unhide for a hidden row and restores the row", async () => {
    const onHideRow = vi.fn();
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <ContextMenuHarness
          headerSelection={{ kind: "row", index: 7 }}
          hiddenRows={{ 7: true }}
          onHideRow={onHideRow}
        />,
      );
    });

    const target = host.querySelector("[data-testid='host']");
    await act(async () => {
      target?.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          button: 2,
          clientX: 120,
          clientY: 90,
        }),
      );
    });

    const unhideButton = host.querySelector("[data-testid='grid-context-action-unhide-row']");
    expect(unhideButton).not.toBeNull();

    await act(async () => {
      unhideButton?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(onHideRow).toHaveBeenCalledWith(7, false);
    expect(host.querySelector("[data-testid='grid-context-menu']")).toBeNull();

    await act(async () => {
      root.unmount();
    });
  });

  it("commits the active edit before opening a column context menu", async () => {
    const onCommitEdit = vi.fn();
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <ContextMenuHarness
          headerSelection={{ kind: "column", index: 3 }}
          isEditingCell
          onCommitEdit={onCommitEdit}
          onHideColumn={vi.fn()}
        />,
      );
    });

    const target = host.querySelector("[data-testid='host']");
    await act(async () => {
      target?.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          button: 2,
          clientX: 64,
          clientY: 24,
        }),
      );
    });

    expect(onCommitEdit).toHaveBeenCalledTimes(1);
    expect(host.querySelector("[data-testid='grid-context-action-hide-column']")).not.toBeNull();

    await act(async () => {
      root.unmount();
    });
  });

  it("closes the context menu on escape", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <ContextMenuHarness headerSelection={{ kind: "row", index: 1 }} onHideRow={vi.fn()} />,
      );
    });

    const target = host.querySelector("[data-testid='host']");
    await act(async () => {
      target?.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          button: 2,
          clientX: 32,
          clientY: 60,
        }),
      );
    });

    expect(host.querySelector("[data-testid='grid-context-menu']")).not.toBeNull();

    await act(async () => {
      window.dispatchEvent(new KeyboardEvent("keydown", { bubbles: true, key: "Escape" }));
    });

    expect(host.querySelector("[data-testid='grid-context-menu']")).toBeNull();

    await act(async () => {
      root.unmount();
    });
  });

  it("opens a context menu programmatically for keyboard-triggered header actions", async () => {
    const onHideColumn = vi.fn();
    const onSelect = vi.fn();
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <ContextMenuHarness
          headerSelection={null}
          onHideColumn={onHideColumn}
          onSelect={onSelect}
          openTarget={{ target: { kind: "column", index: 2 }, x: 180, y: 24 }}
        />,
      );
    });

    const openButton = host.querySelector("[data-testid='open-context-menu']");
    await act(async () => {
      openButton?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(host.querySelector("[data-testid='grid-context-menu']")).not.toBeNull();
    expect(onSelect).toHaveBeenCalledWith("C3");

    const hideButton = host.querySelector("[data-testid='grid-context-action-hide-column']");
    await act(async () => {
      hideButton?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(onHideColumn).toHaveBeenCalledWith(2, true);

    await act(async () => {
      root.unmount();
    });
  });
});
