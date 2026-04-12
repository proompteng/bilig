// @vitest-environment jsdom
import { act, type MutableRefObject } from "react";
import { createRoot } from "react-dom/client";
import { ValueTag, type CellRangeRef } from "@bilig/protocol";
import { afterEach, describe, expect, it, vi } from "vitest";
import { getWorkbookShortcutLabel } from "../shortcut-registry.js";
import { WorkbookToolbar } from "../workbook-toolbar.js";
import { useWorkbookToolbar } from "../use-workbook-toolbar.js";

function ToolbarHookHarness(props: {
  readonly invokeMutation: (method: string, ...args: unknown[]) => Promise<void>;
  readonly selectionRangeRef: MutableRefObject<CellRangeRef>;
}) {
  const { ribbon } = useWorkbookToolbar({
    canHideCurrentColumn: false,
    canHideCurrentRow: false,
    canRedo: false,
    canUndo: false,
    canUnhideCurrentColumn: false,
    canUnhideCurrentRow: false,
    connectionStateName: "connected",
    currentFillColor: "#ffffff",
    currentNumberFormatKind: "general",
    currentTextColor: "#111827",
    horizontalAlignment: null,
    invokeMutation: props.invokeMutation,
    localPersistenceMode: "persistent",
    onApplyBorderPreset: () => {},
    onClearStyle: () => {},
    onFillColorReset: () => {},
    onFillColorSelect: () => {},
    onFontSizeChange: () => {},
    onHideCurrentColumn: () => {},
    onHideCurrentRow: () => {},
    onHorizontalAlignmentChange: () => {},
    onNumberFormatChange: () => {},
    onRedo: () => {},
    onTextColorReset: () => {},
    onTextColorSelect: () => {},
    onToggleBold: () => {},
    onToggleItalic: () => {},
    onToggleUnderline: () => {},
    onToggleWrap: () => {},
    onUndo: () => {},
    onUnhideCurrentColumn: () => {},
    onUnhideCurrentRow: () => {},
    remoteSyncAvailable: true,
    runtimeReady: true,
    selectedCell: {
      address: "A1",
      sheetName: "Sheet1",
      flags: 0,
      value: { tag: ValueTag.Empty },
      version: 0,
    },
    selectedStyle: undefined,
    selection: { sheetName: "Sheet1" },
    selectionRangeRef: props.selectionRangeRef,
    trailingContent: null,
    writesAllowed: true,
    zeroConfigured: true,
    zeroHealthReady: true,
  });

  return <>{ribbon}</>;
}

afterEach(() => {
  document.body.innerHTML = "";
});

describe("WorkbookToolbar", () => {
  it("opens the structure menu and invokes current row actions", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;

    const onHideCurrentRow = vi.fn();
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <WorkbookToolbar
          canHideCurrentColumn={false}
          canHideCurrentRow
          canRedo={false}
          canUndo={false}
          canUnhideCurrentColumn={false}
          canUnhideCurrentRow={false}
          currentFillColor="#ffffff"
          currentNumberFormatKind="general"
          currentTextColor="#111827"
          horizontalAlignment={null}
          isBoldActive={false}
          isItalicActive={false}
          isUnderlineActive={false}
          isWrapActive={false}
          onApplyBorderPreset={() => {}}
          onClearStyle={() => {}}
          onFillColorReset={() => {}}
          onFillColorSelect={() => {}}
          onFontSizeChange={() => {}}
          onHideCurrentColumn={() => {}}
          onHideCurrentRow={onHideCurrentRow}
          onHorizontalAlignmentChange={() => {}}
          onNumberFormatChange={() => {}}
          onRedo={() => {}}
          onTextColorReset={() => {}}
          onTextColorSelect={() => {}}
          onToggleBold={() => {}}
          onToggleItalic={() => {}}
          onToggleUnderline={() => {}}
          onToggleWrap={() => {}}
          onUndo={() => {}}
          onUnhideCurrentColumn={() => {}}
          onUnhideCurrentRow={() => {}}
          recentFillColors={[]}
          recentTextColors={[]}
          selectedFontSize="11"
          writesAllowed
        />,
      );
    });

    const trigger = document.querySelector("[aria-label='Structure']");
    expect(trigger).not.toBeNull();

    await act(async () => {
      trigger?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    const hideRowButton = document.querySelector("[aria-label='Hide row']");
    expect(hideRowButton).not.toBeNull();
    expect(hideRowButton?.getAttribute("disabled")).toBeNull();

    await act(async () => {
      hideRowButton?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(onHideCurrentRow).toHaveBeenCalledTimes(1);

    await act(async () => {
      root.unmount();
    });
  });

  it("shows shared shortcut labels on alignment buttons", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <WorkbookToolbar
          canHideCurrentColumn={false}
          canHideCurrentRow={false}
          canRedo={false}
          canUndo={false}
          canUnhideCurrentColumn={false}
          canUnhideCurrentRow={false}
          currentFillColor="#ffffff"
          currentNumberFormatKind="general"
          currentTextColor="#111827"
          horizontalAlignment="left"
          isBoldActive={false}
          isItalicActive={false}
          isUnderlineActive={false}
          isWrapActive={false}
          onApplyBorderPreset={() => {}}
          onClearStyle={() => {}}
          onFillColorReset={() => {}}
          onFillColorSelect={() => {}}
          onFontSizeChange={() => {}}
          onHideCurrentColumn={() => {}}
          onHideCurrentRow={() => {}}
          onHorizontalAlignmentChange={() => {}}
          onNumberFormatChange={() => {}}
          onRedo={() => {}}
          onTextColorReset={() => {}}
          onTextColorSelect={() => {}}
          onToggleBold={() => {}}
          onToggleItalic={() => {}}
          onToggleUnderline={() => {}}
          onToggleWrap={() => {}}
          onUndo={() => {}}
          onUnhideCurrentColumn={() => {}}
          onUnhideCurrentRow={() => {}}
          recentFillColors={[]}
          recentTextColors={[]}
          selectedFontSize="11"
          writesAllowed
        />,
      );
    });

    const alignLeftButton = document.querySelector("[aria-label='Align left']");
    const alignCenterButton = document.querySelector("[aria-label='Align center']");
    const alignRightButton = document.querySelector("[aria-label='Align right']");

    expect(alignLeftButton?.getAttribute("title")).toBe(
      `Align left (${getWorkbookShortcutLabel("align-left")})`,
    );
    expect(alignCenterButton?.getAttribute("title")).toBe(
      `Align center (${getWorkbookShortcutLabel("align-center")})`,
    );
    expect(alignRightButton?.getAttribute("title")).toBe(
      `Align right (${getWorkbookShortcutLabel("align-right")})`,
    );

    await act(async () => {
      root.unmount();
    });
  });

  it("keeps toolbar controls on one shared height system", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <WorkbookToolbar
          canHideCurrentColumn={false}
          canHideCurrentRow
          canRedo
          canUndo
          canUnhideCurrentColumn={false}
          canUnhideCurrentRow={false}
          currentFillColor="#ffffff"
          currentNumberFormatKind="general"
          currentTextColor="#111827"
          horizontalAlignment="left"
          isBoldActive={false}
          isItalicActive={false}
          isUnderlineActive={false}
          isWrapActive={false}
          onApplyBorderPreset={() => {}}
          onClearStyle={() => {}}
          onFillColorReset={() => {}}
          onFillColorSelect={() => {}}
          onFontSizeChange={() => {}}
          onHideCurrentColumn={() => {}}
          onHideCurrentRow={() => {}}
          onHorizontalAlignmentChange={() => {}}
          onNumberFormatChange={() => {}}
          onRedo={() => {}}
          onTextColorReset={() => {}}
          onTextColorSelect={() => {}}
          onToggleBold={() => {}}
          onToggleItalic={() => {}}
          onToggleUnderline={() => {}}
          onToggleWrap={() => {}}
          onUndo={() => {}}
          onUnhideCurrentColumn={() => {}}
          onUnhideCurrentRow={() => {}}
          recentFillColors={[]}
          recentTextColors={[]}
          selectedFontSize="11"
          writesAllowed
        />,
      );
    });

    const historyGroup = host.querySelector("[aria-label='History']");
    const undoButton = host.querySelector("[aria-label='Undo']");
    const numberFormatTrigger = host.querySelector("[aria-label='Number format']");
    const fontSizeTrigger = host.querySelector("[aria-label='Font size']");
    const structureTrigger = host.querySelector("[aria-label='Structure']");

    expect(historyGroup?.className).toContain("h-8");
    expect(undoButton?.className).toContain("h-8");
    expect(numberFormatTrigger?.className).toContain("h-8");
    expect(fontSizeTrigger?.className).toContain("h-8");
    expect(structureTrigger?.className).toContain("h-8");

    await act(async () => {
      root.unmount();
    });
  });

  it("renders trailing controls in the toolbar's right slot", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <WorkbookToolbar
          canHideCurrentColumn={false}
          canHideCurrentRow={false}
          canRedo={false}
          canUndo={false}
          canUnhideCurrentColumn={false}
          canUnhideCurrentRow={false}
          currentFillColor="#ffffff"
          currentNumberFormatKind="general"
          currentTextColor="#111827"
          horizontalAlignment={null}
          isBoldActive={false}
          isItalicActive={false}
          isUnderlineActive={false}
          isWrapActive={false}
          onApplyBorderPreset={() => {}}
          onClearStyle={() => {}}
          onFillColorReset={() => {}}
          onFillColorSelect={() => {}}
          onFontSizeChange={() => {}}
          onHideCurrentColumn={() => {}}
          onHideCurrentRow={() => {}}
          onHorizontalAlignmentChange={() => {}}
          onNumberFormatChange={() => {}}
          onRedo={() => {}}
          onTextColorReset={() => {}}
          onTextColorSelect={() => {}}
          onToggleBold={() => {}}
          onToggleItalic={() => {}}
          onToggleUnderline={() => {}}
          onToggleWrap={() => {}}
          onUndo={() => {}}
          onUnhideCurrentColumn={() => {}}
          onUnhideCurrentRow={() => {}}
          recentFillColors={[]}
          recentTextColors={[]}
          selectedFontSize="11"
          trailingContent={<div data-testid="toolbar-tail-probe">Tail</div>}
          writesAllowed
        />,
      );
    });

    const trailingSlot = host.querySelector("[data-testid='toolbar-trailing-content']");
    expect(trailingSlot).not.toBeNull();
    expect(trailingSlot?.className).toContain("ml-auto");
    expect(trailingSlot?.textContent).toContain("Tail");

    await act(async () => {
      root.unmount();
    });
  });

  it("targets the live selection range for formatting actions without waiting for a rerender", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;

    const invokeMutation = vi.fn(async () => {});
    const selectionRangeRef: MutableRefObject<CellRangeRef> = {
      current: {
        sheetName: "Sheet1",
        startAddress: "A1",
        endAddress: "A1",
      },
    };
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <ToolbarHookHarness
          invokeMutation={invokeMutation}
          selectionRangeRef={selectionRangeRef}
        />,
      );
    });

    selectionRangeRef.current = {
      sheetName: "Sheet1",
      startAddress: "B2",
      endAddress: "D5",
    };

    await act(async () => {
      host
        .querySelector("[aria-label='Bold']")
        ?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(invokeMutation).toHaveBeenCalledWith(
      "setRangeStyle",
      {
        sheetName: "Sheet1",
        startAddress: "B2",
        endAddress: "D5",
      },
      {
        font: { bold: true },
      },
    );

    await act(async () => {
      root.unmount();
    });
  });

  it("hides native toolbar overflow scrollbars while preserving horizontal scrolling", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <WorkbookToolbar
          canHideCurrentColumn={false}
          canHideCurrentRow={false}
          canRedo={false}
          canUndo={false}
          canUnhideCurrentColumn={false}
          canUnhideCurrentRow={false}
          currentFillColor="#ffffff"
          currentNumberFormatKind="general"
          currentTextColor="#111827"
          horizontalAlignment={null}
          isBoldActive={false}
          isItalicActive={false}
          isUnderlineActive={false}
          isWrapActive={false}
          onApplyBorderPreset={() => {}}
          onClearStyle={() => {}}
          onFillColorReset={() => {}}
          onFillColorSelect={() => {}}
          onFontSizeChange={() => {}}
          onHideCurrentColumn={() => {}}
          onHideCurrentRow={() => {}}
          onHorizontalAlignmentChange={() => {}}
          onNumberFormatChange={() => {}}
          onRedo={() => {}}
          onTextColorReset={() => {}}
          onTextColorSelect={() => {}}
          onToggleBold={() => {}}
          onToggleItalic={() => {}}
          onToggleUnderline={() => {}}
          onToggleWrap={() => {}}
          onUndo={() => {}}
          onUnhideCurrentColumn={() => {}}
          onUnhideCurrentRow={() => {}}
          recentFillColors={[]}
          recentTextColors={[]}
          selectedFontSize="11"
          trailingContent={<div>Trailing</div>}
          writesAllowed
        />,
      );
    });

    const toolbar = host.querySelector("[aria-label='Formatting toolbar']");
    expect(toolbar?.className).toContain("overflow-x-auto");
    expect(toolbar?.className).toContain("overflow-y-hidden");
    expect(toolbar?.className).toContain("wb-scrollbar-none");

    await act(async () => {
      root.unmount();
    });
  });
});
