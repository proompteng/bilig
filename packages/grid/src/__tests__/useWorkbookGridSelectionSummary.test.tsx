// @vitest-environment jsdom
import { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  createGridSelection,
  createRangeSelection,
  createSheetSelection,
} from "../gridSelection.js";
import { useWorkbookGridSelectionSummary } from "../useWorkbookGridSelectionSummary.js";

function Harness(props: {
  readonly selection: Parameters<typeof useWorkbookGridSelectionSummary>[0]["gridSelection"];
  readonly selectedAddr: string;
  readonly onSelectionLabelChange?: ((label: string) => void) | undefined;
  readonly onSelectionRangeChange?:
    | ((range: { startAddress: string; endAddress: string }) => void)
    | undefined;
}) {
  useWorkbookGridSelectionSummary({
    gridSelection: props.selection,
    selectedAddr: props.selectedAddr,
    onSelectionLabelChange: props.onSelectionLabelChange,
    onSelectionRangeChange: props.onSelectionRangeChange,
  });
  return null;
}

beforeEach(() => {
  (
    globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
  ).IS_REACT_ACT_ENVIRONMENT = true;
});

afterEach(() => {
  document.body.innerHTML = "";
  vi.restoreAllMocks();
});

describe("useWorkbookGridSelectionSummary", () => {
  it("reports both the selection label and authoritative range bounds for rectangular ranges", async () => {
    const onSelectionLabelChange = vi.fn();
    const onSelectionRangeChange = vi.fn();
    const selection = createRangeSelection(createGridSelection(1, 1), [1, 1], [3, 4]);
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <Harness
          selection={selection}
          selectedAddr="B2"
          onSelectionLabelChange={onSelectionLabelChange}
          onSelectionRangeChange={onSelectionRangeChange}
        />,
      );
    });

    expect(onSelectionLabelChange).toHaveBeenCalledWith("B2:D5");
    expect(onSelectionRangeChange).toHaveBeenCalledWith({
      startAddress: "B2",
      endAddress: "D5",
    });

    await act(async () => {
      root.unmount();
    });
  });

  it("reports full-sheet bounds without relying on the formatted label", async () => {
    const onSelectionRangeChange = vi.fn();
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <Harness
          selection={createSheetSelection()}
          selectedAddr="A1"
          onSelectionRangeChange={onSelectionRangeChange}
        />,
      );
    });

    expect(onSelectionRangeChange).toHaveBeenCalledWith({
      startAddress: "A1",
      endAddress: "XFD1048576",
    });

    await act(async () => {
      root.unmount();
    });
  });
});
