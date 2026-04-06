// @vitest-environment jsdom
import { createRoot } from "react-dom/client";
import { act } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { ValueTag } from "@bilig/protocol";
import type { GridEngineLike } from "@bilig/grid";
import { WorkbookView } from "../../../../packages/grid/src/WorkbookView.js";

vi.mock("../../../../packages/grid/src/FormulaBar.js", () => ({
  FormulaBar: () => <div data-testid="formula-bar" />,
}));

vi.mock("../../../../packages/grid/src/WorkbookGridSurface.js", () => ({
  WorkbookGridSurface: () => <div data-testid="grid-surface" />,
}));

vi.mock("../../../../packages/grid/src/WorkbookSheetTabs.js", () => ({
  WorkbookSheetTabs: () => <div data-testid="sheet-tabs" />,
}));

afterEach(() => {
  document.body.innerHTML = "";
});

describe("workbook layout", () => {
  it("renders the assistant as a docked right-side rail beside the spreadsheet", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;

    const engine: GridEngineLike = {
      getCell: () => ({
        sheetName: "Sheet1",
        address: "A1",
        value: { tag: ValueTag.Empty },
        flags: 0,
        version: 0,
      }),
      getCellStyle: () => undefined,
      subscribeCells: () => () => {},
      workbook: {
        getSheet: () => undefined,
      },
    };

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <WorkbookView
          engine={engine}
          sheetNames={["Sheet1"]}
          sheetName="Sheet1"
          selectedAddr="A1"
          selectedCellSnapshot={{
            sheetName: "Sheet1",
            address: "A1",
            value: { tag: ValueTag.Empty },
            flags: 0,
            version: 0,
          }}
          editorValue=""
          editorSelectionBehavior="select-all"
          resolvedValue=""
          isEditing={false}
          isEditingCell={false}
          onSelectSheet={() => {}}
          onSelect={() => {}}
          onAddressCommit={() => {}}
          onBeginEdit={() => {}}
          onBeginFormulaEdit={() => {}}
          onEditorChange={() => {}}
          onCommitEdit={() => {}}
          onCancelEdit={() => {}}
          onClearCell={() => {}}
          onFillRange={() => {}}
          onCopyRange={() => {}}
          onMoveRange={() => {}}
          onPaste={() => {}}
          sideRail={<div data-testid="assistant-rail">Assistant rail</div>}
        />,
      );
    });

    const sideRail = host.querySelector("[data-testid='workbook-side-rail']");
    const gridSurface = host.querySelector("[data-testid='grid-surface']");
    expect(sideRail).not.toBeNull();
    expect(sideRail?.textContent).toContain("Assistant rail");
    expect(gridSurface instanceof Node).toBe(true);
    expect(sideRail instanceof Node).toBe(true);
    expect(
      gridSurface instanceof Node && sideRail instanceof Node
        ? gridSurface.compareDocumentPosition(sideRail)
        : 0,
    ).toBe(Node.DOCUMENT_POSITION_FOLLOWING);

    await act(async () => {
      root.unmount();
    });
  });
});
