import type { WorkbookAgentUiContext } from "@bilig/contracts";
import type { Viewport } from "@bilig/protocol";
import type { WorkerRuntimeSelection } from "./runtime-session.js";

export interface WorkbookAgentSelectionRange {
  readonly startAddress: string;
  readonly endAddress: string;
}

export function singleCellAgentSelectionRange(
  selection: WorkerRuntimeSelection,
): WorkbookAgentSelectionRange {
  return {
    startAddress: selection.address,
    endAddress: selection.address,
  };
}

export function buildWorkbookAgentContext(input: {
  readonly selection: WorkerRuntimeSelection;
  readonly selectionRange: WorkbookAgentSelectionRange;
  readonly viewport: Viewport;
}): WorkbookAgentUiContext {
  return {
    selection: {
      sheetName: input.selection.sheetName,
      address: input.selection.address,
      range: {
        startAddress: input.selectionRange.startAddress,
        endAddress: input.selectionRange.endAddress,
      },
    },
    viewport: input.viewport,
  };
}
