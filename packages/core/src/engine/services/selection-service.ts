import { Effect } from "effect";
import type { SelectionState } from "@bilig/protocol";
import type { EngineRuntimeState } from "../runtime-state.js";

interface SetSelectionOptions {
  anchorAddress?: string | null;
  range?: { startAddress: string; endAddress: string } | null;
  editMode?: SelectionState["editMode"];
}

export interface EngineSelectionService {
  readonly subscribe: (listener: () => void) => Effect.Effect<() => void>;
  readonly getSelectionState: () => Effect.Effect<SelectionState>;
  readonly setSelection: (
    sheetName: string,
    address: string | null,
    options?: SetSelectionOptions,
  ) => Effect.Effect<boolean>;
}

export function createEngineSelectionService(
  state: Pick<EngineRuntimeState, "selectionListeners" | "getSelection" | "setSelection">,
): EngineSelectionService {
  return {
    subscribe(listener) {
      return Effect.sync(() => {
        state.selectionListeners.add(listener);
        return () => {
          state.selectionListeners.delete(listener);
        };
      });
    },
    getSelectionState() {
      return Effect.sync(() => state.getSelection());
    },
    setSelection(sheetName, address, options = {}) {
      return Effect.sync(() => {
        const current = state.getSelection();
        const nextSelection: SelectionState = {
          sheetName,
          address,
          anchorAddress: options.anchorAddress ?? address,
          range: options.range ?? (address ? { startAddress: address, endAddress: address } : null),
          editMode: options.editMode ?? current.editMode,
        };

        if (
          current.sheetName === nextSelection.sheetName &&
          current.address === nextSelection.address &&
          current.anchorAddress === nextSelection.anchorAddress &&
          current.editMode === nextSelection.editMode &&
          current.range?.startAddress === nextSelection.range?.startAddress &&
          current.range?.endAddress === nextSelection.range?.endAddress
        ) {
          return false;
        }

        state.setSelection(nextSelection);
        state.selectionListeners.forEach((listener) => listener());
        return true;
      });
    },
  };
}
