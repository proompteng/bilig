import { useEffect, useMemo } from "react";
import { formatSelectionSummary } from "./gridSelection.js";
import type { GridSelection } from "./gridTypes.js";

export function useWorkbookGridSelectionSummary(input: {
  gridSelection: GridSelection;
  selectedAddr: string;
  onSelectionLabelChange?: ((label: string) => void) | undefined;
}) {
  const { gridSelection, onSelectionLabelChange, selectedAddr } = input;
  const selectionSummary = useMemo(
    () => formatSelectionSummary(gridSelection, selectedAddr),
    [gridSelection, selectedAddr],
  );

  useEffect(() => {
    onSelectionLabelChange?.(selectionSummary);
  }, [onSelectionLabelChange, selectionSummary]);
}
