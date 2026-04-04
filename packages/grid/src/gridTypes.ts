export type Item = readonly [number, number];

export interface Rectangle {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface CompactSelectionState {
  readonly ranges: readonly (readonly [number, number])[];
  readonly length: number;
  first(): number | undefined;
  last(): number | undefined;
  hasIndex(index: number): boolean;
}

class CompactSelectionImpl implements CompactSelectionState {
  readonly ranges: readonly (readonly [number, number])[];
  readonly length: number;

  constructor(ranges: readonly (readonly [number, number])[]) {
    this.ranges = ranges;
    this.length = ranges.reduce((total, [start, end]) => total + Math.max(0, end - start), 0);
  }

  first(): number | undefined {
    return this.ranges.length === 0 ? undefined : this.ranges[0]?.[0];
  }

  last(): number | undefined {
    const lastRange = this.ranges.at(-1);
    return lastRange ? lastRange[1] - 1 : undefined;
  }

  hasIndex(index: number): boolean {
    return this.ranges.some(([start, end]) => index >= start && index < end);
  }
}

export const CompactSelection = {
  empty(): CompactSelectionState {
    return new CompactSelectionImpl([]);
  },
  fromSingleSelection(selection: number | readonly [number, number]): CompactSelectionState {
    return typeof selection === "number"
      ? new CompactSelectionImpl([[selection, selection + 1]])
      : new CompactSelectionImpl([[selection[0], selection[1]]]);
  },
};

export interface GridSelectionRange {
  cell: Item;
  range: Rectangle;
  rangeStack: readonly Rectangle[];
}

export interface GridSelection {
  current?: GridSelectionRange | undefined;
  columns: CompactSelectionState;
  rows: CompactSelectionState;
}
