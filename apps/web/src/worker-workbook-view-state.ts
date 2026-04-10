type ViewportStoreLike = {
  getColumnWidths(sheetName: string): Readonly<Record<number, number>>;
  getRowHeights(sheetName: string): Readonly<Record<number, number>>;
  getHiddenColumns(sheetName: string): Readonly<Record<number, true>>;
  getHiddenRows(sheetName: string): Readonly<Record<number, true>>;
};

type WorkerHandleLike = {
  viewportStore: ViewportStoreLike;
};

const EMPTY_AXIS_SIZES: Readonly<Record<number, number>> = Object.freeze({});
const EMPTY_HIDDEN_AXES: Readonly<Record<number, true>> = Object.freeze({});

export function readViewportColumnWidths(
  workerHandle: WorkerHandleLike | null | undefined,
  sheetName: string,
): Readonly<Record<number, number>> {
  return workerHandle?.viewportStore.getColumnWidths(sheetName) ?? EMPTY_AXIS_SIZES;
}

export function readViewportRowHeights(
  workerHandle: WorkerHandleLike | null | undefined,
  sheetName: string,
): Readonly<Record<number, number>> {
  return workerHandle?.viewportStore.getRowHeights(sheetName) ?? EMPTY_AXIS_SIZES;
}

export function readViewportHiddenColumns(
  workerHandle: WorkerHandleLike | null | undefined,
  sheetName: string,
): Readonly<Record<number, true>> {
  return workerHandle?.viewportStore.getHiddenColumns(sheetName) ?? EMPTY_HIDDEN_AXES;
}

export function readViewportHiddenRows(
  workerHandle: WorkerHandleLike | null | undefined,
  sheetName: string,
): Readonly<Record<number, true>> {
  return workerHandle?.viewportStore.getHiddenRows(sheetName) ?? EMPTY_HIDDEN_AXES;
}
