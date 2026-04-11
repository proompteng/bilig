export interface ProjectedViewportLocalAxisState {
  readonly sizes: Record<number, number>;
  readonly renderedSizes: Record<number, number>;
  readonly pendingSizes: Record<number, number>;
  readonly hiddenAxes: Record<number, true>;
}

export interface ProjectedViewportLocalAxisResult {
  readonly sizes: Record<number, number>;
  readonly renderedSizes: Record<number, number>;
  readonly pendingSizes: Record<number, number>;
  readonly hiddenAxes: Record<number, true>;
  readonly changed: boolean;
}

function hasOwnValue<T extends object>(record: T, index: number): boolean {
  return Object.prototype.hasOwnProperty.call(record, index);
}

function cloneWithoutIndex<T extends Record<number, unknown>>(record: T, index: number): T {
  const nextRecord = { ...record };
  delete nextRecord[index];
  return nextRecord;
}

export function setProjectedViewportLocalAxisSize(args: {
  state: ProjectedViewportLocalAxisState;
  index: number;
  size: number;
}): ProjectedViewportLocalAxisResult {
  const { state, index, size } = args;
  const nextRenderedSize = state.hiddenAxes[index] ? 0 : size;
  const changed =
    state.sizes[index] !== size ||
    state.renderedSizes[index] !== nextRenderedSize ||
    state.pendingSizes[index] !== size;
  if (!changed) {
    return { ...state, changed: false };
  }
  return {
    sizes: { ...state.sizes, [index]: size },
    renderedSizes: { ...state.renderedSizes, [index]: nextRenderedSize },
    pendingSizes: { ...state.pendingSizes, [index]: size },
    hiddenAxes: { ...state.hiddenAxes },
    changed: true,
  };
}

export function ackProjectedViewportLocalAxisSize(args: {
  state: ProjectedViewportLocalAxisState;
  index: number;
  size: number;
}): ProjectedViewportLocalAxisResult {
  const { state, index, size } = args;
  if (!hasOwnValue(state.pendingSizes, index) || state.pendingSizes[index] !== size) {
    return { ...state, changed: false };
  }
  return {
    sizes: { ...state.sizes },
    renderedSizes: { ...state.renderedSizes },
    pendingSizes: cloneWithoutIndex(state.pendingSizes, index),
    hiddenAxes: { ...state.hiddenAxes },
    changed: true,
  };
}

export function rollbackProjectedViewportLocalAxisSize(args: {
  state: ProjectedViewportLocalAxisState;
  index: number;
  size: number | undefined;
}): ProjectedViewportLocalAxisResult {
  const { state, index, size } = args;
  const nextSizes =
    size === undefined ? cloneWithoutIndex(state.sizes, index) : { ...state.sizes, [index]: size };
  const nextRenderedSizes =
    size === undefined
      ? cloneWithoutIndex(state.renderedSizes, index)
      : {
          ...state.renderedSizes,
          [index]: state.hiddenAxes[index] ? 0 : size,
        };
  const nextPendingSizes = cloneWithoutIndex(state.pendingSizes, index);
  const changed =
    state.sizes[index] !== nextSizes[index] ||
    state.renderedSizes[index] !== nextRenderedSizes[index] ||
    hasOwnValue(state.pendingSizes, index);
  if (!changed) {
    return { ...state, changed: false };
  }
  return {
    sizes: nextSizes,
    renderedSizes: nextRenderedSizes,
    pendingSizes: nextPendingSizes,
    hiddenAxes: { ...state.hiddenAxes },
    changed: true,
  };
}

export function setProjectedViewportLocalAxisHidden(args: {
  state: ProjectedViewportLocalAxisState;
  index: number;
  hidden: boolean;
  size: number;
}): ProjectedViewportLocalAxisResult {
  const { state, index, hidden, size } = args;
  const currentlyHidden = Boolean(state.hiddenAxes[index]);
  const nextRenderedSize = hidden ? 0 : size;
  const changed =
    currentlyHidden !== hidden ||
    state.sizes[index] !== size ||
    state.renderedSizes[index] !== nextRenderedSize;
  if (!changed) {
    return { ...state, changed: false };
  }

  const nextHiddenAxes: Record<number, true> = hidden
    ? { ...state.hiddenAxes, [index]: true as const }
    : cloneWithoutIndex(state.hiddenAxes, index);
  return {
    sizes: { ...state.sizes, [index]: size },
    renderedSizes: { ...state.renderedSizes, [index]: nextRenderedSize },
    pendingSizes: { ...state.pendingSizes },
    hiddenAxes: nextHiddenAxes,
    changed: true,
  };
}

export function rollbackProjectedViewportLocalAxisHidden(args: {
  state: ProjectedViewportLocalAxisState;
  index: number;
  previous: {
    hidden: boolean;
    size: number | undefined;
  };
}): ProjectedViewportLocalAxisResult {
  const { state, index, previous } = args;
  const nextSizes =
    previous.size === undefined
      ? cloneWithoutIndex(state.sizes, index)
      : { ...state.sizes, [index]: previous.size };
  const nextHiddenAxes: Record<number, true> = previous.hidden
    ? { ...state.hiddenAxes, [index]: true as const }
    : cloneWithoutIndex(state.hiddenAxes, index);
  const nextRenderedSizes =
    previous.size === undefined
      ? cloneWithoutIndex(state.renderedSizes, index)
      : {
          ...state.renderedSizes,
          [index]: previous.hidden ? 0 : previous.size,
        };
  const changed =
    state.sizes[index] !== nextSizes[index] ||
    Boolean(state.hiddenAxes[index]) !== previous.hidden ||
    state.renderedSizes[index] !== nextRenderedSizes[index];
  if (!changed) {
    return { ...state, changed: false };
  }

  return {
    sizes: nextSizes,
    renderedSizes: nextRenderedSizes,
    pendingSizes: { ...state.pendingSizes },
    hiddenAxes: nextHiddenAxes,
    changed: true,
  };
}
