export interface ProjectedViewportAxisPatch {
  readonly index: number
  readonly size: number
  readonly hidden: boolean
}

function hasOwnValue<T extends object>(record: T, index: number): boolean {
  return Object.prototype.hasOwnProperty.call(record, index)
}

export function applyProjectedViewportAxisPatches(args: {
  patches: readonly ProjectedViewportAxisPatch[]
  sizes: Record<number, number>
  renderedSizes: Record<number, number>
  pendingSizes: Record<number, number>
  pendingHiddenAxes: Record<number, boolean>
  hiddenAxes: Record<number, true>
}): {
  sizes: Record<number, number>
  renderedSizes: Record<number, number>
  pendingSizes: Record<number, number>
  pendingHiddenAxes: Record<number, boolean>
  hiddenAxes: Record<number, true>
  axisChanged: boolean
} {
  const sizes = { ...args.sizes }
  const renderedSizes = { ...args.renderedSizes }
  const pendingSizes = { ...args.pendingSizes }
  const pendingHiddenAxes = { ...args.pendingHiddenAxes }
  const hiddenAxes = { ...args.hiddenAxes }
  let axisChanged = false

  args.patches.forEach((patch) => {
    const hasPendingHidden = hasOwnValue(pendingHiddenAxes, patch.index)
    if (hasPendingHidden && pendingHiddenAxes[patch.index] !== patch.hidden) {
      return
    }
    if (hasPendingHidden) {
      delete pendingHiddenAxes[patch.index]
    }

    const wasHidden = hiddenAxes[patch.index] === true
    const previousSize = sizes[patch.index]
    const previousRenderedSize = renderedSizes[patch.index]
    sizes[patch.index] = patch.size
    const pending = pendingSizes[patch.index]
    if (patch.hidden) {
      hiddenAxes[patch.index] = true
      renderedSizes[patch.index] = 0
      if (!wasHidden || previousRenderedSize !== 0 || previousSize !== patch.size) {
        axisChanged = true
      }
      return
    }

    delete hiddenAxes[patch.index]
    if (pending !== undefined && pending !== patch.size && !wasHidden) {
      return
    }
    if (pending === patch.size) {
      delete pendingSizes[patch.index]
    }
    renderedSizes[patch.index] = patch.size
    if (wasHidden || previousRenderedSize !== patch.size) {
      axisChanged = true
    }
  })

  return {
    sizes,
    renderedSizes,
    pendingSizes,
    pendingHiddenAxes,
    hiddenAxes,
    axisChanged,
  }
}
