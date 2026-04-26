import { createRectangleSelectionFromRange, rectangleToAddresses } from './gridSelection.js'
import type { GridHoverState } from './gridHover.js'
import { resolveMovedRange, sameRectangle } from './gridRangeMove.js'
import type { Item, Rectangle } from './gridTypes.js'

const RANGE_MOVE_AUTOSCROLL_EDGE_PX = 36
const RANGE_MOVE_AUTOSCROLL_MAX_STEP_PX = 32

interface PointerEventLike {
  clientX: number
  clientY: number
}

interface PointerListenerTarget {
  addEventListener(type: 'pointermove' | 'pointerup', listener: (event: PointerEventLike) => void, useCapture: boolean): void
  removeEventListener(type: 'pointermove' | 'pointerup', listener: (event: PointerEventLike) => void, useCapture: boolean): void
}

interface RangeMoveScrollViewport {
  readonly clientHeight: number
  readonly clientWidth: number
  readonly scrollHeight: number
  readonly scrollWidth: number
  scrollLeft: number
  scrollTop: number
  dispatchEvent(event: Event): boolean
  getBoundingClientRect(): Pick<DOMRect, 'bottom' | 'left' | 'right' | 'top'>
}

type RequestRangeMoveFrame = (callback: FrameRequestCallback) => number
type CancelRangeMoveFrame = (handle: number) => void

export function beginWorkbookGridRangeMove(input: {
  cleanupRef: { current: (() => void) | null }
  listenerTarget: PointerListenerTarget
  sourceRange: Rectangle
  pointerCell: Item
  resolvePointerCell: (clientX: number, clientY: number) => Item | null
  setGridSelection: (selection: ReturnType<typeof createRectangleSelectionFromRange>) => void
  onSelectionChange: (selection: ReturnType<typeof createRectangleSelectionFromRange>) => void
  onMoveRange: (sourceStartAddress: string, sourceEndAddress: string, targetStartAddress: string, targetEndAddress: string) => void
  refreshHoverState: (clientX: number, clientY: number, buttons: number) => void
  setIsRangeMoveDragging: (isDragging: boolean) => void
  setHoverState: (state: GridHoverState) => void
  scrollViewport?: RangeMoveScrollViewport | null | undefined
  requestAnimationFrame?: RequestRangeMoveFrame | undefined
  cancelAnimationFrame?: CancelRangeMoveFrame | undefined
}): void {
  const {
    cancelAnimationFrame = resolveDefaultCancelAnimationFrame(),
    cleanupRef,
    listenerTarget,
    onMoveRange,
    pointerCell,
    requestAnimationFrame = resolveDefaultRequestAnimationFrame(),
    refreshHoverState,
    resolvePointerCell,
    scrollViewport = null,
    setGridSelection,
    onSelectionChange,
    setHoverState,
    setIsRangeMoveDragging,
    sourceRange,
  } = input
  const anchorOffset: Item = [pointerCell[0] - sourceRange.x, pointerCell[1] - sourceRange.y]
  let previewRange = sourceRange
  let lastPointerEvent: PointerEventLike | null = null
  let autoScrollFrame: number | null = null

  cleanupRef.current?.()
  setIsRangeMoveDragging(true)
  setHoverState({ cell: null, header: null, cursor: 'grabbing' })

  const updatePreview = (nativeEvent: PointerEventLike) => {
    const nextPointerCell = resolvePointerCell(nativeEvent.clientX, nativeEvent.clientY)
    if (!nextPointerCell) {
      return
    }
    const nextRange = resolveMovedRange(sourceRange, nextPointerCell, anchorOffset)
    if (sameRectangle(previewRange, nextRange)) {
      return
    }
    previewRange = nextRange
    setGridSelection(createRectangleSelectionFromRange(nextRange))
  }

  const cancelAutoScroll = () => {
    if (autoScrollFrame === null) {
      return
    }
    cancelAnimationFrame?.(autoScrollFrame)
    autoScrollFrame = null
  }

  const scheduleAutoScroll = () => {
    if (!scrollViewport || !requestAnimationFrame || autoScrollFrame !== null) {
      return
    }
    autoScrollFrame = requestAnimationFrame(() => {
      autoScrollFrame = null
      if (!lastPointerEvent || !applyRangeMoveAutoScroll(scrollViewport, lastPointerEvent)) {
        return
      }
      updatePreview(lastPointerEvent)
      scheduleAutoScroll()
    })
  }

  const move = (nativeEvent: PointerEventLike) => {
    lastPointerEvent = nativeEvent
    updatePreview(nativeEvent)
    scheduleAutoScroll()
  }

  const cleanup = (nativeEvent?: PointerEventLike) => {
    cancelAutoScroll()
    listenerTarget.removeEventListener('pointermove', move, true)
    listenerTarget.removeEventListener('pointerup', up, true)
    cleanupRef.current = null
    setIsRangeMoveDragging(false)
    if (nativeEvent) {
      refreshHoverState(nativeEvent.clientX, nativeEvent.clientY, 0)
    }
  }

  const up = (nativeEvent: PointerEventLike) => {
    lastPointerEvent = nativeEvent
    updatePreview(nativeEvent)
    cleanup(nativeEvent)
    const nextSelection = createRectangleSelectionFromRange(previewRange)
    setGridSelection(nextSelection)
    onSelectionChange(nextSelection)
    if (sameRectangle(sourceRange, previewRange)) {
      return
    }
    const sourceAddresses = rectangleToAddresses(sourceRange)
    const targetAddresses = rectangleToAddresses(previewRange)
    onMoveRange(sourceAddresses.startAddress, sourceAddresses.endAddress, targetAddresses.startAddress, targetAddresses.endAddress)
  }

  cleanupRef.current = () => {
    cleanup()
  }
  listenerTarget.addEventListener('pointermove', move, true)
  listenerTarget.addEventListener('pointerup', up, true)
}

function resolveDefaultRequestAnimationFrame(): RequestRangeMoveFrame | undefined {
  return typeof window === 'undefined' ? undefined : window.requestAnimationFrame.bind(window)
}

function resolveDefaultCancelAnimationFrame(): CancelRangeMoveFrame | undefined {
  return typeof window === 'undefined' ? undefined : window.cancelAnimationFrame.bind(window)
}

function applyRangeMoveAutoScroll(scrollViewport: RangeMoveScrollViewport, pointer: PointerEventLike): boolean {
  const rect = scrollViewport.getBoundingClientRect()
  const deltaX = resolveRangeMoveAutoScrollStep(pointer.clientX, rect.left, rect.right)
  const deltaY = resolveRangeMoveAutoScrollStep(pointer.clientY, rect.top, rect.bottom)
  if (deltaX === 0 && deltaY === 0) {
    return false
  }

  const nextScrollLeft = clamp(scrollViewport.scrollLeft + deltaX, 0, Math.max(0, scrollViewport.scrollWidth - scrollViewport.clientWidth))
  const nextScrollTop = clamp(scrollViewport.scrollTop + deltaY, 0, Math.max(0, scrollViewport.scrollHeight - scrollViewport.clientHeight))
  if (nextScrollLeft === scrollViewport.scrollLeft && nextScrollTop === scrollViewport.scrollTop) {
    return false
  }

  scrollViewport.scrollLeft = nextScrollLeft
  scrollViewport.scrollTop = nextScrollTop
  scrollViewport.dispatchEvent(new Event('scroll'))
  return true
}

function resolveRangeMoveAutoScrollStep(position: number, start: number, end: number): number {
  if (position < start + RANGE_MOVE_AUTOSCROLL_EDGE_PX) {
    return -resolveRangeMoveAutoScrollMagnitude(start + RANGE_MOVE_AUTOSCROLL_EDGE_PX - position)
  }
  if (position > end - RANGE_MOVE_AUTOSCROLL_EDGE_PX) {
    return resolveRangeMoveAutoScrollMagnitude(position - (end - RANGE_MOVE_AUTOSCROLL_EDGE_PX))
  }
  return 0
}

function resolveRangeMoveAutoScrollMagnitude(distanceInsideEdge: number): number {
  const intensity = clamp(distanceInsideEdge / RANGE_MOVE_AUTOSCROLL_EDGE_PX, 0, 1)
  return Math.max(1, Math.round(intensity * RANGE_MOVE_AUTOSCROLL_MAX_STEP_PX))
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value))
}
