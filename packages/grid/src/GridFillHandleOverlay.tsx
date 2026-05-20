import { useLayoutEffect, useRef, type PointerEventHandler } from 'react'
import { resolveFillHandleHitTargetBounds, type FillHandleOverlayBounds } from './gridFillHandle.js'
import type { GridGeometrySnapshot } from './gridGeometry.js'
import type { Rectangle } from './gridTypes.js'
import type { WorkbookGridScrollStore } from './workbookGridScrollStore.js'

interface GridFillHandleOverlayProps {
  readonly scrollTransformStore: WorkbookGridScrollStore
  readonly getGeometrySnapshot: () => GridGeometrySnapshot | null
  readonly selectionRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  readonly hidden: boolean
  readonly hostWidth: number
  readonly hostHeight: number
  readonly minX: number
  readonly minY: number
  readonly onPointerDown: PointerEventHandler<HTMLDivElement>
}

export function GridFillHandleOverlay(props: GridFillHandleOverlayProps) {
  const { getGeometrySnapshot, hidden, hostHeight, hostWidth, minX, minY, onPointerDown, scrollTransformStore, selectionRange } = props
  const handleRef = useRef<HTMLDivElement | null>(null)

  useLayoutEffect(() => {
    const handle = handleRef.current
    if (!handle) {
      return
    }

    const applyBounds = () => {
      let nextBounds: FillHandleOverlayBounds | undefined
      if (!hidden && selectionRange) {
        const visualBounds = getGeometrySnapshot()?.fillHandleScreenRect(selectionRange) ?? null
        nextBounds = visualBounds
          ? resolveFillHandleHitTargetBounds({
              hostBounds: {
                width: hostWidth,
                height: hostHeight,
              },
              minX,
              minY,
              visualBounds,
            })
          : undefined
      }

      if (!nextBounds) {
        handle.style.display = 'none'
        return
      }

      handle.style.display = ''
      handle.style.left = `${nextBounds.x}px`
      handle.style.top = `${nextBounds.y}px`
      handle.style.width = `${nextBounds.width}px`
      handle.style.height = `${nextBounds.height}px`
    }

    applyBounds()
    return scrollTransformStore.subscribe(applyBounds)
  }, [getGeometrySnapshot, hidden, hostHeight, hostWidth, minX, minY, scrollTransformStore, selectionRange])

  return (
    <div
      aria-hidden="true"
      className="absolute z-30 cursor-crosshair rounded-[2px] border border-white bg-[var(--wb-accent)] shadow-[0_0_0_1px_rgba(33,86,58,0.32)] outline-none"
      data-grid-fill-handle="true"
      ref={handleRef}
      onClick={(event) => {
        event.preventDefault()
        event.stopPropagation()
      }}
      onPointerDown={onPointerDown}
      style={{
        display: hidden || !selectionRange ? 'none' : undefined,
        opacity: 0,
        touchAction: 'none',
      }}
    />
  )
}
