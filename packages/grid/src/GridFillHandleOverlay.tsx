import { useEffect, useRef, type PointerEventHandler } from 'react'
import { resolveFillHandleOverlayBounds, type FillHandleOverlayBounds } from './gridFillHandle.js'
import type { Rectangle } from './gridTypes.js'
import type { WorkbookGridScrollStore } from './workbookGridScrollStore.js'

interface GridFillHandleOverlayProps {
  readonly scrollTransformStore: WorkbookGridScrollStore
  readonly selectionRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  readonly hidden: boolean
  readonly hostWidth: number
  readonly hostHeight: number
  readonly minX: number
  readonly minY: number
  readonly getCellBounds: (col: number, row: number) => Rectangle | undefined
  readonly onPointerDown: PointerEventHandler<HTMLButtonElement>
}

export function GridFillHandleOverlay(props: GridFillHandleOverlayProps) {
  const buttonRef = useRef<HTMLButtonElement | null>(null)

  useEffect(() => {
    const button = buttonRef.current
    if (!button) {
      return
    }

    const applyBounds = () => {
      const nextBounds: FillHandleOverlayBounds | undefined =
        props.hidden || !props.selectionRange
          ? undefined
          : resolveFillHandleOverlayBounds({
              sourceRange: props.selectionRange,
              hostBounds: {
                left: 0,
                top: 0,
                width: props.hostWidth,
                height: props.hostHeight,
              },
              getCellBounds: props.getCellBounds,
              minX: props.minX,
              minY: props.minY,
            })

      if (!nextBounds) {
        button.style.display = 'none'
        return
      }

      button.style.display = ''
      button.style.left = `${nextBounds.x}px`
      button.style.top = `${nextBounds.y}px`
      button.style.width = `${nextBounds.width}px`
      button.style.height = `${nextBounds.height}px`
    }

    applyBounds()
    return props.scrollTransformStore.subscribe(applyBounds)
  }, [
    props.getCellBounds,
    props.hidden,
    props.hostHeight,
    props.hostWidth,
    props.minX,
    props.minY,
    props.scrollTransformStore,
    props.selectionRange,
  ])

  return (
    <button
      aria-label="Fill handle"
      className="absolute z-30 cursor-crosshair rounded-[2px] border border-white bg-[var(--wb-accent)] shadow-[0_0_0_1px_rgba(33,86,58,0.32)] outline-none"
      data-grid-fill-handle="true"
      ref={buttonRef}
      onClick={(event) => {
        event.preventDefault()
        event.stopPropagation()
      }}
      onPointerDown={props.onPointerDown}
      style={{
        display: props.hidden || !props.selectionRange ? 'none' : undefined,
        touchAction: 'none',
      }}
      type="button"
    />
  )
}
