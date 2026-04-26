import { useLayoutEffect, useState } from 'react'

export interface GridElementSize {
  readonly width: number
  readonly height: number
}

const EMPTY_ELEMENT_SIZE: GridElementSize = Object.freeze({ width: 0, height: 0 })

export function useGridElementSize(element: HTMLElement | null): GridElementSize {
  const [size, setSize] = useState<GridElementSize>(EMPTY_ELEMENT_SIZE)

  useLayoutEffect(() => {
    if (!element) {
      setSize((current) => (current.width === 0 && current.height === 0 ? current : EMPTY_ELEMENT_SIZE))
      return
    }

    const syncSize = () => {
      const next = {
        width: element.clientWidth,
        height: element.clientHeight,
      }
      setSize((current) => (current.width === next.width && current.height === next.height ? current : next))
    }

    syncSize()
    if (typeof ResizeObserver === 'undefined') {
      return
    }

    const observer = new ResizeObserver(syncSize)
    observer.observe(element)
    return () => {
      observer.disconnect()
    }
  }, [element])

  return size
}
