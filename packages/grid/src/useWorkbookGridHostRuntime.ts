import { useCallback, useRef, useState, type Dispatch, type MutableRefObject, type SetStateAction } from 'react'
import type { VisibleRegionState } from './gridPointer.js'
import { useGridElementSize } from './useGridElementSize.js'

export interface WorkbookGridHostRuntimeState {
  readonly focusTargetRef: MutableRefObject<HTMLDivElement | null>
  readonly getVisibleRegion: () => VisibleRegionState
  readonly handleHostRef: (node: HTMLDivElement | null) => void
  readonly hostClientHeight: number
  readonly hostClientWidth: number
  readonly hostElement: HTMLDivElement | null
  readonly hostRef: MutableRefObject<HTMLDivElement | null>
  readonly liveVisibleRegionRef: MutableRefObject<VisibleRegionState>
  readonly scrollViewportRef: MutableRefObject<HTMLDivElement | null>
  readonly setVisibleRegion: Dispatch<SetStateAction<VisibleRegionState>>
  readonly visibleRegion: VisibleRegionState
}

export function useWorkbookGridHostRuntime(input: {
  readonly freezeCols: number
  readonly freezeRows: number
}): WorkbookGridHostRuntimeState {
  const { freezeCols, freezeRows } = input
  const hostRef = useRef<HTMLDivElement | null>(null)
  const focusTargetRef = useRef<HTMLDivElement | null>(null)
  const scrollViewportRef = useRef<HTMLDivElement | null>(null)
  const [hostElement, setHostElement] = useState<HTMLDivElement | null>(null)
  const [visibleRegion, setVisibleRegion] = useState<VisibleRegionState>({
    range: { x: 0, y: 0, width: 12, height: 24 },
    tx: 0,
    ty: 0,
    freezeRows,
    freezeCols,
  })
  const liveVisibleRegionRef = useRef<VisibleRegionState>(visibleRegion)
  const hostElementSize = useGridElementSize(hostElement)
  const handleHostRef = useCallback((node: HTMLDivElement | null) => {
    hostRef.current = node
    setHostElement(node)
  }, [])
  const getVisibleRegion = useCallback(() => liveVisibleRegionRef.current, [])

  return {
    focusTargetRef,
    getVisibleRegion,
    handleHostRef,
    hostClientHeight: hostElementSize.height,
    hostClientWidth: hostElementSize.width,
    hostElement,
    hostRef,
    liveVisibleRegionRef,
    scrollViewportRef,
    setVisibleRegion,
    visibleRegion,
  }
}
