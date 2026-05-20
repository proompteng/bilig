import { useEffect, useRef, useState } from 'react'

interface ToolbarScrollCueState {
  readonly hasOverflow: boolean
  readonly isAtStart: boolean
  readonly isAtEnd: boolean
}

export function useToolbarScrollCue() {
  const scrollContainerRef = useRef<HTMLDivElement | null>(null)
  const [state, setState] = useState<ToolbarScrollCueState>({
    hasOverflow: false,
    isAtStart: true,
    isAtEnd: true,
  })

  useEffect(() => {
    const scrollContainer = scrollContainerRef.current
    if (!scrollContainer) {
      return
    }

    const updateScrollCue = () => {
      const maxScrollLeft = Math.max(0, scrollContainer.scrollWidth - scrollContainer.clientWidth)
      const nextState: ToolbarScrollCueState = {
        hasOverflow: maxScrollLeft > 1,
        isAtStart: scrollContainer.scrollLeft <= 1,
        isAtEnd: scrollContainer.scrollLeft >= maxScrollLeft - 1,
      }

      setState((currentState) =>
        currentState.hasOverflow === nextState.hasOverflow &&
        currentState.isAtStart === nextState.isAtStart &&
        currentState.isAtEnd === nextState.isAtEnd
          ? currentState
          : nextState,
      )
    }

    updateScrollCue()
    scrollContainer.addEventListener('scroll', updateScrollCue, { passive: true })
    window.addEventListener('resize', updateScrollCue)

    let resizeObserver: ResizeObserver | null = null
    if (typeof ResizeObserver !== 'undefined') {
      resizeObserver = new ResizeObserver(updateScrollCue)
      resizeObserver.observe(scrollContainer)
    }

    return () => {
      scrollContainer.removeEventListener('scroll', updateScrollCue)
      window.removeEventListener('resize', updateScrollCue)
      resizeObserver?.disconnect()
    }
  }, [])

  return {
    scrollContainerRef,
    showBackwardCue: state.hasOverflow && !state.isAtStart,
    showForwardCue: state.hasOverflow && !state.isAtEnd,
  } as const
}
