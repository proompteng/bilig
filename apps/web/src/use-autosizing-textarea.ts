import { useLayoutEffect, useRef } from 'react'

type AutoSizingTextareaOptions = Readonly<{
  readonly value: string
  readonly minHeight: number
  readonly maxHeight: number
}>

export function useAutoSizingTextarea({ value, minHeight, maxHeight }: AutoSizingTextareaOptions) {
  const textareaRef = useRef<HTMLTextAreaElement | null>(null)
  const viewportRef = useRef<HTMLDivElement | null>(null)

  useLayoutEffect(() => {
    const textarea = textareaRef.current
    const viewport = viewportRef.current
    if (!textarea || !viewport) {
      return
    }

    textarea.style.height = '0px'
    const measuredHeight = Math.max(textarea.scrollHeight, minHeight)
    const viewportHeight = Math.min(measuredHeight, maxHeight)
    textarea.style.height = `${measuredHeight}px`
    viewport.style.height = `${viewportHeight}px`
    viewport.scrollTop = viewport.scrollHeight
  }, [value, minHeight, maxHeight])

  return { textareaRef, viewportRef }
}
