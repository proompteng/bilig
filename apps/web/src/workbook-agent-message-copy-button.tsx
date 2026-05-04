import { useCallback, useEffect, useRef, useState } from 'react'
import { Button } from '@base-ui/react/button'
import { Check, Copy } from 'lucide-react'
import { cn } from './cn.js'

async function copyTextToClipboard(text: string): Promise<void> {
  if (navigator.clipboard?.writeText) {
    try {
      await navigator.clipboard.writeText(text)
      return
    } catch {
      // fallback below
    }
  }

  const textarea = document.createElement('textarea')
  textarea.value = text
  textarea.readOnly = true
  textarea.style.position = 'fixed'
  textarea.style.inset = '0 auto auto 0'
  textarea.style.width = '1px'
  textarea.style.height = '1px'
  textarea.style.opacity = '0'
  textarea.style.pointerEvents = 'none'

  const selection = document.getSelection()
  const previousRange = selection && selection.rangeCount > 0 ? selection.getRangeAt(0).cloneRange() : null

  document.body.appendChild(textarea)
  let copied = false
  try {
    textarea.select()
    textarea.setSelectionRange(0, textarea.value.length)
    copied = typeof document.execCommand === 'function' ? document.execCommand('copy') : false
  } finally {
    textarea.remove()

    if (selection) {
      selection.removeAllRanges()
      if (previousRange) {
        selection.addRange(previousRange)
      }
    }
  }

  if (!copied) {
    throw new Error('Clipboard write failed')
  }
}

export function MessageCopyButton(props: {
  readonly entryId: string
  readonly text: string
  readonly messageKind: 'assistant' | 'user'
  readonly className?: string
}) {
  const resetTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null)
  const [copyStatus, setCopyStatus] = useState<'idle' | 'copied' | 'failed'>('idle')

  const clearResetTimer = useCallback(() => {
    if (resetTimerRef.current !== null) {
      clearTimeout(resetTimerRef.current)
      resetTimerRef.current = null
    }
  }, [])

  useEffect(() => {
    setCopyStatus('idle')
  }, [props.text])

  useEffect(() => clearResetTimer, [clearResetTimer])

  const handleCopy = useCallback(async () => {
    clearResetTimer()
    try {
      await copyTextToClipboard(props.text)
      setCopyStatus('copied')
      resetTimerRef.current = setTimeout(() => {
        setCopyStatus('idle')
        resetTimerRef.current = null
      }, 1800)
    } catch {
      setCopyStatus('failed')
    }
  }, [clearResetTimer, props.text])

  const copied = copyStatus === 'copied'
  const labelPrefix = props.messageKind === 'assistant' ? 'assistant' : 'user'
  const label =
    copyStatus === 'copied'
      ? `Copied ${labelPrefix} message`
      : copyStatus === 'failed'
        ? `Copy failed. Retry ${labelPrefix} message`
        : `Copy ${labelPrefix} message`

  return (
    <Button
      aria-label={label}
      className={cn(
        'inline-flex size-7 shrink-0 items-center justify-center rounded-[var(--wb-radius-control)] border border-transparent text-[var(--wb-text-subtle)] transition-colors hover:border-[var(--wb-border)] hover:bg-[var(--wb-surface)] hover:text-[var(--wb-text)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1',
        props.className,
      )}
      data-copy-state={copyStatus}
      data-testid={`workbook-agent-message-copy-${props.entryId}`}
      title={label}
      type="button"
      onClick={() => {
        void handleCopy()
      }}
    >
      {copied ? (
        <Check aria-hidden="true" className="size-4" strokeWidth={2.2} />
      ) : (
        <Copy aria-hidden="true" className="size-4" strokeWidth={1.8} />
      )}
    </Button>
  )
}
