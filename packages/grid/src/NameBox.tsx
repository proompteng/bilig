import { forwardRef, useCallback, useEffect, useId, useRef, useState, type ForwardedRef } from 'react'
import type { WorkbookDefinedNameSnapshot } from '@bilig/protocol'
import { resolveNameBoxDisplayValue } from './formulaAssist.js'
import { formulaInlineMessageClass, formulaStandaloneInputClass } from './formula-bar-theme.js'
import { workbookTextControlProps } from './workbookTextControls.js'

interface NameBoxProps {
  readonly address: string
  readonly definedNames?: readonly WorkbookDefinedNameSnapshot[]
  readonly sheetName: string
  readonly selectionLabel?: string | undefined
  readonly onCommit: (next: string) => boolean
  readonly onCommitSuccess?: (() => void) | undefined
}

export const NameBox = forwardRef<HTMLInputElement, NameBoxProps>(function NameBox(
  { address, definedNames, sheetName, selectionLabel, onCommit, onCommitSuccess },
  ref,
) {
  const displayValue = resolveNameBoxDisplayValue({
    sheetName,
    address,
    ...(selectionLabel !== undefined ? { selectionLabel } : {}),
    ...(definedNames ? { definedNames } : {}),
  })
  const inputRef = useRef<HTMLInputElement | null>(null)
  const [inputValue, setInputValue] = useState(displayValue)
  const [errorMessage, setErrorMessage] = useState<string | null>(null)
  const errorId = useId()
  const isDirtyRef = useRef(false)
  const setInputRef = useCallback(
    (node: HTMLInputElement | null) => {
      inputRef.current = node
      assignForwardedRef(ref, node)
    },
    [ref],
  )

  useEffect(() => {
    if (isDirtyRef.current) {
      return
    }
    setInputValue(displayValue)
    setErrorMessage(null)
  }, [displayValue, sheetName])

  return (
    <div className="w-20 shrink-0 sm:w-28">
      <label className="sr-only" htmlFor="name-box-input">
        Name
      </label>
      <input
        aria-label="Name box"
        aria-describedby={errorMessage ? errorId : undefined}
        aria-invalid={errorMessage ? 'true' : undefined}
        className={formulaStandaloneInputClass({ invalid: Boolean(errorMessage) })}
        data-testid="name-box"
        id="name-box-input"
        ref={setInputRef}
        {...workbookTextControlProps}
        value={inputValue}
        onBlur={() => {
          isDirtyRef.current = false
          if (!errorMessage) {
            setInputValue(displayValue)
          }
        }}
        onChange={(event) => {
          isDirtyRef.current = true
          setInputValue(event.target.value)
          if (errorMessage) {
            setErrorMessage(null)
          }
        }}
        onKeyDown={(event) => {
          event.stopPropagation()
          if (event.key === 'Enter') {
            event.preventDefault()
            const didCommit = onCommit(event.currentTarget.value)
            if (!didCommit) {
              setErrorMessage('Unknown range or name')
            } else {
              isDirtyRef.current = false
              setErrorMessage(null)
              event.currentTarget.blur()
              onCommitSuccess?.()
            }
          }
          if (event.key === 'Escape') {
            event.preventDefault()
            isDirtyRef.current = false
            setErrorMessage(null)
            setInputValue(displayValue)
          }
        }}
      />
      {errorMessage ? (
        <p className={formulaInlineMessageClass()} data-testid="name-box-error" id={errorId}>
          {errorMessage}
        </p>
      ) : null}
    </div>
  )
})

function assignForwardedRef(ref: ForwardedRef<HTMLInputElement>, node: HTMLInputElement | null): void {
  if (typeof ref === 'function') {
    ref(node)
    return
  }
  if (ref) {
    ref.current = node
  }
}
