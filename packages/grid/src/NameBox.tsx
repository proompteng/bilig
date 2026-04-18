import { forwardRef, useEffect, useId, useState } from 'react'
import type { WorkbookDefinedNameSnapshot } from '@bilig/protocol'
import { resolveNameBoxDisplayValue } from './formulaAssist.js'
import { formulaInlineMessageClass, formulaStandaloneInputClass } from './formula-bar-theme.js'

interface NameBoxProps {
  readonly address: string
  readonly definedNames?: readonly WorkbookDefinedNameSnapshot[]
  readonly sheetName: string
  readonly selectionLabel?: string | undefined
  readonly onCommit: (next: string) => boolean
}

export const NameBox = forwardRef<HTMLInputElement, NameBoxProps>(function NameBox(
  { address, definedNames, sheetName, selectionLabel, onCommit },
  ref,
) {
  const displayValue = resolveNameBoxDisplayValue({
    sheetName,
    address,
    ...(selectionLabel !== undefined ? { selectionLabel } : {}),
    ...(definedNames ? { definedNames } : {}),
  })
  const [inputValue, setInputValue] = useState(displayValue)
  const [errorMessage, setErrorMessage] = useState<string | null>(null)
  const errorId = useId()

  useEffect(() => {
    setInputValue(displayValue)
    setErrorMessage(null)
  }, [displayValue, sheetName])

  return (
    <div className="w-[168px] shrink-0">
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
        ref={ref}
        value={inputValue}
        onBlur={() => {
          if (!errorMessage) {
            setInputValue(displayValue)
          }
        }}
        onChange={(event) => {
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
              setErrorMessage(null)
            }
          }
          if (event.key === 'Escape') {
            event.preventDefault()
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
