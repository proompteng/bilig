import { forwardRef, useEffect, useState } from 'react'
import type { WorkbookDefinedNameSnapshot } from '@bilig/protocol'
import { resolveNameBoxDisplayValue } from './formulaAssist.js'
import { formulaStandaloneInputClass } from './formula-bar-theme.js'

interface NameBoxProps {
  readonly address: string
  readonly definedNames?: readonly WorkbookDefinedNameSnapshot[]
  readonly sheetName: string
  readonly selectionLabel?: string
  readonly onCommit: (next: string) => void
}

export const NameBox = forwardRef<HTMLInputElement, NameBoxProps>(function NameBox(
  { address, definedNames, sheetName, selectionLabel, onCommit },
  ref,
) {
  const displayValue = resolveNameBoxDisplayValue({
    sheetName,
    address,
    selectionLabel,
    ...(definedNames ? { definedNames } : {}),
  })
  const [inputValue, setInputValue] = useState(displayValue)

  useEffect(() => {
    setInputValue(displayValue)
  }, [displayValue, sheetName])

  return (
    <div className="w-[168px] shrink-0">
      <label className="sr-only" htmlFor="name-box-input">
        Name
      </label>
      <input
        aria-label="Name box"
        className={formulaStandaloneInputClass()}
        data-testid="name-box"
        id="name-box-input"
        ref={ref}
        value={inputValue}
        onBlur={() => setInputValue(displayValue)}
        onChange={(event) => setInputValue(event.target.value)}
        onKeyDown={(event) => {
          event.stopPropagation()
          if (event.key === 'Enter') {
            event.preventDefault()
            onCommit(event.currentTarget.value)
          }
          if (event.key === 'Escape') {
            event.preventDefault()
            setInputValue(displayValue)
          }
        }}
      />
    </div>
  )
})
