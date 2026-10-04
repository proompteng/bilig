import { useId, useState } from 'react'
import type { ModelFormat } from './model-document.js'

export function ModelField(props: {
  readonly label: string
  readonly value: string
  readonly onCommit: (value: string) => void
  readonly kind?: 'text' | 'number' | 'positive-number' | 'formula'
  readonly className?: string
  readonly maxLength?: number
}) {
  const [draft, setDraft] = useState<string | null>(null)
  const [error, setError] = useState<string | null>(null)
  const errorId = useId()
  const commit = () => {
    if (draft === null) return
    const next = draft.trim()
    if (!next) {
      setError('Enter a value.')
      return
    }
    if ((props.kind === 'number' || props.kind === 'positive-number') && !Number.isFinite(Number(next))) {
      setError('Enter a finite number.')
      return
    }
    if (props.kind === 'positive-number' && Number(next) <= 0) {
      setError('Enter a positive number.')
      return
    }
    if (props.kind === 'formula' && !next.startsWith('=')) {
      setError('Start formulas with =.')
      return
    }
    setError(null)
    setDraft(null)
    if (next !== props.value) props.onCommit(next)
  }
  return (
    <div className={`model-field ${props.className ?? ''}`}>
      <input
        aria-label={props.label}
        aria-invalid={error !== null}
        aria-describedby={error ? errorId : undefined}
        value={draft ?? props.value}
        inputMode={props.kind === 'number' || props.kind === 'positive-number' ? 'decimal' : undefined}
        maxLength={props.maxLength ?? 120}
        spellCheck={props.kind !== 'formula'}
        onChange={(event) => {
          setDraft(event.target.value)
          setError(null)
        }}
        onBlur={commit}
        onKeyDown={(event) => {
          if (event.key === 'Enter') {
            event.preventDefault()
            commit()
          }
          if (event.key === 'Escape') {
            setDraft(null)
            setError(null)
          }
        }}
      />
      {error ? (
        <span className="model-field-error" id={errorId} role="alert">
          {error}
        </span>
      ) : null}
    </div>
  )
}

export function ModelFormatSelect(props: {
  readonly label: string
  readonly value: ModelFormat
  readonly onChange: (format: ModelFormat) => void
}) {
  return (
    <select
      aria-label={props.label}
      className="model-format-select"
      value={props.value}
      onChange={(event) => {
        const value = event.target.value
        if (value === 'number' || value === 'currency' || value === 'percent') props.onChange(value)
      }}
    >
      <option value="number">Number</option>
      <option value="currency">USD</option>
      <option value="percent">Percent</option>
    </select>
  )
}
