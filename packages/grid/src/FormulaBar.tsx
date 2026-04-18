import { useEffect, useMemo, useRef, useState } from 'react'
import type { WorkbookDefinedNameSnapshot } from '@bilig/protocol'
import { FormulaArgumentHint } from './FormulaArgumentHint.js'
import { FormulaAutocomplete } from './FormulaAutocomplete.js'
import { applyFormulaSuggestion, resolveFormulaAssistState, type FormulaSuggestion } from './formulaAssist.js'
import { formulaBarRootClass, formulaFieldAddonClass, formulaFieldShellClass, formulaInputClass } from './formula-bar-theme.js'
import { NameBox } from './NameBox.js'

interface FormulaBarProps {
  sheetName: string
  address: string
  selectionLabel?: string
  value: string
  resolvedValue: string
  isEditing: boolean
  definedNames?: readonly WorkbookDefinedNameSnapshot[]
  onBeginEdit(this: void, seed?: string): void
  onAddressCommit(this: void, next: string): void
  onChange(this: void, next: string): void
  onCommit(this: void): void
  onCancel(this: void): void
}

export function FormulaBar({
  sheetName,
  address,
  selectionLabel,
  value,
  resolvedValue,
  isEditing,
  definedNames,
  onBeginEdit,
  onAddressCommit,
  onChange,
  onCommit,
  onCancel,
}: FormulaBarProps) {
  const inputRef = useRef<HTMLTextAreaElement | null>(null)
  const nameBoxRef = useRef<HTMLInputElement | null>(null)
  const [isFormulaFocused, setIsFormulaFocused] = useState(false)
  const [formulaCaret, setFormulaCaret] = useState(value.length)
  const [highlightedSuggestionIndex, setHighlightedSuggestionIndex] = useState(0)
  const [dismissedAutocompleteValue, setDismissedAutocompleteValue] = useState<string | null>(null)
  const pendingSelectionRef = useRef<{ start: number; end: number } | null>(null)
  const MAX_FORMULA_HEIGHT = 128

  useEffect(() => {
    setFormulaCaret((current) => Math.min(value.length, current === 0 ? value.length : current))
    if (dismissedAutocompleteValue !== null && dismissedAutocompleteValue !== value) {
      setDismissedAutocompleteValue(null)
    }
  }, [dismissedAutocompleteValue, value])

  useEffect(() => {
    const textarea = inputRef.current
    if (!textarea) {
      return
    }
    textarea.style.height = '0px'
    const measuredHeight = Math.min(Math.max(textarea.scrollHeight, 32), MAX_FORMULA_HEIGHT)
    textarea.style.height = `${measuredHeight}px`
    textarea.style.overflowY = textarea.scrollHeight > MAX_FORMULA_HEIGHT ? 'auto' : 'hidden'
  }, [value])

  useEffect(() => {
    const pending = pendingSelectionRef.current
    if (!pending) {
      return
    }
    const input = inputRef.current
    if (!input) {
      return
    }
    input.focus()
    input.setSelectionRange(pending.start, pending.end)
    pendingSelectionRef.current = null
  }, [value])

  useEffect(() => {
    const handleGoToShortcut = (event: KeyboardEvent) => {
      if (event.defaultPrevented || event.altKey || event.shiftKey) {
        return
      }
      const hasPrimaryModifier = event.metaKey || event.ctrlKey
      if (!hasPrimaryModifier || event.key.toLowerCase() !== 'g') {
        return
      }
      event.preventDefault()
      nameBoxRef.current?.focus()
      nameBoxRef.current?.select()
    }

    window.addEventListener('keydown', handleGoToShortcut, true)
    return () => {
      window.removeEventListener('keydown', handleGoToShortcut, true)
    }
  }, [])

  const assistState = useMemo(
    () =>
      resolveFormulaAssistState({
        value,
        caret: formulaCaret,
        ...(definedNames ? { definedNames } : {}),
      }),
    [definedNames, formulaCaret, value],
  )
  const showAutocomplete = isEditing && assistState.suggestions.length > 0 && dismissedAutocompleteValue !== value

  useEffect(() => {
    setHighlightedSuggestionIndex((current) => Math.min(current, Math.max(assistState.suggestions.length - 1, 0)))
  }, [assistState.suggestions.length])

  const activeSuggestion = showAutocomplete
    ? (assistState.suggestions[highlightedSuggestionIndex] ?? assistState.suggestions[0] ?? null)
    : null
  const selectionStatus = `${sheetName}!${selectionLabel ?? address}`

  const commitSuggestion = (suggestion: FormulaSuggestion) => {
    if (assistState.tokenStart === null || assistState.tokenEnd === null) {
      return
    }
    const next = applyFormulaSuggestion({
      value,
      tokenStart: assistState.tokenStart,
      tokenEnd: assistState.tokenEnd,
      suggestion,
    })
    if (!isEditing) {
      onBeginEdit(value)
    }
    pendingSelectionRef.current = { start: next.caret, end: next.caret }
    setFormulaCaret(next.caret)
    setDismissedAutocompleteValue(null)
    onChange(next.value)
  }

  return (
    <div className={formulaBarRootClass()} data-testid="formula-bar">
      <NameBox
        address={address}
        onCommit={onAddressCommit}
        ref={nameBoxRef}
        selectionLabel={selectionLabel}
        sheetName={sheetName}
        {...(definedNames ? { definedNames } : {})}
      />
      <div className="min-w-0 flex-1">
        <label className="sr-only" htmlFor="formula-input">
          Formula
        </label>
        <div className="relative">
          <div className={formulaFieldShellClass({ focused: isFormulaFocused })} data-testid="formula-input-frame">
            <span aria-hidden="true" className={`${formulaFieldAddonClass()} w-10`}>
              fx
            </span>
            <textarea
              aria-activedescendant={showAutocomplete ? `formula-autocomplete-option-${highlightedSuggestionIndex}` : undefined}
              aria-controls={showAutocomplete ? 'formula-autocomplete' : undefined}
              aria-expanded={showAutocomplete ? 'true' : 'false'}
              aria-label="Formula"
              className={formulaInputClass()}
              data-testid="formula-input"
              id="formula-input"
              placeholder="Type a literal or =formula"
              ref={inputRef}
              role="combobox"
              value={value}
              rows={1}
              onBlur={(event) => {
                setIsFormulaFocused(false)
                const nextTarget = event.relatedTarget
                if (nextTarget instanceof Node && event.currentTarget.closest('.formula-bar')?.contains(nextTarget)) {
                  return
                }
                if (isEditing) {
                  onCommit()
                }
              }}
              onChange={(event) => {
                const nextValue = event.target.value
                if (!isEditing) {
                  onBeginEdit(nextValue)
                }
                setFormulaCaret(event.target.selectionStart ?? nextValue.length)
                onChange(nextValue)
              }}
              onClick={(event) => {
                setFormulaCaret(event.currentTarget.selectionStart ?? event.currentTarget.value.length)
              }}
              onFocus={(event) => {
                setIsFormulaFocused(true)
                setFormulaCaret(event.currentTarget.selectionStart ?? event.currentTarget.value.length)
                if (!isEditing) {
                  onBeginEdit(value)
                }
              }}
              onKeyDown={(event) => {
                event.stopPropagation()
                if (event.key === 'Enter' && event.altKey) {
                  return
                }
                if (showAutocomplete && event.key === 'ArrowDown') {
                  event.preventDefault()
                  setHighlightedSuggestionIndex((current) => Math.min(current + 1, assistState.suggestions.length - 1))
                  return
                }
                if (showAutocomplete && event.key === 'ArrowUp') {
                  event.preventDefault()
                  setHighlightedSuggestionIndex((current) => Math.max(current - 1, 0))
                  return
                }
                if (activeSuggestion && showAutocomplete && (event.key === 'Enter' || event.key === 'Tab')) {
                  event.preventDefault()
                  commitSuggestion(activeSuggestion)
                  return
                }
                if (event.key === 'Enter') {
                  event.preventDefault()
                  onCommit()
                  return
                }
                if (event.key === 'Escape') {
                  event.preventDefault()
                  if (showAutocomplete) {
                    setDismissedAutocompleteValue(value)
                    return
                  }
                  onCancel()
                }
              }}
              onKeyUp={(event) => {
                setFormulaCaret(event.currentTarget.selectionStart ?? event.currentTarget.value.length)
              }}
              onSelect={(event) => {
                setFormulaCaret(event.currentTarget.selectionStart ?? event.currentTarget.value.length)
              }}
            />
          </div>
          <div
            className="mt-1.5 flex items-center justify-between gap-3 text-[11px] text-[var(--wb-text-subtle)]"
            data-testid="formula-bar-meta"
          >
            <span className="truncate font-medium text-[var(--wb-text-muted)]" data-testid="formula-selection-label">
              {selectionLabel ?? address}
            </span>
            <span className="truncate">{resolvedValue || '∅'}</span>
          </div>
          {showAutocomplete ? (
            <FormulaAutocomplete
              highlightedIndex={highlightedSuggestionIndex}
              suggestions={assistState.suggestions}
              onSelect={(index) => {
                const suggestion = assistState.suggestions[index]
                if (!suggestion) {
                  return
                }
                setHighlightedSuggestionIndex(index)
                commitSuggestion(suggestion)
              }}
            />
          ) : null}
        </div>
        {isEditing && assistState.activeFunction ? (
          <div className="mt-1.5">
            <FormulaArgumentHint
              activeArgumentIndex={assistState.activeFunction.activeArgumentIndex}
              entry={assistState.activeFunction.entry}
            />
          </div>
        ) : null}
      </div>
      <span className="sr-only" data-testid="status-selection">
        {selectionStatus}
      </span>
      <span className="sr-only" data-testid="formula-resolved-value">
        {resolvedValue || '∅'}
      </span>
    </div>
  )
}
