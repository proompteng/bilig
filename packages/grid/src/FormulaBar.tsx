import { useEffect, useMemo, useRef, useState } from "react";
import type { WorkbookDefinedNameSnapshot } from "@bilig/protocol";
import { FormulaArgumentHint } from "./FormulaArgumentHint.js";
import { FormulaAutocomplete } from "./FormulaAutocomplete.js";
import {
  applyFormulaSuggestion,
  resolveFormulaAssistState,
  type FormulaSuggestion,
} from "./formulaAssist.js";
import {
  formulaBarRootClass,
  formulaFieldAddonClass,
  formulaFieldShellClass,
  formulaInputClass,
} from "./formula-bar-theme.js";
import { NameBox } from "./NameBox.js";

interface FormulaBarProps {
  sheetName: string;
  address: string;
  value: string;
  resolvedValue: string;
  isEditing: boolean;
  definedNames?: readonly WorkbookDefinedNameSnapshot[];
  onBeginEdit(this: void, seed?: string): void;
  onAddressCommit(this: void, next: string): void;
  onChange(this: void, next: string): void;
  onCommit(this: void): void;
  onCancel(this: void): void;
}

export function FormulaBar({
  sheetName,
  address,
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
  const inputRef = useRef<HTMLInputElement | null>(null);
  const [isFormulaFocused, setIsFormulaFocused] = useState(false);
  const [formulaCaret, setFormulaCaret] = useState(value.length);
  const [highlightedSuggestionIndex, setHighlightedSuggestionIndex] = useState(0);
  const [dismissedAutocompleteValue, setDismissedAutocompleteValue] = useState<string | null>(null);
  const pendingSelectionRef = useRef<{ start: number; end: number } | null>(null);

  useEffect(() => {
    setFormulaCaret((current) => Math.min(value.length, current === 0 ? value.length : current));
    if (dismissedAutocompleteValue !== null && dismissedAutocompleteValue !== value) {
      setDismissedAutocompleteValue(null);
    }
  }, [dismissedAutocompleteValue, value]);

  useEffect(() => {
    const pending = pendingSelectionRef.current;
    if (!pending) {
      return;
    }
    const input = inputRef.current;
    if (!input) {
      return;
    }
    input.focus();
    input.setSelectionRange(pending.start, pending.end);
    pendingSelectionRef.current = null;
  }, [value]);

  const assistState = useMemo(
    () =>
      resolveFormulaAssistState({
        value,
        caret: formulaCaret,
        ...(definedNames ? { definedNames } : {}),
      }),
    [definedNames, formulaCaret, value],
  );
  const showAutocomplete =
    isEditing && assistState.suggestions.length > 0 && dismissedAutocompleteValue !== value;

  useEffect(() => {
    setHighlightedSuggestionIndex((current) =>
      Math.min(current, Math.max(assistState.suggestions.length - 1, 0)),
    );
  }, [assistState.suggestions.length]);

  const activeSuggestion = showAutocomplete
    ? (assistState.suggestions[highlightedSuggestionIndex] ?? assistState.suggestions[0] ?? null)
    : null;

  const commitSuggestion = (suggestion: FormulaSuggestion) => {
    if (assistState.tokenStart === null || assistState.tokenEnd === null) {
      return;
    }
    const next = applyFormulaSuggestion({
      value,
      tokenStart: assistState.tokenStart,
      tokenEnd: assistState.tokenEnd,
      suggestion,
    });
    if (!isEditing) {
      onBeginEdit(value);
    }
    pendingSelectionRef.current = { start: next.caret, end: next.caret };
    setFormulaCaret(next.caret);
    setDismissedAutocompleteValue(null);
    onChange(next.value);
  };

  return (
    <div className={formulaBarRootClass()} data-testid="formula-bar">
      <NameBox
        address={address}
        onCommit={onAddressCommit}
        sheetName={sheetName}
        {...(definedNames ? { definedNames } : {})}
      />
      <div className="min-w-0 flex-1">
        <label className="sr-only" htmlFor="formula-input">
          Formula
        </label>
        <div className="relative">
          <div
            className={formulaFieldShellClass({ focused: isFormulaFocused })}
            data-testid="formula-input-frame"
          >
            <span aria-hidden="true" className={`${formulaFieldAddonClass()} w-10`}>
              fx
            </span>
            <input
              aria-activedescendant={
                showAutocomplete
                  ? `formula-autocomplete-option-${highlightedSuggestionIndex}`
                  : undefined
              }
              aria-controls={showAutocomplete ? "formula-autocomplete" : undefined}
              aria-expanded={showAutocomplete ? "true" : "false"}
              aria-label="Formula"
              className={formulaInputClass()}
              data-testid="formula-input"
              id="formula-input"
              placeholder="Type a literal or =formula"
              ref={inputRef}
              role="combobox"
              value={value}
              onBlur={(event) => {
                setIsFormulaFocused(false);
                const nextTarget = event.relatedTarget;
                if (
                  nextTarget instanceof Node &&
                  event.currentTarget.closest(".formula-bar")?.contains(nextTarget)
                ) {
                  return;
                }
                if (isEditing) {
                  onCommit();
                }
              }}
              onChange={(event) => {
                const nextValue = event.target.value;
                if (!isEditing) {
                  onBeginEdit(nextValue);
                }
                setFormulaCaret(event.target.selectionStart ?? nextValue.length);
                onChange(nextValue);
              }}
              onClick={(event) => {
                setFormulaCaret(
                  event.currentTarget.selectionStart ?? event.currentTarget.value.length,
                );
              }}
              onFocus={(event) => {
                setIsFormulaFocused(true);
                setFormulaCaret(
                  event.currentTarget.selectionStart ?? event.currentTarget.value.length,
                );
                if (!isEditing) {
                  onBeginEdit(value);
                }
              }}
              onKeyDown={(event) => {
                event.stopPropagation();
                if (showAutocomplete && event.key === "ArrowDown") {
                  event.preventDefault();
                  setHighlightedSuggestionIndex((current) =>
                    Math.min(current + 1, assistState.suggestions.length - 1),
                  );
                  return;
                }
                if (showAutocomplete && event.key === "ArrowUp") {
                  event.preventDefault();
                  setHighlightedSuggestionIndex((current) => Math.max(current - 1, 0));
                  return;
                }
                if (
                  activeSuggestion &&
                  showAutocomplete &&
                  (event.key === "Enter" || event.key === "Tab")
                ) {
                  event.preventDefault();
                  commitSuggestion(activeSuggestion);
                  return;
                }
                if (event.key === "Enter") {
                  event.preventDefault();
                  onCommit();
                  return;
                }
                if (event.key === "Escape") {
                  event.preventDefault();
                  if (showAutocomplete) {
                    setDismissedAutocompleteValue(value);
                    return;
                  }
                  onCancel();
                }
              }}
              onKeyUp={(event) => {
                setFormulaCaret(
                  event.currentTarget.selectionStart ?? event.currentTarget.value.length,
                );
              }}
              onSelect={(event) => {
                setFormulaCaret(
                  event.currentTarget.selectionStart ?? event.currentTarget.value.length,
                );
              }}
            />
          </div>
          {showAutocomplete ? (
            <FormulaAutocomplete
              highlightedIndex={highlightedSuggestionIndex}
              suggestions={assistState.suggestions}
              onSelect={(index) => {
                const suggestion = assistState.suggestions[index];
                if (!suggestion) {
                  return;
                }
                setHighlightedSuggestionIndex(index);
                commitSuggestion(suggestion);
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
      <span className="sr-only" data-testid="formula-resolved-value">
        {resolvedValue || "∅"}
      </span>
    </div>
  );
}
