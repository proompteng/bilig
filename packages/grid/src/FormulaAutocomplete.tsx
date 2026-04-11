import type { FormulaSuggestion } from "./formulaAssist.js";
import { formulaPopupClass, formulaPopupOptionClass } from "./formula-bar-theme.js";

interface FormulaAutocompleteProps {
  readonly suggestions: readonly FormulaSuggestion[];
  readonly highlightedIndex: number;
  readonly onSelect: (index: number) => void;
}

export function FormulaAutocomplete({
  suggestions,
  highlightedIndex,
  onSelect,
}: FormulaAutocompleteProps) {
  if (suggestions.length === 0) {
    return null;
  }

  return (
    <div
      aria-label="Formula suggestions"
      className={`${formulaPopupClass()} absolute left-0 right-0 top-[calc(100%+0.375rem)] z-40`}
      data-testid="formula-autocomplete"
      id="formula-autocomplete"
      role="listbox"
    >
      <ul className="max-h-72 overflow-auto py-1">
        {suggestions.map((suggestion, index) => {
          const active = index === highlightedIndex;
          return (
            <li
              aria-selected={active}
              className={formulaPopupOptionClass({ active })}
              data-testid="formula-autocomplete-option"
              id={`formula-autocomplete-option-${index}`}
              key={`${suggestion.kind}:${suggestion.name}`}
              role="option"
              onMouseDown={(event) => {
                event.preventDefault();
                onSelect(index);
              }}
            >
              <div className="flex items-center justify-between gap-3">
                <div className="min-w-0">
                  <div className="truncate text-[12px] font-semibold text-[var(--color-mauve-950)]">
                    {suggestion.name}
                  </div>
                  <div className="truncate text-[11px] text-[var(--color-mauve-600)]">
                    {suggestion.kind === "function" ? suggestion.signature : suggestion.summary}
                  </div>
                </div>
                <span className="shrink-0 text-[10px] font-medium uppercase tracking-[0.08em] text-[var(--color-mauve-500)]">
                  {suggestion.kind === "function" ? suggestion.category : "Name"}
                </span>
              </div>
              {suggestion.kind === "function" ? (
                <div className="mt-1 truncate text-[11px] text-[var(--color-mauve-600)]">
                  {suggestion.summary}
                </div>
              ) : null}
            </li>
          );
        })}
      </ul>
    </div>
  );
}
