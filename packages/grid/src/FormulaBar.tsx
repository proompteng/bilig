import { useEffect, useState } from "react";

interface FormulaBarProps {
  sheetName: string;
  address: string;
  value: string;
  resolvedValue: string;
  isEditing: boolean;
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
  onBeginEdit,
  onAddressCommit,
  onChange,
  onCommit,
  onCancel,
}: FormulaBarProps) {
  const [addressValue, setAddressValue] = useState(address);

  useEffect(() => {
    setAddressValue(address);
  }, [address, sheetName]);

  return (
    <div
      className="formula-bar flex items-center gap-2 border-b border-[var(--wb-border)] bg-[var(--wb-surface)] px-2.5 py-1.5 font-sans"
      data-testid="formula-bar"
    >
      <div className="w-[92px] shrink-0">
        <label className="sr-only" htmlFor="name-box-input">
          Name
        </label>
        <input
          aria-label="Name box"
          className="box-border h-8 w-full rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-2.5 text-[12px] font-medium leading-none text-[var(--wb-text)] outline-none transition-[border-color,box-shadow,background-color] focus:border-[var(--wb-accent)] focus:bg-[var(--wb-surface)] focus:ring-2 focus:ring-[var(--wb-accent-ring)]"
          data-testid="name-box"
          id="name-box-input"
          value={addressValue}
          onBlur={() => setAddressValue(address)}
          onChange={(event) => setAddressValue(event.target.value.toUpperCase())}
          onKeyDown={(event) => {
            event.stopPropagation();
            if (event.key === "Enter") {
              event.preventDefault();
              const nextValue = event.currentTarget.value.toUpperCase();
              setAddressValue(nextValue);
              onAddressCommit(nextValue);
            }
            if (event.key === "Escape") {
              event.preventDefault();
              event.currentTarget.value = address;
              setAddressValue(address);
            }
          }}
        />
      </div>
      <div className="min-w-0 flex-1">
        <label className="sr-only" htmlFor="formula-input">
          Formula
        </label>
        <div
          className="box-border flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)]"
          data-testid="formula-input-frame"
        >
          <span
            aria-hidden="true"
            className="inline-flex h-full w-8 shrink-0 items-center justify-center border-r border-[var(--wb-border)] bg-[var(--wb-surface-muted)] text-[11px] font-semibold uppercase tracking-[0.1em] leading-none text-[var(--wb-text-subtle)]"
          >
            fx
          </span>
          <input
            aria-label="Formula"
            className="h-full min-w-0 flex-1 border-0 bg-[var(--wb-surface)] px-3 text-[12px] leading-none text-[var(--wb-text)] outline-none"
            data-testid="formula-input"
            id="formula-input"
            placeholder="Type a literal or =formula"
            value={value}
            onBlur={(event) => {
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
              if (!isEditing) {
                onBeginEdit(event.target.value);
                return;
              }
              onChange(event.target.value);
            }}
            onFocus={() => {
              if (!isEditing) {
                onBeginEdit();
              }
            }}
            onKeyDown={(event) => {
              event.stopPropagation();
              if (event.key === "Enter") {
                event.preventDefault();
                onCommit();
              }
              if (event.key === "Escape") {
                event.preventDefault();
                onCancel();
              }
            }}
          />
        </div>
      </div>
      <span className="sr-only" data-testid="formula-resolved-value">
        {resolvedValue || "∅"}
      </span>
    </div>
  );
}
