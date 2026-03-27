import { useEffect, useState } from "react";

interface FormulaBarProps {
  variant?: "playground" | "product";
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
  onClear(this: void): void;
}

export function FormulaBar({
  variant = "playground",
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
  onClear,
}: FormulaBarProps) {
  const [addressValue, setAddressValue] = useState(address);
  const product = variant === "product";

  useEffect(() => {
    setAddressValue(address);
  }, [address, sheetName]);

  const showResolvedValue = variant !== "product";

  return (
    <div
      className={
        product
          ? "formula-bar flex items-center gap-2 border-b border-[#dadce0] bg-white px-3 py-2"
          : "formula-bar"
      }
      data-testid="formula-bar"
    >
      <div className={product ? "w-[112px] shrink-0" : "name-box-shell"}>
        <label className={product ? "sr-only" : "formula-meta-label"} htmlFor="name-box-input">
          Name
        </label>
        <input
          aria-label="Name box"
          className={
            product
              ? "box-border h-8 w-full rounded-[2px] border border-[#dadce0] bg-white px-3 text-[14px] leading-none text-[#202124] outline-none focus:border-[#1a73e8]"
              : undefined
          }
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
      <div className={product ? "min-w-0 flex-1" : "formula-input-shell"}>
        <label className={product ? "sr-only" : "formula-meta-label"} htmlFor="formula-input">
          Formula
        </label>
        <div
          className={
            product
              ? "box-border flex h-8 items-center rounded-[2px] border border-[#dadce0] bg-white"
              : "formula-input-frame"
          }
          data-testid={product ? "formula-input-frame" : undefined}
        >
          <span
            aria-hidden="true"
            className={
              product
                ? "inline-flex h-full w-8 shrink-0 items-center justify-center border-r border-[#dadce0] text-[16px] font-semibold leading-none text-[#5f6368]"
                : "formula-fx"
            }
          >
            fx
          </span>
          <input
            aria-label="Formula"
            className={
              product
                ? "h-full min-w-0 flex-1 border-0 bg-white px-3 text-[14px] leading-none text-[#202124] outline-none"
                : undefined
            }
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
      {showResolvedValue ? (
        <div className={product ? "grid gap-0" : "formula-result-shell"}>
          <span className={product ? "sr-only" : "formula-meta-label"}>Value</span>
          <div
            className={
              product
                ? "inline-flex min-h-8 items-center justify-end rounded-[2px] border border-[#dadce0] bg-white px-3 text-[13px] font-medium text-[#202124]"
                : "formula-result"
            }
            data-testid="formula-resolved-value"
          >
            {resolvedValue || "∅"}
          </div>
        </div>
      ) : (
        <span className="sr-only" data-testid="formula-resolved-value">
          {resolvedValue || "∅"}
        </span>
      )}
      {variant === "playground" ? (
        <div className="formula-actions">
          <button className="ghost-button" onClick={onCancel} type="button">
            Cancel
          </button>
          <button className="ghost-button" onClick={onClear} type="button">
            Clear
          </button>
          <button onClick={onCommit} type="button">
            Commit
          </button>
        </div>
      ) : null}
    </div>
  );
}
