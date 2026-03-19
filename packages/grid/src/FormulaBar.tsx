import React, { useEffect, useState } from "react";

interface FormulaBarProps {
  variant?: "playground" | "product";
  sheetName: string;
  address: string;
  value: string;
  resolvedValue: string;
  isEditing: boolean;
  onBeginEdit(seed?: string): void;
  onAddressCommit(next: string): void;
  onChange(next: string): void;
  onCommit(): void;
  onCancel(): void;
  onClear(): void;
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
  onClear
}: FormulaBarProps) {
  const [addressValue, setAddressValue] = useState(address);

  useEffect(() => {
    setAddressValue(address);
  }, [address, sheetName]);

  const showResolvedValue = variant !== "product";

  return (
    <div className="formula-bar" data-testid="formula-bar">
      <div className="name-box-shell">
        <label className="formula-meta-label" htmlFor="name-box-input">
          Name
        </label>
        <input
          aria-label="Name box"
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
      <div className="formula-input-shell">
        <label className="formula-meta-label" htmlFor="formula-input">
          Formula
        </label>
        <div className="formula-input-frame">
          <span aria-hidden="true" className="formula-fx">
            fx
          </span>
          <input
            aria-label="Formula"
            data-testid="formula-input"
            id="formula-input"
            placeholder="Type a literal or =formula"
            value={value}
            onBlur={(event) => {
              const nextTarget = event.relatedTarget;
              if (nextTarget instanceof Node && event.currentTarget.closest(".formula-bar")?.contains(nextTarget)) {
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
        <div className="formula-result-shell">
          <span className="formula-meta-label">Value</span>
          <div className="formula-result" data-testid="formula-resolved-value">
            {resolvedValue || "∅"}
          </div>
        </div>
      ) : (
        <span className="visually-hidden" data-testid="formula-resolved-value">
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
