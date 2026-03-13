import React from "react";

interface FormulaBarProps {
  value: string;
  onChange(next: string): void;
  onCommit(): void;
}

export function FormulaBar({ value, onChange, onCommit }: FormulaBarProps) {
  return (
    <div className="panel formula-bar">
      <label htmlFor="formula-input">Formula</label>
      <input
        id="formula-input"
        value={value}
        onChange={(event) => onChange(event.target.value)}
        onKeyDown={(event) => {
          if (event.key === "Enter") onCommit();
        }}
      />
      <button onClick={onCommit}>Commit</button>
    </div>
  );
}
