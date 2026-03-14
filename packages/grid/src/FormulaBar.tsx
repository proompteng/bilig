import React from "react";

interface FormulaBarProps {
  label: string;
  value: string;
  onChange(next: string): void;
  onCommit(): void;
  onClear(): void;
}

export function FormulaBar({ label, value, onChange, onCommit, onClear }: FormulaBarProps) {
  return (
    <div className="panel formula-bar">
      <div className="formula-copy">
        <p className="panel-eyebrow">Formula Bar</p>
        <label htmlFor="formula-input">{label}</label>
      </div>
      <input
        aria-label="Formula"
        id="formula-input"
        placeholder="Type a literal or =formula"
        value={value}
        onChange={(event) => onChange(event.target.value)}
        onKeyDown={(event) => {
          if (event.key === "Enter") onCommit();
        }}
      />
      <div className="formula-actions">
        <button className="ghost-button" onClick={onClear} type="button">Clear</button>
        <button onClick={onCommit} type="button">Commit</button>
      </div>
    </div>
  );
}
