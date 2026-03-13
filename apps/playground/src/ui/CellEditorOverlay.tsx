import React from "react";

interface CellEditorOverlayProps {
  label: string;
  value: string;
  onChange(next: string): void;
  onCommit(): void;
}

export function CellEditorOverlay({ label, value, onChange, onCommit }: CellEditorOverlayProps) {
  return (
    <div className="panel editor-panel">
      <span>{label}</span>
      <input
        value={value}
        onChange={(event) => onChange(event.target.value)}
        onKeyDown={(event) => {
          if (event.key === "Enter") onCommit();
        }}
      />
      <button onClick={onCommit}>Apply</button>
    </div>
  );
}
