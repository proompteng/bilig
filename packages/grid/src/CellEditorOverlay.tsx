import React from "react";

interface CellEditorOverlayProps {
  label: string;
  value: string;
  resolvedValue: string;
  onChange(next: string): void;
  onCommit(): void;
  onClear(): void;
}

export function CellEditorOverlay({ label, value, resolvedValue, onChange, onCommit, onClear }: CellEditorOverlayProps) {
  return (
    <div className="panel editor-panel">
      <div className="editor-copy">
        <p className="panel-eyebrow">Selected Cell</p>
        <strong>{label}</strong>
        <span className="editor-value">Current value: {resolvedValue || "∅"}</span>
      </div>
      <input
        aria-label={`${label} editor`}
        value={value}
        onChange={(event) => onChange(event.target.value)}
        onKeyDown={(event) => {
          if (event.key === "Enter") onCommit();
        }}
      />
      <div className="editor-actions">
        <button className="ghost-button" onClick={onClear} type="button">Clear</button>
        <button onClick={onCommit} type="button">Apply</button>
      </div>
    </div>
  );
}
