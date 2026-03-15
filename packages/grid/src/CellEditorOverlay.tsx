import React, { useEffect, useRef, type CSSProperties } from "react";
import type { EditMovement } from "./SheetGridView.js";

interface CellEditorOverlayProps {
  label: string;
  value: string;
  resolvedValue: string;
  onChange(next: string): void;
  onCommit(movement?: EditMovement): void;
  onCancel(): void;
  style?: CSSProperties;
}

export function CellEditorOverlay({
  label,
  value,
  resolvedValue,
  onChange,
  onCommit,
  onCancel,
  style
}: CellEditorOverlayProps) {
  const inputRef = useRef<HTMLInputElement | null>(null);
  const completionRef = useRef<"idle" | "commit" | "cancel">("idle");
  const blurArmedRef = useRef(false);

  useEffect(() => {
    blurArmedRef.current = false;
    inputRef.current?.focus();
    inputRef.current?.select();
    const blurArm = window.requestAnimationFrame(() => {
      blurArmedRef.current = true;
    });

    return () => {
      window.cancelAnimationFrame(blurArm);
    };
  }, []);

  const commit = (movement?: EditMovement) => {
    if (completionRef.current !== "idle") {
      return;
    }
    completionRef.current = "commit";
    onCommit(movement);
  };

  const cancel = () => {
    if (completionRef.current !== "idle") {
      return;
    }
    completionRef.current = "cancel";
    onCancel();
  };

  return (
    <div className="cell-editor-overlay" data-testid="cell-editor-overlay" style={style}>
      <div className="cell-editor-overlay-head">
        <div className="cell-editor-overlay-copy">
          <p className="panel-eyebrow">Editing</p>
          <strong>{label}</strong>
        </div>
        <span className="cell-editor-overlay-value">Current: {resolvedValue || "∅"}</span>
      </div>
      <input
        aria-label={`${label} editor`}
        ref={inputRef}
        value={value}
        onBlur={() => {
          if (!blurArmedRef.current) {
            return;
          }
          commit();
        }}
        onChange={(event) => onChange(event.target.value)}
        onKeyDown={(event) => {
          if (event.key === "Enter") {
            event.preventDefault();
            commit([0, event.shiftKey ? -1 : 1]);
            return;
          }
          if (event.key === "Tab") {
            event.preventDefault();
            commit([event.shiftKey ? -1 : 1, 0]);
            return;
          }
          if (event.key === "Escape") {
            event.preventDefault();
            cancel();
          }
        }}
      />
      <div className="cell-editor-overlay-hint">
        <span>Enter or blur to commit</span>
        <span>Esc to cancel</span>
      </div>
    </div>
  );
}
