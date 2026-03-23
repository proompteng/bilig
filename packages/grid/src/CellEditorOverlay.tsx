import { useEffect, useRef, type CSSProperties } from "react";
import type { EditMovement } from "./SheetGridView.js";

function normalizeNumpadKey(key: string, code: string): string | null {
  if (!code.startsWith("Numpad")) {
    return null;
  }
  const suffix = code.slice("Numpad".length);
  if (/^\d$/.test(suffix)) {
    return suffix;
  }
  if (suffix === "Decimal") {
    return ".";
  }
  if (suffix === "Add") {
    return "+";
  }
  if (suffix === "Subtract") {
    return "-";
  }
  if (suffix === "Multiply") {
    return "*";
  }
  if (suffix === "Divide") {
    return "/";
  }
  return key.length === 1 ? key : null;
}

interface CellEditorOverlayProps {
  label: string;
  value: string;
  resolvedValue: string;
  selectionBehavior?: "select-all" | "caret-end";
  textAlign?: "left" | "right";
  onChange(this: void, next: string): void;
  onCommit(this: void, movement?: EditMovement): void;
  onCancel(this: void): void;
  style?: CSSProperties;
}

export function CellEditorOverlay({
  label,
  value,
  resolvedValue: _resolvedValue,
  selectionBehavior = "select-all",
  textAlign = "left",
  onChange,
  onCommit,
  onCancel,
  style,
}: CellEditorOverlayProps) {
  const inputRef = useRef<HTMLInputElement | null>(null);
  const completionRef = useRef<"idle" | "commit" | "cancel">("idle");
  const blurArmedRef = useRef(false);

  useEffect(() => {
    blurArmedRef.current = false;
    inputRef.current?.focus();
    if (selectionBehavior === "select-all") {
      inputRef.current?.select();
    } else {
      const caretPosition = value.length;
      inputRef.current?.setSelectionRange(caretPosition, caretPosition);
    }
    const blurArm = window.requestAnimationFrame(() => {
      blurArmedRef.current = true;
    });

    return () => {
      window.cancelAnimationFrame(blurArm);
    };
  }, [selectionBehavior, value]);

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
      <input
        aria-label={`${label} editor`}
        data-testid="cell-editor-input"
        ref={inputRef}
        style={{ textAlign }}
        value={value}
        onBlur={() => {
          if (!blurArmedRef.current) {
            return;
          }
          commit();
        }}
        onChange={(event) => onChange(event.target.value)}
        onKeyDown={(event) => {
          const normalizedNumpadKey = normalizeNumpadKey(event.key, event.code);
          if (normalizedNumpadKey !== null && event.key !== normalizedNumpadKey) {
            event.preventDefault();
            const input = event.currentTarget;
            const selectionStart = input.selectionStart ?? value.length;
            const selectionEnd = input.selectionEnd ?? value.length;
            const nextValue = `${value.slice(0, selectionStart)}${normalizedNumpadKey}${value.slice(selectionEnd)}`;
            onChange(nextValue);
            window.requestAnimationFrame(() => {
              const caretPosition = selectionStart + normalizedNumpadKey.length;
              inputRef.current?.setSelectionRange(caretPosition, caretPosition);
            });
            return;
          }
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
    </div>
  );
}
