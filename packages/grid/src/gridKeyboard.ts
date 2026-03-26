import type { GridKeyEventArgs } from "@glideapps/glide-data-grid";

export function isPrintableKey(
  event: Pick<GridKeyEventArgs, "altKey" | "ctrlKey" | "key" | "metaKey">,
): boolean {
  if (event.ctrlKey || event.metaKey || event.altKey) {
    return false;
  }
  return event.key.length === 1;
}

export function normalizeKeyboardKey(key: string, code?: string): string {
  if (code?.startsWith("Numpad")) {
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
  }
  return key;
}

export function isNumericEditorSeed(value: string): boolean {
  const normalized = value.trim();
  if (normalized.length === 0 || normalized.startsWith("=")) {
    return false;
  }
  return /^-?\d+(\.\d+)?$/.test(normalized);
}

export function isNavigationKey(key: string): boolean {
  return key === "ArrowUp" || key === "ArrowDown" || key === "ArrowLeft" || key === "ArrowRight";
}

export function isClipboardShortcut(
  event: Pick<GridKeyEventArgs, "altKey" | "ctrlKey" | "key" | "metaKey">,
): boolean {
  if (!(event.ctrlKey || event.metaKey) || event.altKey) {
    return false;
  }
  const normalizedKey = event.key.toLowerCase();
  return normalizedKey === "c" || normalizedKey === "x" || normalizedKey === "v";
}

export function isHandledGridKey(
  event: Pick<GridKeyEventArgs, "altKey" | "ctrlKey" | "key" | "metaKey">,
): boolean {
  return (
    isPrintableKey(event) ||
    isClipboardShortcut(event) ||
    isNavigationKey(event.key) ||
    event.key === "Enter" ||
    event.key === "Tab" ||
    event.key === "Escape" ||
    event.key === "F2" ||
    event.key === "Backspace" ||
    event.key === "Delete"
  );
}
