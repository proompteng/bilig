interface GridKeyboardModifierState {
  altKey: boolean
  ctrlKey: boolean
  key: string
  metaKey: boolean
}

export function isPrintableKey(event: GridKeyboardModifierState): boolean {
  if (event.ctrlKey || event.metaKey || event.altKey) {
    return false
  }
  return event.key.length === 1
}

export function normalizeKeyboardKey(key: string, code?: string): string {
  if (code?.startsWith('Numpad')) {
    const suffix = code.slice('Numpad'.length)
    if (/^\d$/.test(suffix)) {
      return suffix
    }
    if (suffix === 'Decimal') {
      return '.'
    }
    if (suffix === 'Add') {
      return '+'
    }
    if (suffix === 'Subtract') {
      return '-'
    }
    if (suffix === 'Multiply') {
      return '*'
    }
    if (suffix === 'Divide') {
      return '/'
    }
  }
  return key
}

export function isNumericEditorSeed(value: string): boolean {
  const normalized = value.trim()
  if (normalized.length === 0 || normalized.startsWith('=')) {
    return false
  }
  return /^-?\d+(\.\d+)?$/.test(normalized)
}

export function isNavigationKey(key: string): boolean {
  return key === 'ArrowUp' || key === 'ArrowDown' || key === 'ArrowLeft' || key === 'ArrowRight'
}

export function isNavigationShortcut(
  event: GridKeyboardModifierState & {
    shiftKey?: boolean
  },
): boolean {
  return isNavigationKey(event.key) && !event.altKey
}

export function isClipboardShortcut(event: GridKeyboardModifierState): boolean {
  if (!(event.ctrlKey || event.metaKey) || event.altKey) {
    return false
  }
  const normalizedKey = event.key.toLowerCase()
  return normalizedKey === 'c' || normalizedKey === 'x' || normalizedKey === 'v'
}

export function isDeleteKey(key: string): boolean {
  return key === 'Backspace' || key === 'Delete'
}

export function isClearCellKey(
  event: GridKeyboardModifierState & {
    shiftKey?: boolean
  },
): boolean {
  return isDeleteKey(event.key) && !event.altKey && !event.ctrlKey && !event.metaKey && !event.shiftKey
}

export function isHandledGridKey(
  event: GridKeyboardModifierState & {
    shiftKey?: boolean
  },
): boolean {
  const hasPrimaryModifier = event.ctrlKey || event.metaKey
  return (
    isPrintableKey(event) ||
    isClipboardShortcut(event) ||
    isNavigationShortcut(event) ||
    (hasPrimaryModifier && !event.altKey && event.key.toLowerCase() === 'a') ||
    (event.key === ' ' && !event.altKey && (hasPrimaryModifier || event.shiftKey)) ||
    (event.key === 'Enter' && !event.altKey && !hasPrimaryModifier) ||
    (event.key === 'Tab' && !event.altKey && !hasPrimaryModifier) ||
    (event.key === 'Escape' && !event.altKey && !hasPrimaryModifier) ||
    (event.key === 'F2' && !event.altKey && !hasPrimaryModifier && !event.shiftKey) ||
    isClearCellKey(event) ||
    ((event.key === 'Home' || event.key === 'End' || event.key === 'PageUp' || event.key === 'PageDown') && !event.altKey)
  )
}
