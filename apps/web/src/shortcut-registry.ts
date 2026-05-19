export interface WorkbookShortcutEntry {
  readonly id: string
  readonly label: string
  readonly category: 'Editing' | 'Formatting' | 'Navigation' | 'Selection' | 'Structure'
  readonly mac: string
  readonly windows: string
  readonly keywords?: readonly string[]
}

const WORKBOOK_SHORTCUTS: readonly WorkbookShortcutEntry[] = [
  {
    id: 'undo',
    label: 'Undo',
    category: 'Editing',
    mac: '⌘Z',
    windows: 'Ctrl+Z',
  },
  {
    id: 'redo',
    label: 'Redo',
    category: 'Editing',
    mac: '⇧⌘Z',
    windows: 'Ctrl+Y',
  },
  {
    id: 'begin-edit',
    label: 'Edit cell',
    category: 'Editing',
    mac: 'F2',
    windows: 'F2',
    keywords: ['formula', 'input'],
  },
  {
    id: 'commit-edit',
    label: 'Commit edit',
    category: 'Editing',
    mac: 'Enter',
    windows: 'Enter',
  },
  {
    id: 'insert-line-break',
    label: 'Insert line break',
    category: 'Editing',
    mac: '⌥Enter / ⌘Enter',
    windows: 'Alt+Enter / Ctrl+Enter',
    keywords: ['multiline', 'newline'],
  },
  {
    id: 'cancel-edit',
    label: 'Cancel edit',
    category: 'Editing',
    mac: 'Escape',
    windows: 'Escape',
  },
  {
    id: 'clear-selection',
    label: 'Clear selected cells',
    category: 'Editing',
    mac: 'Delete / Backspace',
    windows: 'Delete / Backspace',
    keywords: ['delete', 'backspace', 'clear cells'],
  },
  {
    id: 'copy',
    label: 'Copy',
    category: 'Editing',
    mac: '⌘C',
    windows: 'Ctrl+C',
  },
  {
    id: 'cut',
    label: 'Cut',
    category: 'Editing',
    mac: '⌘X',
    windows: 'Ctrl+X',
  },
  {
    id: 'paste',
    label: 'Paste',
    category: 'Editing',
    mac: '⌘V',
    windows: 'Ctrl+V',
  },
  {
    id: 'paste-values-only',
    label: 'Paste values only',
    category: 'Editing',
    mac: '⇧⌘V',
    windows: 'Ctrl+Shift+V',
    keywords: ['paste special', 'values only'],
  },
  {
    id: 'fill-down',
    label: 'Fill down',
    category: 'Editing',
    mac: '⌘D',
    windows: 'Ctrl+D',
    keywords: ['copy down', 'autofill'],
  },
  {
    id: 'fill-right',
    label: 'Fill right',
    category: 'Editing',
    mac: '⌘R',
    windows: 'Ctrl+R',
    keywords: ['copy right', 'autofill'],
  },
  {
    id: 'fill-range',
    label: 'Fill range',
    category: 'Editing',
    mac: '⌘Enter',
    windows: 'Ctrl+Enter',
    keywords: ['copy selected range', 'autofill'],
  },
  {
    id: 'bold',
    label: 'Bold',
    category: 'Formatting',
    mac: '⌘B',
    windows: 'Ctrl+B',
  },
  {
    id: 'italic',
    label: 'Italic',
    category: 'Formatting',
    mac: '⌘I',
    windows: 'Ctrl+I',
  },
  {
    id: 'underline',
    label: 'Underline',
    category: 'Formatting',
    mac: '⌘U',
    windows: 'Ctrl+U',
  },
  {
    id: 'format-number',
    label: 'Number format',
    category: 'Formatting',
    mac: '⇧⌘1',
    windows: 'Ctrl+Shift+1',
  },
  {
    id: 'format-currency',
    label: 'Currency format',
    category: 'Formatting',
    mac: '⇧⌘4',
    windows: 'Ctrl+Shift+4',
  },
  {
    id: 'format-percent',
    label: 'Percent format',
    category: 'Formatting',
    mac: '⇧⌘5',
    windows: 'Ctrl+Shift+5',
  },
  {
    id: 'align-left',
    label: 'Align left',
    category: 'Formatting',
    mac: '⇧⌘L',
    windows: 'Ctrl+Shift+L',
  },
  {
    id: 'align-center',
    label: 'Align center',
    category: 'Formatting',
    mac: '⇧⌘E',
    windows: 'Ctrl+Shift+E',
  },
  {
    id: 'align-right',
    label: 'Align right',
    category: 'Formatting',
    mac: '⇧⌘R',
    windows: 'Ctrl+Shift+R',
  },
  {
    id: 'border-outer',
    label: 'Outer borders',
    category: 'Formatting',
    mac: '⇧⌘7',
    windows: 'Ctrl+Shift+7',
    keywords: ['border'],
  },
  {
    id: 'clear-formatting',
    label: 'Clear formatting',
    category: 'Formatting',
    mac: '⌘\\',
    windows: 'Ctrl+\\',
    keywords: ['clear style'],
  },
  {
    id: 'context-menu',
    label: 'Open context menu',
    category: 'Structure',
    mac: '⇧⌘\\',
    windows: 'Ctrl+Shift+\\',
    keywords: ['right click', 'row menu', 'column menu'],
  },
  {
    id: 'delete-selected-structure',
    label: 'Delete selected rows or columns',
    category: 'Structure',
    mac: '⌘⌥-',
    windows: 'Ctrl+Alt+-',
    keywords: ['delete row', 'delete column', 'remove rows', 'remove columns'],
  },
  {
    id: 'move-selection',
    label: 'Move selection',
    category: 'Navigation',
    mac: 'Arrow keys',
    windows: 'Arrow keys',
    keywords: ['cell', 'navigate'],
  },
  {
    id: 'extend-selection',
    label: 'Extend selection',
    category: 'Selection',
    mac: 'Shift+Arrow keys',
    windows: 'Shift+Arrow keys',
    keywords: ['range'],
  },
  {
    id: 'goto-range',
    label: 'Go to cell or range',
    category: 'Navigation',
    mac: '⌘G',
    windows: 'Ctrl+G',
    keywords: ['name box', 'jump', 'range'],
  },
  {
    id: 'next-sheet',
    label: 'Move to next sheet',
    category: 'Navigation',
    mac: '⌥↓',
    windows: 'Alt+Down',
    keywords: ['tab', 'worksheet'],
  },
  {
    id: 'previous-sheet',
    label: 'Move to previous sheet',
    category: 'Navigation',
    mac: '⌥↑',
    windows: 'Alt+Up',
    keywords: ['tab', 'worksheet'],
  },
  {
    id: 'jump-row-start',
    label: 'Jump to row start',
    category: 'Navigation',
    mac: 'Home',
    windows: 'Home',
  },
  {
    id: 'jump-sheet-edge',
    label: 'Jump to sheet edge',
    category: 'Navigation',
    mac: '⌘Arrow',
    windows: 'Ctrl+Arrow',
  },
  {
    id: 'select-row',
    label: 'Select row',
    category: 'Selection',
    mac: 'Shift+Space',
    windows: 'Shift+Space',
  },
  {
    id: 'select-column',
    label: 'Select column',
    category: 'Selection',
    mac: '⌃Space',
    windows: 'Ctrl+Space',
  },
  {
    id: 'select-current-region',
    label: 'Select current region',
    category: 'Selection',
    mac: '⇧⌘*',
    windows: 'Ctrl+Shift+*',
    keywords: ['data region', 'table', 'range'],
  },
  {
    id: 'select-all',
    label: 'Select all',
    category: 'Selection',
    mac: '⌘A',
    windows: 'Ctrl+A',
    keywords: ['sheet'],
  },
  {
    id: 'toggle-shortcuts',
    label: 'Show shortcuts',
    category: 'Structure',
    mac: '⌘/',
    windows: 'Ctrl+/',
    keywords: ['?', 'help', 'keyboard'],
  },
] as const

const SHORTCUT_BY_ID = new Map(WORKBOOK_SHORTCUTS.map((entry) => [entry.id, entry]))

function normalizeSearchText(value: string): string {
  return value.trim().toLowerCase()
}

function isMacPlatform(platform?: string): boolean {
  if (!platform) {
    return false
  }
  return /mac|iphone|ipad|ipod/i.test(platform)
}

function getWorkbookShortcutEntry(id: string): WorkbookShortcutEntry | undefined {
  return SHORTCUT_BY_ID.get(id)
}

export function getWorkbookShortcutLabel(id: string, platform = globalThis.navigator?.platform): string {
  const entry = getWorkbookShortcutEntry(id)
  if (!entry) {
    return ''
  }
  return isMacPlatform(platform) ? entry.mac : entry.windows
}

function splitMacShortcutLabel(label: string): readonly string[] {
  const parts: string[] = []
  let trailing = ''
  for (const char of label) {
    if (char === '⌘' || char === '⇧' || char === '⌥' || char === '⌃') {
      parts.push(char)
      continue
    }
    trailing += char
  }
  if (trailing.length > 0) {
    parts.push(trailing)
  }
  return parts
}

export function getWorkbookShortcutParts(id: string, platform = globalThis.navigator?.platform): readonly string[] {
  const label = getWorkbookShortcutLabel(id, platform)
  if (label.length === 0) {
    return []
  }
  if (label.includes('+')) {
    return label.split('+').map((part) => part.trim())
  }
  if (/[⌘⇧⌥⌃]/.test(label)) {
    return splitMacShortcutLabel(label)
  }
  return [label]
}

export function searchWorkbookShortcutEntries(query: string): readonly WorkbookShortcutEntry[] {
  const normalizedQuery = normalizeSearchText(query)
  if (normalizedQuery.length === 0) {
    return WORKBOOK_SHORTCUTS
  }
  return WORKBOOK_SHORTCUTS.filter((entry) => {
    const searchParts = [entry.label, entry.category, entry.mac, entry.windows, ...(entry.keywords ?? [])]
    return searchParts.some((part) => normalizeSearchText(part).includes(normalizedQuery))
  })
}

export function groupWorkbookShortcutEntries(entries: readonly WorkbookShortcutEntry[]): readonly {
  readonly category: WorkbookShortcutEntry['category']
  readonly entries: readonly WorkbookShortcutEntry[]
}[] {
  const order: readonly WorkbookShortcutEntry['category'][] = ['Editing', 'Formatting', 'Navigation', 'Selection', 'Structure']
  return order
    .map((category) => ({
      category,
      entries: entries.filter((entry) => entry.category === category),
    }))
    .filter((group) => group.entries.length > 0)
}
