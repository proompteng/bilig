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
    id: 'cancel-edit',
    label: 'Cancel edit',
    category: 'Editing',
    mac: 'Escape',
    windows: 'Escape',
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
    id: 'move-selection',
    label: 'Move selection',
    category: 'Navigation',
    mac: 'Arrow keys',
    windows: 'Arrow keys',
    keywords: ['cell', 'navigate'],
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
    mac: '?',
    windows: '?',
    keywords: ['help', 'keyboard'],
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
