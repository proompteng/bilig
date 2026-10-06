import { useRef, useState, type CSSProperties, type RefObject } from 'react'
import type { ColorSwatch } from './workbook-colors.js'
import { colorPickerSwatchClass } from './workbook-toolbar-theme.js'

export function ColorSwatchGrid({
  ariaLabel,
  currentColor,
  rows,
  focusRef,
  columnCount = 10,
  onSelect,
}: {
  readonly ariaLabel: string
  readonly currentColor: string
  readonly rows: readonly (readonly ColorSwatch[])[]
  readonly columnCount?: number | undefined
  readonly focusRef?: RefObject<HTMLButtonElement | null> | undefined
  readonly onSelect: (color: string) => void
}) {
  const [active, setActive] = useState(() => {
    const row = Math.max(
      0,
      rows.findIndex((entries) => entries.some((entry) => entry.value === currentColor)),
    )
    const col = Math.max(0, rows[row]?.findIndex((entry) => entry.value === currentColor) ?? 0)
    return { row, col }
  })
  const buttons = useRef(new Map<string, HTMLButtonElement>())

  return (
    <div aria-label={ariaLabel} className="space-y-1" role="grid">
      {rows.map((entries, row) => (
        <div
          className="grid gap-1"
          key={entries.map((entry) => entry.label).join('|')}
          role="row"
          style={{ gridTemplateColumns: `repeat(${columnCount}, minmax(0, 1fr))` } satisfies CSSProperties}
        >
          {entries.map((swatch, col) => (
            <div key={swatch.label} role="gridcell">
              <button
                aria-label={swatch.label}
                aria-pressed={swatch.value === currentColor}
                className={colorPickerSwatchClass()}
                data-color={swatch.value}
                type="button"
                ref={(node) => {
                  const key = `${row}:${col}`
                  if (node) buttons.current.set(key, node)
                  else buttons.current.delete(key)
                  if (focusRef && row === active.row && col === active.col) focusRef.current = node
                }}
                style={{ backgroundColor: swatch.value } satisfies CSSProperties}
                tabIndex={row === active.row && col === active.col ? 0 : -1}
                onClick={() => onSelect(swatch.value)}
                onFocus={() => setActive({ row, col })}
                onKeyDown={(event) => {
                  let nextRow = row
                  let nextCol = col
                  switch (event.key) {
                    case 'ArrowLeft':
                      nextCol -= 1
                      break
                    case 'ArrowRight':
                      nextCol += 1
                      break
                    case 'ArrowUp':
                      nextRow -= 1
                      break
                    case 'ArrowDown':
                      nextRow += 1
                      break
                    case 'Home':
                      if (event.ctrlKey || event.metaKey) nextRow = 0
                      nextCol = 0
                      break
                    case 'End':
                      if (event.ctrlKey || event.metaKey) nextRow = rows.length - 1
                      nextCol = (rows[nextRow]?.length ?? 1) - 1
                      break
                    default:
                      return
                  }
                  event.preventDefault()
                  event.stopPropagation()
                  nextRow = Math.max(0, Math.min(rows.length - 1, nextRow))
                  nextCol = Math.max(0, Math.min((rows[nextRow]?.length ?? 1) - 1, nextCol))
                  buttons.current.get(`${nextRow}:${nextCol}`)?.focus()
                }}
              />
            </div>
          ))}
        </div>
      ))}
    </div>
  )
}
