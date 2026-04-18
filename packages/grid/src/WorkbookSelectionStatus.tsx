import React from 'react'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import { ValueTag } from '@bilig/protocol'
import { Select } from '@base-ui/react/select'
import type { GridEngineLike } from './grid-engine.js'
import type { GridSelectionSnapshot } from './gridTypes.js'

type SelectionAggregateMetric = 'sum' | 'avg' | 'min' | 'max' | 'count' | 'countNumbers'

interface WorkbookSelectionStatusProps {
  engine: GridEngineLike
  selectionLabel: string
  selectionSnapshot: GridSelectionSnapshot
}

interface SelectionAggregateSummary {
  readonly materializedAddresses: readonly string[]
  readonly nonEmptyCount: number
  readonly numericValues: readonly number[]
}

interface SelectionAggregateOption {
  readonly metric: SelectionAggregateMetric
  readonly label: string
  readonly value: SelectionAggregateMetric
  readonly valueText: string
}

const decimalFormatter = new Intl.NumberFormat('en-US', {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
})

const integerFormatter = new Intl.NumberFormat('en-US')

function isCellIncludedInSelection(
  selection: GridSelectionSnapshot,
  row: number,
  col: number,
  bounds: { rowStart: number; rowEnd: number; colStart: number; colEnd: number },
): boolean {
  switch (selection.kind) {
    case 'sheet':
      return true
    case 'row':
      return row >= bounds.rowStart && row <= bounds.rowEnd
    case 'column':
      return col >= bounds.colStart && col <= bounds.colEnd
    case 'range':
    case 'cell':
      return row >= bounds.rowStart && row <= bounds.rowEnd && col >= bounds.colStart && col <= bounds.colEnd
  }
}

function collectSelectionAggregateSummary(engine: GridEngineLike, selection: GridSelectionSnapshot): SelectionAggregateSummary {
  const sheet = engine.workbook.getSheet(selection.sheetName)
  if (!sheet) {
    return {
      materializedAddresses: [],
      nonEmptyCount: 0,
      numericValues: [],
    }
  }

  const start = parseCellAddress(selection.range.startAddress, selection.sheetName)
  const end = parseCellAddress(selection.range.endAddress, selection.sheetName)
  const bounds = {
    rowStart: Math.min(start.row, end.row),
    rowEnd: Math.max(start.row, end.row),
    colStart: Math.min(start.col, end.col),
    colEnd: Math.max(start.col, end.col),
  }

  const materializedAddresses: string[] = []
  const numericValues: number[] = []
  let nonEmptyCount = 0

  sheet.grid.forEachCellEntry((_cellIndex, row, col) => {
    if (!isCellIncludedInSelection(selection, row, col, bounds)) {
      return
    }
    const address = formatAddress(row, col)
    materializedAddresses.push(address)
    const snapshot = engine.getCell(selection.sheetName, address)
    if (snapshot.value.tag !== ValueTag.Empty) {
      nonEmptyCount += 1
    }
    if (snapshot.value.tag === ValueTag.Number && Number.isFinite(snapshot.value.value)) {
      numericValues.push(snapshot.value.value)
    }
  })

  return {
    materializedAddresses,
    nonEmptyCount,
    numericValues,
  }
}

function formatMetricValue(metric: SelectionAggregateMetric, summary: SelectionAggregateSummary): string {
  switch (metric) {
    case 'sum':
      return decimalFormatter.format(summary.numericValues.reduce((total, value) => total + value, 0))
    case 'avg':
      return decimalFormatter.format(summary.numericValues.reduce((total, value) => total + value, 0) / summary.numericValues.length)
    case 'min':
      return decimalFormatter.format(Math.min(...summary.numericValues))
    case 'max':
      return decimalFormatter.format(Math.max(...summary.numericValues))
    case 'count':
      return integerFormatter.format(summary.nonEmptyCount)
    case 'countNumbers':
      return integerFormatter.format(summary.numericValues.length)
  }
}

function buildSelectionAggregateOptions(summary: SelectionAggregateSummary): readonly SelectionAggregateOption[] {
  if (summary.numericValues.length === 0) {
    return []
  }
  const definitions: ReadonlyArray<{ metric: SelectionAggregateMetric; label: string }> = [
    { metric: 'sum', label: 'Sum' },
    { metric: 'avg', label: 'Avg' },
    { metric: 'min', label: 'Min' },
    { metric: 'max', label: 'Max' },
    { metric: 'count', label: 'Count' },
    { metric: 'countNumbers', label: 'Count Numbers' },
  ]
  return definitions.map((definition) => ({
    metric: definition.metric,
    label: definition.label,
    value: definition.metric,
    valueText: formatMetricValue(definition.metric, summary),
  }))
}

const statusTriggerClass =
  'inline-flex h-8 items-center gap-2 rounded-full border border-[var(--wb-border)] bg-[var(--wb-muted)] px-3 text-[12px] font-medium text-[var(--wb-text)] shadow-[var(--wb-shadow-sm)] transition-[background-color,border-color,color,box-shadow] hover:border-[var(--wb-border-strong)] hover:bg-[var(--wb-surface)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1 focus-visible:ring-offset-[var(--wb-surface-subtle)]'

const statusMenuClass =
  'w-max min-w-[220px] max-w-[calc(100vw-2rem)] overflow-hidden rounded-[var(--wb-radius-panel)] border border-[var(--wb-border)] bg-[var(--wb-surface)] p-1 shadow-[var(--wb-shadow-md)] outline-none'

const statusMenuItemClass =
  'flex items-center gap-2 whitespace-nowrap rounded-[var(--wb-radius-control)] px-2.5 py-2 text-[12px] text-[var(--wb-text)] outline-none transition-colors data-[highlighted]:bg-[var(--wb-muted)] data-[selected]:font-semibold'

const statusMenuItemIndicatorSlotClass = 'inline-flex h-4 w-4 shrink-0 items-center justify-center'

const statusMenuItemIndicatorClass = 'text-[var(--wb-accent)]'

const statusMenuItemTextClass = 'whitespace-nowrap leading-none'

function SelectionStatusChevronIcon() {
  return (
    <svg aria-hidden="true" className="h-3.5 w-3.5" fill="none" viewBox="0 0 16 16">
      <path d="M4 6.25 8 10l4-3.75" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.75" />
    </svg>
  )
}

function SelectionStatusCheckIcon() {
  return (
    <svg aria-hidden="true" className="h-3.5 w-3.5" fill="none" viewBox="0 0 16 16">
      <path d="m3.5 8.25 2.25 2.25 6-6" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" />
    </svg>
  )
}

export function WorkbookSelectionStatus({ engine, selectionLabel, selectionSnapshot }: WorkbookSelectionStatusProps) {
  const [selectedMetric, setSelectedMetric] = React.useState<SelectionAggregateMetric>('sum')
  const [, setRevision] = React.useState(0)

  const summary = collectSelectionAggregateSummary(engine, selectionSnapshot)

  const addressesKey = React.useMemo(() => summary.materializedAddresses.join('\u001f'), [summary.materializedAddresses])

  React.useEffect(() => {
    if (summary.materializedAddresses.length === 0) {
      return
    }
    return engine.subscribeCells(selectionSnapshot.sheetName, summary.materializedAddresses, () => {
      setRevision((current) => current + 1)
    })
  }, [addressesKey, engine, selectionSnapshot.sheetName, summary.materializedAddresses])

  const options = React.useMemo(() => buildSelectionAggregateOptions(summary), [summary])
  const activeOption = React.useMemo(
    () => options.find((option) => option.metric === selectedMetric) ?? options[0] ?? null,
    [options, selectedMetric],
  )

  if (selectionSnapshot.kind === 'cell' || activeOption === null) {
    return <span data-testid="workbook-selection-summary">{selectionLabel}</span>
  }

  return (
    <div data-testid="workbook-selection-summary">
      <Select.Root
        items={options}
        value={activeOption.metric}
        onValueChange={(nextValue: string | null) => {
          if (
            nextValue === 'sum' ||
            nextValue === 'avg' ||
            nextValue === 'min' ||
            nextValue === 'max' ||
            nextValue === 'count' ||
            nextValue === 'countNumbers'
          ) {
            setSelectedMetric(nextValue)
          }
        }}
      >
        <Select.Trigger aria-label="Selection calculations" className={statusTriggerClass} data-testid="workbook-selection-status-trigger">
          <span className="whitespace-nowrap">{`${activeOption.label}: ${activeOption.valueText}`}</span>
          <Select.Icon className="text-[var(--wb-text-muted)]">
            <SelectionStatusChevronIcon />
          </Select.Icon>
        </Select.Trigger>
        <Select.Portal>
          <Select.Positioner align="end" className="z-[1200]" side="top" sideOffset={8}>
            <Select.Popup className={statusMenuClass} data-testid="workbook-selection-status-menu">
              <Select.List className="py-1">
                {options.map((option) => (
                  <Select.Item
                    className={statusMenuItemClass}
                    key={option.metric}
                    label={option.label}
                    value={option.metric}
                    data-testid={`workbook-selection-status-option-${option.metric}`}
                    onClick={() => {
                      setSelectedMetric(option.metric)
                    }}
                  >
                    <span
                      className={statusMenuItemIndicatorSlotClass}
                      data-testid={`workbook-selection-status-option-${option.metric}-indicator-slot`}
                    >
                      <Select.ItemIndicator className={statusMenuItemIndicatorClass}>
                        <SelectionStatusCheckIcon />
                      </Select.ItemIndicator>
                    </span>
                    <Select.ItemText
                      className={statusMenuItemTextClass}
                      data-testid={`workbook-selection-status-option-${option.metric}-text`}
                    >
                      {`${option.label}: ${option.valueText}`}
                    </Select.ItemText>
                  </Select.Item>
                ))}
              </Select.List>
            </Select.Popup>
          </Select.Positioner>
        </Select.Portal>
      </Select.Root>
    </div>
  )
}
