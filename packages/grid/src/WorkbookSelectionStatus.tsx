import React from 'react'
import { Select } from '@base-ui/react/select'
import type { GridSelectionSnapshot } from './gridTypes.js'
import type { SelectionAggregateSummary } from './selectionAggregateSummary.js'

type SelectionAggregateMetric = 'sum' | 'avg' | 'min' | 'max' | 'count' | 'countNumbers'

interface WorkbookSelectionStatusProps {
  selectionLabel: string
  selectionSnapshot: GridSelectionSnapshot
  summary: SelectionAggregateSummary | null
}

interface SelectionAggregateOption {
  readonly metric: SelectionAggregateMetric
  readonly label: string
  readonly value: SelectionAggregateMetric
  readonly valueText: string
}

const numberFormatter = new Intl.NumberFormat('en-US', { maximumSignificantDigits: 15 })

function buildSelectionAggregateOptions(summary: SelectionAggregateSummary): readonly SelectionAggregateOption[] {
  if (summary.nonEmptyCount === 0) return []
  const values: ReadonlyArray<{ metric: SelectionAggregateMetric; label: string; value: number | null }> = [
    { metric: 'sum', label: 'Sum', value: summary.numericCount ? summary.sum : null },
    { metric: 'avg', label: 'Avg', value: summary.numericCount ? summary.sum / summary.numericCount : null },
    { metric: 'min', label: 'Min', value: summary.min },
    { metric: 'max', label: 'Max', value: summary.max },
    { metric: 'count', label: 'Count', value: summary.nonEmptyCount },
    { metric: 'countNumbers', label: 'Count Numbers', value: summary.numericCount },
  ]
  return values.flatMap(({ metric, label, value }) =>
    value === null ? [] : [{ metric, label, value: metric, valueText: numberFormatter.format(value) }],
  )
}

const statusTriggerClass =
  'inline-flex h-8 items-center gap-2 rounded-[var(--wb-radius-control)] border border-transparent bg-transparent px-3 text-[12px] font-medium text-[var(--wb-text)] transition-[background-color,border-color,color,box-shadow] hover:border-[var(--wb-border-strong)] hover:bg-[var(--wb-surface)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1 focus-visible:ring-offset-[var(--wb-surface-subtle)]'

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

export function WorkbookSelectionStatus({ selectionLabel, selectionSnapshot, summary }: WorkbookSelectionStatusProps) {
  const [selectedMetric, setSelectedMetric] = React.useState<SelectionAggregateMetric>('sum')
  const options = summary ? buildSelectionAggregateOptions(summary) : []
  const activeOption = options.find((option) => option.metric === selectedMetric) ?? options[0] ?? null

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
