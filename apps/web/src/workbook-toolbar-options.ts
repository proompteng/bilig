import { TableCellsMerge, TableCellsSplit, type LucideIcon } from 'lucide-react'
import {
  BorderAllIcon,
  BorderBottomIcon,
  BorderFullIcon,
  BorderLeftIcon,
  BorderNoneIcon,
  BorderRightIcon,
  BorderTopIcon,
} from '@hugeicons/core-free-icons'
import type { IconSvgElement } from '@hugeicons/react'

export type BorderPreset = 'all' | 'outer' | 'left' | 'top' | 'right' | 'bottom' | 'clear'

export interface BorderPresetOption {
  readonly key: BorderPreset
  readonly icon: IconSvgElement
  readonly label: string
  readonly shortLabel: string
}

export const BORDER_PRESET_OPTIONS: readonly BorderPresetOption[] = [
  { key: 'all', label: 'All borders', shortLabel: 'All', icon: BorderAllIcon },
  { key: 'outer', label: 'Outer borders', shortLabel: 'Outer', icon: BorderFullIcon },
  { key: 'left', label: 'Left border', shortLabel: 'Left', icon: BorderLeftIcon },
  { key: 'top', label: 'Top border', shortLabel: 'Top', icon: BorderTopIcon },
  { key: 'right', label: 'Right border', shortLabel: 'Right', icon: BorderRightIcon },
  { key: 'bottom', label: 'Bottom border', shortLabel: 'Bottom', icon: BorderBottomIcon },
  { key: 'clear', label: 'Clear borders', shortLabel: 'Clear', icon: BorderNoneIcon },
] as const

export type StructureActionTemplate =
  | 'mergeSelectedCells'
  | 'unmergeSelectedCells'
  | 'hideCurrentRow'
  | 'unhideCurrentRow'
  | 'hideCurrentColumn'
  | 'unhideCurrentColumn'

export interface StructureActionOption {
  readonly key: string
  readonly label: string
  readonly template: StructureActionTemplate
  readonly icon?: LucideIcon
}

export const STRUCTURE_ACTIONS: readonly StructureActionOption[] = [
  { key: 'merge-cells', label: 'Merge cells', template: 'mergeSelectedCells', icon: TableCellsMerge },
  { key: 'unmerge-cells', label: 'Unmerge cells', template: 'unmergeSelectedCells', icon: TableCellsSplit },
  { key: 'hide-row', label: 'Hide row', template: 'hideCurrentRow' },
  { key: 'unhide-row', label: 'Unhide row', template: 'unhideCurrentRow' },
  { key: 'hide-column', label: 'Hide column', template: 'hideCurrentColumn' },
  { key: 'unhide-column', label: 'Unhide column', template: 'unhideCurrentColumn' },
] as const

export type StructureActionAvailability = Record<StructureActionTemplate, boolean>

export function hasAvailableStructureAction(availability: StructureActionAvailability): boolean {
  return STRUCTURE_ACTIONS.some((action) => availability[action.template])
}
