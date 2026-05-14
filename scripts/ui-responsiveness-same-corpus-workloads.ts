export const requiredUiResponsivenessSameCorpusWorkloads = [
  'open-workbook',
  'select-cell',
  'edit-visible-cell',
  'scroll-vertical',
  'scroll-horizontal',
  'jump-deep-row',
  'formula-edit',
  'fill-format-change',
  'wide-sheet-navigation',
] as const

export type UiResponsivenessSameCorpusWorkload = (typeof requiredUiResponsivenessSameCorpusWorkloads)[number]

export function isUiResponsivenessSameCorpusWorkload(value: string): value is UiResponsivenessSameCorpusWorkload {
  return (requiredUiResponsivenessSameCorpusWorkloads as readonly string[]).includes(value)
}
