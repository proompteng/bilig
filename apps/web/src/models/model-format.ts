import type { ModelFormat } from './model-document.js'

const numberFormats: Record<ModelFormat, Intl.NumberFormat> = {
  number: new Intl.NumberFormat('en-US', { maximumFractionDigits: 2 }),
  currency: new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 2 }),
  percent: new Intl.NumberFormat('en-US', { style: 'percent', maximumFractionDigits: 2 }),
}

export function formatModelNumber(value: number, format: ModelFormat): string {
  return numberFormats[format].format(value)
}
