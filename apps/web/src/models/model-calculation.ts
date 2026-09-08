import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless/browser'
import { ValueTag } from '@bilig/protocol'
import type { ModelDefinition, ModelDocument } from './model-document.js'
import { formatModelNumber } from './model-format.js'

export type ModelResult = {
  readonly id: string
  readonly text: string
} & ({ readonly kind: 'number'; readonly value: number } | { readonly kind: 'text' | 'error' })

export interface ModelCalculation {
  readonly current: readonly ModelResult[]
  readonly scenarios: readonly { readonly id: string; readonly values: readonly ModelResult[] }[]
  readonly workpaperJson: string
  readonly restoreVerified: boolean
}

function buildWorkbook(definition: ModelDefinition): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      Inputs: [['Assumption', 'Value'], ...definition.inputs.map((input) => [input.label, input.value])],
      Results: [['Result', 'Formula'], ...definition.outputs.map((output) => [output.label, output.formula])],
    },
    { evaluationTimeoutMs: 200, maxRows: 1000, maxColumns: 64 },
  )
}

function readResults(workbook: WorkPaper, definition: ModelDefinition): ModelResult[] {
  const sheet = workbook.getSheetId('Results')
  if (sheet === undefined) throw new Error('The model has no Results sheet.')
  return definition.outputs.map((output, index) => {
    const address = { sheet, col: 1, row: index + 1 }
    const value = workbook.getCellValue(address)
    if (value.tag === ValueTag.Number && Number.isFinite(value.value)) {
      return { id: output.id, kind: 'number', value: value.value, text: formatModelNumber(value.value, output.format) }
    }
    return { id: output.id, kind: value.tag === ValueTag.Error ? 'error' : 'text', text: workbook.getCellDisplayValue(address) }
  })
}

export function calculateModel(model: ModelDocument): ModelCalculation {
  const workbook = buildWorkbook(model)
  try {
    const current = readResults(workbook, model)
    const workpaperJson = serializeWorkPaperDocument(exportWorkPaperDocument(workbook))
    const restored = createWorkPaperFromDocument(parseWorkPaperDocument(workpaperJson))
    let restoreVerified: boolean
    try {
      restoreVerified = JSON.stringify(readResults(restored, model)) === JSON.stringify(current)
    } finally {
      restored.dispose()
    }
    const scenarios = model.scenarios.map((scenario) => {
      const scenarioWorkbook = buildWorkbook(scenario)
      try {
        return { id: scenario.id, values: readResults(scenarioWorkbook, scenario) }
      } finally {
        scenarioWorkbook.dispose()
      }
    })
    return { current, scenarios, workpaperJson, restoreVerified }
  } finally {
    workbook.dispose()
  }
}
