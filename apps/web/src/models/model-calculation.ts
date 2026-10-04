import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/workpaper/browser'
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
  readonly sensitivity: ModelSensitivity | null
}

export interface ModelSensitivityRequest {
  readonly inputId: string
  readonly step: number
}

export interface ModelSensitivity {
  readonly inputId: string
  readonly rows: readonly { readonly offset: number; readonly inputValue: number; readonly values: readonly ModelResult[] }[]
}

export function parseModelSensitivityRequest(value: unknown): ModelSensitivityRequest | null {
  if (value === null) return null
  if (typeof value !== 'object' || !value || !('inputId' in value) || typeof value.inputId !== 'string') {
    throw new Error('Choose an assumption for what-if analysis.')
  }
  if (!('step' in value) || typeof value.step !== 'number' || !Number.isFinite(value.step) || value.step <= 0) {
    throw new Error('Step size must be a positive finite number.')
  }
  return { inputId: value.inputId, step: value.step }
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

export function calculateModel(model: ModelDocument, sensitivityRequest: ModelSensitivityRequest | null): ModelCalculation {
  const request = parseModelSensitivityRequest(sensitivityRequest)
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
    let sensitivity: ModelSensitivity | null = null
    if (request) {
      const inputIndex = model.inputs.findIndex((input) => input.id === request.inputId)
      const input = model.inputs[inputIndex]
      if (!input) throw new Error('The selected assumption no longer exists.')
      const sheet = workbook.getSheetId('Inputs')
      if (sheet === undefined) throw new Error('The model has no Inputs sheet.')
      const rows = [-2, -1, 0, 1, 2].map((offset) => {
        const inputValue = input.value + offset * request.step
        if (!Number.isFinite(inputValue)) throw new Error('Step size produces values outside the supported numeric range.')
        workbook.setCellContents({ sheet, row: inputIndex + 1, col: 1 }, inputValue)
        return { offset, inputValue, values: readResults(workbook, model) }
      })
      sensitivity = { inputId: input.id, rows }
    }
    return { current, scenarios, workpaperJson, restoreVerified, sensitivity }
  } finally {
    workbook.dispose()
  }
}
