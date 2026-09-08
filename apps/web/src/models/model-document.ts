export const MODEL_FORMAT = 'bilig.model.v1'
export const MODEL_LIMITS = { inputs: 64, outputs: 32, scenarios: 12, fileBytes: 1_000_000 } as const

export type ModelFormat = 'number' | 'currency' | 'percent'

export interface ModelInput {
  readonly id: string
  readonly label: string
  readonly value: number
  readonly format: ModelFormat
}

export interface ModelOutput {
  readonly id: string
  readonly label: string
  readonly formula: string
  readonly format: ModelFormat
}

export interface ModelDefinition {
  readonly inputs: readonly ModelInput[]
  readonly outputs: readonly ModelOutput[]
}

export interface ModelScenario extends ModelDefinition {
  readonly id: string
  readonly name: string
  readonly createdAt: string
}

export interface ModelDocument extends ModelDefinition {
  readonly format: typeof MODEL_FORMAT
  readonly id: string
  readonly title: string
  readonly description: string
  readonly createdAt: string
  readonly updatedAt: string
  readonly revision: number
  readonly status: 'active' | 'archived'
  readonly scenarios: readonly ModelScenario[]
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function record(value: unknown): Record<string, unknown> {
  if (!isRecord(value)) throw new Error('Expected a model object.')
  return value
}

function text(value: unknown, label: string, maxLength = 120): string {
  if (typeof value !== 'string' || !value.trim() || value.length > maxLength) throw new Error(`Invalid ${label}.`)
  return value
}

function identifier(value: unknown): string {
  const id = text(value, 'identifier')
  if (!/^[a-zA-Z0-9_-]+$/u.test(id)) throw new Error('Invalid model identifier.')
  return id
}

function date(value: unknown): string {
  const result = text(value, 'date')
  if (!Number.isFinite(Date.parse(result))) throw new Error('Invalid model date.')
  return result
}

function format(value: unknown): ModelFormat {
  if (value !== 'number' && value !== 'currency' && value !== 'percent') throw new Error('Invalid number format.')
  return value
}

function items<T extends { readonly id: string }>(value: unknown, limit: number, parse: (entry: unknown) => T): T[] {
  if (!Array.isArray(value) || value.length > limit) throw new Error(`Expected at most ${limit} entries.`)
  const entries = value.map((entry: unknown) => parse(entry))
  if (new Set(entries.map((entry) => entry.id)).size !== entries.length) throw new Error('Duplicate identifiers in model.')
  return entries
}

export function parseModelDefinition(definition: unknown): ModelDefinition {
  const source = record(definition)
  return {
    inputs: items(source['inputs'], MODEL_LIMITS.inputs, (entry) => {
      const input = record(entry)
      const value = input['value']
      if (typeof value !== 'number' || !Number.isFinite(value)) throw new Error('Input values must be finite numbers.')
      return { id: identifier(input['id']), label: text(input['label'], 'input label'), value, format: format(input['format']) }
    }),
    outputs: items(source['outputs'], MODEL_LIMITS.outputs, (entry) => {
      const output = record(entry)
      const formula = text(output['formula'], 'formula', 2048)
      if (!formula.startsWith('=')) throw new Error('Result formulas must start with =.')
      return { id: identifier(output['id']), label: text(output['label'], 'result label'), formula, format: format(output['format']) }
    }),
  }
}

export function parseModelDocument(value: unknown): ModelDocument {
  const source = record(value)
  if (source['format'] !== MODEL_FORMAT) throw new Error('This file is not a supported Bilig model backup.')
  const revision = source['revision']
  if (typeof revision !== 'number' || !Number.isSafeInteger(revision) || revision < 0) throw new Error('Invalid model revision.')
  const status = source['status']
  if (status !== 'active' && status !== 'archived') throw new Error('Invalid model status.')
  const description = source['description']
  if (typeof description !== 'string' || description.length > 1000) throw new Error('Invalid model description.')
  return {
    ...parseModelDefinition(source),
    format: MODEL_FORMAT,
    id: identifier(source['id']),
    title: text(source['title'], 'model title'),
    description,
    createdAt: date(source['createdAt']),
    updatedAt: date(source['updatedAt']),
    revision,
    status,
    scenarios: items(source['scenarios'], MODEL_LIMITS.scenarios, (entry) => {
      const scenario = record(entry)
      return {
        ...parseModelDefinition(scenario),
        id: identifier(scenario['id']),
        name: text(scenario['name'], 'scenario name'),
        createdAt: date(scenario['createdAt']),
      }
    }),
  }
}

export function importModelBackup(json: string, id = crypto.randomUUID(), now = new Date().toISOString()): ModelDocument {
  if (new TextEncoder().encode(json).byteLength > MODEL_LIMITS.fileBytes) throw new Error('Model backups must be smaller than 1 MB.')
  const parsed: unknown = JSON.parse(json)
  const model = parseModelDocument(parsed)
  return { ...model, id, revision: 0, status: 'active', createdAt: now, updatedAt: now }
}

export function captureModelScenario(model: ModelDocument, name: string, id = crypto.randomUUID()): ModelDocument {
  if (model.scenarios.length >= MODEL_LIMITS.scenarios) throw new Error('This model already has 12 scenarios.')
  return {
    ...model,
    scenarios: [
      ...model.scenarios,
      {
        id,
        name: text(name.trim(), 'scenario name'),
        createdAt: new Date().toISOString(),
        ...structuredClone({ inputs: model.inputs, outputs: model.outputs }),
      },
    ],
  }
}
