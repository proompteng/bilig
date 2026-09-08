import { MODEL_FORMAT, type ModelDefinition, type ModelDocument } from './model-document.js'

interface ModelTemplate extends ModelDefinition {
  readonly id: string
  readonly title: string
  readonly description: string
  readonly question: string
}

export const MODEL_TEMPLATES: readonly ModelTemplate[] = [
  {
    id: 'contribution',
    title: 'Contribution model',
    description: 'Revenue, variable costs, and the volume needed to cover fixed costs.',
    question: 'What makes this profitable?',
    inputs: [
      { id: 'units', label: 'Units sold', value: 100, format: 'number' },
      { id: 'price', label: 'Price per unit', value: 100, format: 'currency' },
      { id: 'cost', label: 'Cost per unit', value: 40, format: 'currency' },
      { id: 'fixed', label: 'Fixed costs', value: 3000, format: 'currency' },
    ],
    outputs: [
      { id: 'revenue', label: 'Revenue', formula: '=Inputs!B2*Inputs!B3', format: 'currency' },
      { id: 'profit', label: 'Operating profit', formula: '=Inputs!B2*(Inputs!B3-Inputs!B4)-Inputs!B5', format: 'currency' },
      { id: 'margin', label: 'Operating margin', formula: '=Results!B3/Results!B2', format: 'percent' },
      { id: 'breakeven', label: 'Break-even units', formula: '=ROUNDUP(Inputs!B5/(Inputs!B3-Inputs!B4),0)', format: 'number' },
    ],
  },
  {
    id: 'project',
    title: 'Project budget',
    description: 'Estimate delivery costs, contingency, and the budget remaining.',
    question: 'Can we deliver within budget?',
    inputs: [
      { id: 'budget', label: 'Project budget', value: 50000, format: 'currency' },
      { id: 'hours', label: 'Estimated hours', value: 320, format: 'number' },
      { id: 'rate', label: 'Hourly cost', value: 100, format: 'currency' },
      { id: 'expenses', label: 'Other expenses', value: 5000, format: 'currency' },
      { id: 'contingency', label: 'Contingency', value: 0.1, format: 'percent' },
    ],
    outputs: [
      { id: 'labor', label: 'Labor cost', formula: '=Inputs!B3*Inputs!B4', format: 'currency' },
      { id: 'total', label: 'Cost with contingency', formula: '=(Results!B2+Inputs!B5)*(1+Inputs!B6)', format: 'currency' },
      { id: 'remaining', label: 'Budget remaining', formula: '=Inputs!B2-Results!B3', format: 'currency' },
      { id: 'used', label: 'Budget used', formula: '=Results!B3/Inputs!B2', format: 'percent' },
    ],
  },
  {
    id: 'capacity',
    title: 'Team capacity',
    description: 'Turn available hours and expected demand into a staffing estimate.',
    question: 'Do we have enough capacity?',
    inputs: [
      { id: 'people', label: 'Team members', value: 6, format: 'number' },
      { id: 'hours', label: 'Hours per person / week', value: 40, format: 'number' },
      { id: 'focus', label: 'Delivery time', value: 0.7, format: 'percent' },
      { id: 'demand', label: 'Required hours / week', value: 200, format: 'number' },
    ],
    outputs: [
      { id: 'available', label: 'Delivery hours / week', formula: '=Inputs!B2*Inputs!B3*Inputs!B4', format: 'number' },
      { id: 'gap', label: 'Capacity remaining', formula: '=Results!B2-Inputs!B5', format: 'number' },
      { id: 'needed', label: 'People needed', formula: '=ROUNDUP(Inputs!B5/(Inputs!B3*Inputs!B4),0)', format: 'number' },
    ],
  },
  {
    id: 'custom',
    title: 'Untitled model',
    description: 'Name your assumptions and add formulas for the results you need.',
    question: 'Build your own model',
    inputs: [{ id: 'input-1', label: 'First input', value: 10, format: 'number' }],
    outputs: [{ id: 'output-1', label: 'First result', formula: '=Inputs!B2*2', format: 'number' }],
  },
]

export function createModel(templateId: string, id = crypto.randomUUID(), now = new Date().toISOString()): ModelDocument {
  const template = MODEL_TEMPLATES.find((entry) => entry.id === templateId)
  if (!template) throw new Error('Unknown model template.')
  return {
    format: MODEL_FORMAT,
    id,
    title: template.title,
    description: template.description,
    createdAt: now,
    updatedAt: now,
    revision: 0,
    status: 'active',
    scenarios: [],
    ...structuredClone({ inputs: template.inputs, outputs: template.outputs }),
  }
}
