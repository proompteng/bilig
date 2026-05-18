import { exportWorkPaperDocument, serializeWorkPaperDocument } from './persistence.js'
import type { WorkPaper } from './work-paper.js'
import type { WorkPaperSheetDimensions } from './work-paper-types.js'

type JsonObject = Record<string, unknown>

type WorkPaperMcpResourceUri =
  | 'bilig://workpaper/manifest'
  | 'bilig://workpaper/agent-handoff'
  | 'bilig://workpaper/sheets'
  | 'bilig://workpaper/current-document'

interface WorkPaperMcpResourceDefinition {
  uri: WorkPaperMcpResourceUri
  name: string
  title: string
  description: string
  mimeType: string
}

type WorkPaperMcpPromptName = 'edit_and_verify_workpaper' | 'debug_workpaper_formula'

interface WorkPaperMcpPromptArgument {
  name: string
  title: string
  description: string
  required?: boolean
}

interface WorkPaperMcpPromptDefinition {
  name: WorkPaperMcpPromptName
  title: string
  description: string
  arguments: WorkPaperMcpPromptArgument[]
}

interface WorkPaperMcpToolDiscoveryDefinition {
  name: string
  title: string
  description: string
  annotations: {
    readOnlyHint: boolean
    destructiveHint: boolean
  }
}

interface WorkPaperMcpSheetSummary {
  id: number
  name: string
  dimensions: WorkPaperSheetDimensions
}

function createFileBackedResourceDefinitions(): WorkPaperMcpResourceDefinition[] {
  return [
    {
      uri: 'bilig://workpaper/manifest',
      name: 'workpaper_manifest',
      title: 'WorkPaper MCP Manifest',
      description: 'Live manifest of the current WorkPaper file, available tools, prompts, resources, and verification contract.',
      mimeType: 'application/json',
    },
    {
      uri: 'bilig://workpaper/agent-handoff',
      name: 'workpaper_agent_handoff',
      title: 'WorkPaper Agent Handoff',
      description: 'Compact instructions for agents that need to edit workbook formulas without spreadsheet UI automation.',
      mimeType: 'text/markdown',
    },
    {
      uri: 'bilig://workpaper/sheets',
      name: 'workpaper_sheets',
      title: 'WorkPaper Sheets',
      description: 'Current sheet names and used dimensions for the loaded WorkPaper document.',
      mimeType: 'application/json',
    },
    {
      uri: 'bilig://workpaper/current-document',
      name: 'workpaper_current_document',
      title: 'Current WorkPaper Document',
      description: 'Current persisted WorkPaper JSON document as exported from the in-memory engine.',
      mimeType: 'application/json',
    },
  ]
}

function createFileBackedPromptDefinitions(): WorkPaperMcpPromptDefinition[] {
  return [
    {
      name: 'edit_and_verify_workpaper',
      title: 'Edit And Verify WorkPaper',
      description:
        'Guide an agent through a safe WorkPaper edit: read before, validate target, write one cell, read computed output, export JSON, and report proof.',
      arguments: [
        {
          name: 'task',
          title: 'Task',
          description: 'Human-readable workbook edit request.',
        },
        {
          name: 'target_cell',
          title: 'Target Cell',
          description: 'Optional sheet-qualified A1 target such as Inputs!B3.',
        },
        {
          name: 'output_range',
          title: 'Output Range',
          description: 'Optional dependent output range to read after recalculation, such as Summary!A1:B5.',
        },
      ],
    },
    {
      name: 'debug_workpaper_formula',
      title: 'Debug WorkPaper Formula',
      description: 'Guide an agent through formula validation and readback when a WorkPaper formula or dependent output looks wrong.',
      arguments: [
        {
          name: 'formula',
          title: 'Formula',
          description: 'Optional formula text, including the leading =.',
        },
        {
          name: 'cell',
          title: 'Cell',
          description: 'Optional sheet-qualified A1 cell that contains or should contain the formula.',
        },
        {
          name: 'symptom',
          title: 'Symptom',
          description: 'Optional description of the wrong value, parse failure, or behavior being debugged.',
        },
      ],
    },
  ]
}

function readFileBackedResource(input: {
  workbook: WorkPaper
  writable: boolean
  sourcePath: string | undefined
  capabilities: unknown
  toolDefinitions: WorkPaperMcpToolDiscoveryDefinition[]
  resourceDefinitions: WorkPaperMcpResourceDefinition[]
  promptDefinitions: WorkPaperMcpPromptDefinition[]
  params: JsonObject | undefined
}): JsonObject {
  const params = requireRecord(input.params ?? {}, 'resources/read params')
  const uri = requireString(params['uri'], 'uri')

  if (!isWorkPaperMcpResourceUri(uri)) {
    throw new Error(`Unknown WorkPaper MCP resource: ${uri}`)
  }

  const mimeType = resourceMimeType(input.resourceDefinitions, uri)
  return {
    contents: [
      {
        uri,
        mimeType,
        text: resourceText(input, uri),
      },
    ],
  }
}

function resourceText(
  input: {
    workbook: WorkPaper
    writable: boolean
    sourcePath: string | undefined
    capabilities: unknown
    toolDefinitions: WorkPaperMcpToolDiscoveryDefinition[]
    resourceDefinitions: WorkPaperMcpResourceDefinition[]
    promptDefinitions: WorkPaperMcpPromptDefinition[]
  },
  uri: WorkPaperMcpResourceUri,
): string {
  if (uri === 'bilig://workpaper/manifest') {
    return JSON.stringify(
      {
        server: 'bilig-workpaper-mcp',
        sourcePath: input.sourcePath,
        writable: input.writable,
        capabilities: input.capabilities,
        sheets: sheetSummaries(input.workbook),
        tools: input.toolDefinitions.map((tool) => ({
          name: tool.name,
          title: tool.title,
          description: tool.description,
          readOnly: tool.annotations.readOnlyHint,
          destructive: tool.annotations.destructiveHint,
        })),
        resources: input.resourceDefinitions,
        prompts: input.promptDefinitions,
        verificationContract: verificationContract(),
      },
      null,
      2,
    )
  }

  if (uri === 'bilig://workpaper/agent-handoff') {
    return workPaperAgentHandoff(input)
  }

  if (uri === 'bilig://workpaper/sheets') {
    return JSON.stringify(
      {
        sourcePath: input.sourcePath,
        writable: input.writable,
        sheets: sheetSummaries(input.workbook),
      },
      null,
      2,
    )
  }

  const document = exportWorkPaperDocument(input.workbook, { includeConfig: true })
  const serialized = serializeWorkPaperDocument(document)
  return JSON.stringify(
    {
      sourcePath: input.sourcePath,
      writable: input.writable,
      serializedBytes: Buffer.byteLength(serialized, 'utf8'),
      document,
    },
    null,
    2,
  )
}

function getFileBackedPrompt(input: {
  workbook: WorkPaper
  writable: boolean
  sourcePath: string | undefined
  params: JsonObject | undefined
}): JsonObject {
  const params = requireRecord(input.params ?? {}, 'prompts/get params')
  const name = requireString(params['name'], 'name')
  const args = requireRecord(params['arguments'] ?? {}, `${name} arguments`)

  if (name === 'edit_and_verify_workpaper') {
    return promptResult('Edit and verify a WorkPaper cell change.', editAndVerifyPrompt(input, args))
  }

  if (name === 'debug_workpaper_formula') {
    return promptResult('Debug a WorkPaper formula with validation and readback.', debugFormulaPrompt(input, args))
  }

  throw new Error(`Unknown WorkPaper MCP prompt: ${name}`)
}

function promptResult(description: string, text: string): JsonObject {
  return {
    description,
    messages: [
      {
        role: 'user',
        content: {
          type: 'text',
          text,
        },
      },
    ],
  }
}

function editAndVerifyPrompt(
  input: {
    workbook: WorkPaper
    writable: boolean
    sourcePath: string | undefined
  },
  args: JsonObject,
): string {
  const task = optionalString(args['task']) ?? 'Edit the WorkPaper cell requested by the user and verify dependent formula readback.'
  const targetCell = optionalString(args['target_cell']) ?? 'unknown'
  const outputRange = optionalString(args['output_range']) ?? 'the smallest dependent output range that proves the edit'

  return [
    'Use the Bilig WorkPaper MCP server instead of spreadsheet UI automation.',
    '',
    `Task: ${task}`,
    `Source path: ${input.sourcePath ?? 'in-memory WorkPaper'}`,
    `Writable: ${input.writable ? 'yes, set_cell_contents persists to JSON' : 'no, writes recalculate in memory only'}`,
    `Target cell: ${targetCell}`,
    `Dependent output range: ${outputRange}`,
    `Available sheets: ${formatSheetList(input.workbook)}`,
    '',
    'Workflow:',
    '1. Call resources/read for bilig://workpaper/agent-handoff if you need the compact contract.',
    '2. Call list_sheets, then read_range or read_cell for the relevant inputs and formulas before editing.',
    '3. If writing a formula, call validate_formula before set_cell_contents.',
    '4. Call set_cell_contents once for the smallest safe target.',
    '5. Call read_cell or read_range on the dependent output after recalculation.',
    '6. Call export_workpaper_document or resources/read for bilig://workpaper/current-document.',
    '7. Return editedCell, before, after, restored or exported proof, serializedBytes, verified, and limitations.',
    '',
    'Do not claim success from set_cell_contents alone. The proof is computed readback plus exported or restored WorkPaper JSON.',
  ].join('\n')
}

function debugFormulaPrompt(
  input: {
    workbook: WorkPaper
    writable: boolean
    sourcePath: string | undefined
  },
  args: JsonObject,
): string {
  const formula = optionalString(args['formula']) ?? 'unknown formula'
  const cell = optionalString(args['cell']) ?? 'unknown cell'
  const symptom = optionalString(args['symptom']) ?? 'formula output is wrong or unverified'

  return [
    'Debug the WorkPaper formula through Bilig MCP tools.',
    '',
    `Formula: ${formula}`,
    `Cell: ${cell}`,
    `Symptom: ${symptom}`,
    `Source path: ${input.sourcePath ?? 'in-memory WorkPaper'}`,
    `Writable: ${input.writable ? 'yes' : 'no'}`,
    `Available sheets: ${formatSheetList(input.workbook)}`,
    '',
    'Workflow:',
    '1. Call validate_formula on the proposed formula text.',
    '2. Call read_cell for the formula cell and read_range for nearby precedent/dependent cells.',
    '3. If a fix is needed, validate the replacement formula before writing it.',
    '4. Call set_cell_contents, then read the dependent output cell or range after recalculation.',
    '5. Export the WorkPaper document and report parse validity, before/after values, formula readback, persistence, and limitations.',
    '',
    'If validation fails or the dependent output cannot be identified, report that blocker instead of presenting an edit as complete.',
  ].join('\n')
}

function workPaperAgentHandoff(input: {
  workbook: WorkPaper
  writable: boolean
  sourcePath: string | undefined
  toolDefinitions: WorkPaperMcpToolDiscoveryDefinition[]
  promptDefinitions: WorkPaperMcpPromptDefinition[]
}): string {
  return [
    '# Bilig WorkPaper Agent Handoff',
    '',
    `Source path: ${input.sourcePath ?? 'in-memory WorkPaper'}`,
    `Writable: ${input.writable ? 'yes' : 'no'}`,
    `Sheets: ${formatSheetList(input.workbook)}`,
    '',
    'Use these MCP tools instead of spreadsheet UI automation:',
    ...input.toolDefinitions.map((tool) => `- ${tool.name}: ${tool.description}`),
    '',
    'Reusable prompts:',
    ...input.promptDefinitions.map((prompt) => `- ${prompt.name}: ${prompt.description}`),
    '',
    'Verification contract:',
    ...verificationContract().map((step) => `- ${step}`),
    '',
    'Do not report success unless computed readback and exported or restored WorkPaper JSON support the claim.',
  ].join('\n')
}

function formatSheetList(workbook: WorkPaper): string {
  return sheetSummaries(workbook)
    .map((sheet) => `${sheet.name}(${sheet.dimensions.height}x${sheet.dimensions.width})`)
    .join(', ')
}

function optionalString(value: unknown): string | undefined {
  return typeof value === 'string' && value.trim().length > 0 ? value : undefined
}

function isWorkPaperMcpResourceUri(value: string): value is WorkPaperMcpResourceUri {
  return (
    value === 'bilig://workpaper/manifest' ||
    value === 'bilig://workpaper/agent-handoff' ||
    value === 'bilig://workpaper/sheets' ||
    value === 'bilig://workpaper/current-document'
  )
}

function resourceMimeType(definitions: WorkPaperMcpResourceDefinition[], uri: WorkPaperMcpResourceUri): string {
  const definition = definitions.find((candidate) => candidate.uri === uri)
  if (definition === undefined) {
    throw new Error(`Unknown WorkPaper MCP resource: ${uri}`)
  }
  return definition.mimeType
}

function sheetSummaries(workbook: WorkPaper): WorkPaperMcpSheetSummary[] {
  return workbook.getSheetNames().map((name) => {
    const sheetId = workbook.getSheetId(name)
    if (sheetId === undefined) {
      throw new Error(`Expected sheet "${name}" to exist`)
    }

    return {
      id: sheetId,
      name,
      dimensions: workbook.getSheetDimensions(sheetId),
    }
  })
}

function verificationContract(): string[] {
  return [
    'discover sheets and read the relevant range before editing',
    'validate the target sheet and A1 cell address',
    'validate formulas before writing formula text',
    'write the smallest safe input or formula change',
    'read the dependent computed output after recalculation',
    'export or serialize the WorkPaper document',
    'report editedCell, before, after, persisted document bytes, verified, and limitations',
  ]
}

function requireRecord(value: unknown, label: string): JsonObject {
  if (!isRecord(value)) {
    throw new Error(`Expected ${label} to be an object`)
  }
  return value
}

function requireString(value: unknown, label: string): string {
  if (typeof value !== 'string') {
    throw new Error(`Expected ${label} to be a string`)
  }
  return value
}

function isRecord(value: unknown): value is JsonObject {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

export {
  createFileBackedPromptDefinitions,
  createFileBackedResourceDefinitions,
  getFileBackedPrompt,
  readFileBackedResource,
  type WorkPaperMcpPromptDefinition,
  type WorkPaperMcpResourceDefinition,
}
