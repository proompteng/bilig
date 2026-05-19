export interface AgentDiscoveryMcpCardOptions {
  readonly headlessPackageSpec: string
  readonly headlessPackageVersion: string
  readonly remoteMcpEndpoint: string
  readonly repositoryUrl: string
  readonly siteRoot: string
}

type JsonObject = Record<string, unknown>

export function mcpServerCardManifest(options: AgentDiscoveryMcpCardOptions): string {
  const json = JSON.stringify(
    {
      $schema: 'https://modelcontextprotocol.io/schemas/server-card/v1.0',
      version: '1.0',
      protocolVersion: '2025-11-25',
      serverInfo: {
        name: 'Bilig WorkPaper',
        version: options.headlessPackageVersion,
        description:
          'Formula-backed WorkPaper MCP tools for workbook reads, input edits, recalculation, formula validation, and JSON persistence.',
      },
      repository: {
        type: 'git',
        url: options.repositoryUrl,
      },
      homepage: `${options.siteRoot}/`,
      license: 'MIT',
      authentication: {
        required: false,
        schemes: [],
      },
      transport: npmStdioTransport(options.headlessPackageSpec),
      transports: [
        {
          type: 'streamable-http',
          url: options.remoteMcpEndpoint,
          protocolVersion: '2025-11-25',
          stateless: true,
        },
        npmStdioTransport(options.headlessPackageSpec),
      ],
      capabilities: {
        tools: true,
        resources: true,
        prompts: true,
      },
      tools: fileBackedWorkPaperTools(),
      resources: [
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
      ],
      prompts: [
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
      ],
    },
    null,
    2,
  )
  return `${compactStringOnlyArrays(json)}\n`
}

function npmStdioTransport(headlessPackageSpec: string): JsonObject {
  return {
    type: 'stdio',
    command: 'npm',
    args: ['exec', '--package', headlessPackageSpec, '--', 'bilig-workpaper-mcp', '--demo-workpaper-tools'],
  }
}

function fileBackedWorkPaperTools(): readonly JsonObject[] {
  return [
    {
      name: 'list_sheets',
      title: 'List WorkPaper Sheets',
      description:
        'Discover sheet names and used dimensions before reading or editing a WorkPaper. Returns metadata only; use read_range or read_cell for values.',
      inputSchema: emptySchema(),
      outputSchema: objectSchema({
        required: ['writable', 'sheets'],
        properties: {
          sourcePath: stringOutput('Absolute JSON file path when the server was started with --workpaper.'),
          writable: booleanOutput('Whether set_cell_contents persists edits back to the source JSON file.'),
          sheets: {
            type: 'array',
            description: 'Sheet names, ids, and current used dimensions.',
            items: objectSchema({
              required: ['id', 'name', 'dimensions'],
              properties: {
                id: numberOutput('Stable numeric sheet id inside the WorkPaper engine.'),
                name: stringOutput('Sheet name to pass as sheetName in read and write calls.'),
                dimensions: objectSchema({
                  description: 'Current used rows and columns for the sheet.',
                  properties: {
                    rowCount: numberOutput('Used row count for the sheet.'),
                    columnCount: numberOutput('Used column count for the sheet.'),
                  },
                }),
              },
            }),
          },
        },
      }),
      annotations: toolAnnotations('List WorkPaper Sheets', true),
    },
    {
      name: 'read_range',
      title: 'Read WorkPaper Range',
      description:
        'Read calculated values plus serialized formulas/inputs for an A1 range. Use for audit readback after edits; use read_cell for one address.',
      inputSchema: objectSchema({
        required: ['range'],
        properties: {
          range: stringInput('A1 range such as Summary!A1:B5. If omitted from the range, pass sheetName separately.'),
          sheetName: stringInput('Default sheet name when range omits a sheet name, for example Summary.'),
        },
      }),
      outputSchema: objectSchema({
        required: ['range', 'values', 'serialized'],
        properties: {
          range: stringOutput('Canonical A1 range including the sheet name.'),
          values: arrayOutput('Two-dimensional array of evaluated cell values.'),
          serialized: arrayOutput('Two-dimensional array of raw serialized cell contents, including formulas.'),
        },
      }),
      annotations: toolAnnotations('Read WorkPaper Range', true),
    },
    {
      name: 'read_cell',
      title: 'Read WorkPaper Cell',
      description:
        'Read one cell with calculated value, display text, formula text, and serialized content. Use after set_cell_contents to verify readback.',
      inputSchema: cellAddressInputSchema(
        'Existing sheet name, for example Inputs.',
        'Single A1 cell address such as B3. Ranges are not accepted.',
      ),
      outputSchema: cellReadOutputSchema(),
      annotations: toolAnnotations('Read WorkPaper Cell', true),
    },
    {
      name: 'set_cell_contents',
      title: 'Set WorkPaper Cell Contents',
      description:
        'Write raw content to one cell, recalculate dependents, persist the WorkPaper JSON file when writable, and return before/after/restored readback.',
      inputSchema: objectSchema({
        required: ['sheetName', 'address', 'value'],
        properties: {
          sheetName: stringInput('Existing sheet name, for example Inputs.'),
          address: stringInput('Single A1 cell address such as B3. Ranges are not accepted.'),
          value: {
            type: ['string', 'number', 'boolean', 'null'],
            description: 'Raw cell content. Formula strings must start with =.',
          },
        },
      }),
      outputSchema: objectSchema({
        required: ['editedCell', 'before', 'after', 'restored', 'persistence', 'checks'],
        properties: {
          editedCell: stringOutput('Canonical sheet-qualified address that was edited.'),
          before: cellReadOutputSchema('Cell readback before the edit.'),
          after: cellReadOutputSchema('Cell readback after recalculation.'),
          restored: cellReadOutputSchema('Cell readback after exporting and restoring WorkPaper JSON.'),
          persistence: objectSchema({
            description: 'Persistence result for the WorkPaper JSON document.',
            required: ['persisted', 'serializedBytes'],
            properties: {
              persisted: booleanOutput('True when the server wrote the updated WorkPaper JSON file.'),
              path: stringOutput('Absolute JSON file path written by the server when writable.'),
              serializedBytes: numberOutput('UTF-8 byte length of the serialized WorkPaper document.'),
            },
          }),
          checks: objectSchema({
            description: 'Boolean receipt for persisted state and restored readback.',
            required: ['persisted', 'restoredMatchesAfter', 'previousSerialized', 'newSerialized'],
            properties: {
              persisted: booleanOutput('Echo of whether the edit was persisted to disk.'),
              restoredMatchesAfter: booleanOutput('True when exported and re-imported JSON preserves the edited cell readback.'),
              previousSerialized: rawCellContentOutput('Raw serialized cell content before the edit.'),
              newSerialized: rawCellContentOutput('Raw serialized cell content after the edit.'),
            },
          }),
        },
      }),
      annotations: toolAnnotations('Set WorkPaper Cell Contents', false, true),
    },
    {
      name: 'get_cell_display_value',
      title: 'Get WorkPaper Cell Display Value',
      description:
        'Return the formatted display string for one cell. Use when an agent needs what a user would see, not the raw numeric value.',
      inputSchema: cellAddressInputSchema(
        'Existing sheet name, for example Summary.',
        'Single A1 cell address such as B2. Ranges are not accepted.',
      ),
      outputSchema: objectSchema({
        required: ['address', 'displayValue'],
        properties: {
          address: stringOutput('Canonical sheet-qualified A1 address.'),
          displayValue: stringOutput('Formatted value string suitable for human readback.'),
        },
      }),
      annotations: toolAnnotations('Get WorkPaper Cell Display Value', true),
    },
    {
      name: 'export_workpaper_document',
      title: 'Export WorkPaper Document',
      description:
        'Export the current WorkPaper JSON document for persistence, review, or handoff to another agent. Does not write files by itself.',
      inputSchema: objectSchema({
        properties: {
          includeConfig: {
            type: 'boolean',
            default: true,
            description: 'Include workbook configuration metadata in the exported JSON. Defaults to true.',
          },
        },
      }),
      outputSchema: objectSchema({
        required: ['document', 'serializedBytes'],
        properties: {
          sourcePath: stringOutput('Absolute JSON file path when the server was started with --workpaper.'),
          document: objectSchema({ description: 'Persisted WorkPaper JSON document.' }),
          serializedBytes: numberOutput('UTF-8 byte length of the serialized WorkPaper document.'),
        },
      }),
      annotations: toolAnnotations('Export WorkPaper Document', true),
    },
    {
      name: 'validate_formula',
      title: 'Validate WorkPaper Formula',
      description:
        'Validate formula syntax with the WorkPaper parser before writing it to a cell. This checks syntax only; use set_cell_contents plus readback to evaluate.',
      inputSchema: objectSchema({
        required: ['formula'],
        properties: {
          formula: stringInput('Formula string including the leading =, for example =SUM(Inputs!B2:B4).'),
        },
      }),
      outputSchema: objectSchema({
        required: ['formula', 'valid'],
        properties: {
          formula: stringOutput('Formula string that was validated.'),
          valid: booleanOutput('True when the WorkPaper formula parser accepts the formula syntax.'),
        },
      }),
      annotations: toolAnnotations('Validate WorkPaper Formula', true),
    },
  ]
}

function emptySchema(): JsonObject {
  return objectSchema({})
}

function objectSchema(input: {
  readonly description?: string
  readonly required?: readonly string[]
  readonly properties?: JsonObject
}): JsonObject {
  return {
    type: 'object',
    ...(input.description === undefined ? {} : { description: input.description }),
    ...(input.required === undefined ? {} : { required: input.required }),
    ...(input.properties === undefined ? {} : { properties: input.properties }),
    additionalProperties: false,
  }
}

function cellAddressInputSchema(sheetDescription: string, addressDescription: string): JsonObject {
  return objectSchema({
    required: ['sheetName', 'address'],
    properties: {
      sheetName: stringInput(sheetDescription),
      address: stringInput(addressDescription),
    },
  })
}

function cellReadOutputSchema(description?: string): JsonObject {
  return objectSchema({
    description,
    required: ['address', 'value', 'serialized', 'formula', 'displayValue'],
    properties: {
      address: stringOutput('Canonical sheet-qualified A1 address.'),
      value: {
        description: 'Calculated cell value.',
      },
      serialized: rawCellContentOutput('Raw serialized cell content, preserving formulas.'),
      formula: {
        type: ['string', 'null'],
        description: 'Formula text without losing the original calculated value context, or null for literal cells.',
      },
      displayValue: stringOutput('Formatted value as a user would see it.'),
    },
  })
}

function rawCellContentOutput(description: string): JsonObject {
  return {
    type: ['string', 'number', 'boolean', 'null'],
    description,
  }
}

function toolAnnotations(title: string, readOnlyHint: boolean, destructiveHint = false): JsonObject {
  return {
    title,
    readOnlyHint,
    destructiveHint,
    idempotentHint: true,
    openWorldHint: false,
  }
}

function stringInput(description: string): JsonObject {
  return { type: 'string', description }
}

function stringOutput(description: string): JsonObject {
  return { type: 'string', description }
}

function numberOutput(description: string): JsonObject {
  return { type: 'number', description }
}

function booleanOutput(description: string): JsonObject {
  return { type: 'boolean', description }
}

function arrayOutput(description: string): JsonObject {
  return { type: 'array', description }
}

function compactStringOnlyArrays(json: string): string {
  return json.replace(/\[\n((?:\s+"(?:[^"\\]|\\.)+"(?:,\n)?)+)\s+\]/g, (match) => {
    const values = [...match.matchAll(/"((?:[^"\\]|\\.)*)"/g)].map(([, value]) => `"${value}"`)
    return `[${values.join(', ')}]`
  })
}
