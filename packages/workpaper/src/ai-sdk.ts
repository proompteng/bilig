import { createWorkPaperFromDocument, exportWorkPaperDocument, parseWorkPaperDocument, serializeWorkPaperDocument } from 'bilig-workpaper'
import type { WorkPaper } from 'bilig-workpaper'
import { tool } from 'ai'
import { z } from 'zod'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSheets>
type CellAddress = NonNullable<ReturnType<WorkPaperInstance['simpleCellAddressFromString']>>

export type AiSdkWorkPaperCellValue = string | number | boolean | null

export interface AiSdkWorkPaperToolsOptions {
  readonly workpaper: WorkPaperInstance
  readonly defaultReadRange?: string
  readonly proofRange?: string
  readonly writableSheets?: readonly string[]
  readonly includeSerializedDocument?: boolean
}

export interface AiSdkWorkPaperSetCellArgs {
  readonly sheetName: string
  readonly address: string
  readonly value: AiSdkWorkPaperCellValue
}

export interface AiSdkWorkPaperReadRangeArgs {
  readonly range?: string
}

export interface AiSdkWorkPaperReadResult {
  readonly range: string
  readonly values: unknown[][]
  readonly serialized: unknown[][]
}

export interface AiSdkWorkPaperWriteResult {
  readonly editedCell: string
  readonly before: AiSdkWorkPaperReadResult
  readonly after: AiSdkWorkPaperReadResult
  readonly restored: AiSdkWorkPaperReadResult
  readonly checks: {
    readonly previousValue: unknown
    readonly newValue: unknown
    readonly formulasPersisted: boolean
    readonly restoredMatchesAfter: boolean
    readonly proofRangeChanged: boolean
    readonly serializedBytes: number
  }
  readonly serializedDocument?: string
}

const setInputCellSchema = z.object({
  sheetName: z.string().min(1).describe('Target sheet name, for example Inputs.'),
  address: z
    .string()
    .regex(/^[A-Z]+[1-9][0-9]*$/)
    .describe('A1 cell address inside the target sheet, for example B3.'),
  value: z
    .union([z.string(), z.number(), z.boolean(), z.null()])
    .describe('Literal cell value. Formula strings are accepted only when your WorkPaper contract allows them.'),
})

export function createAiSdkWorkPaperTools(options: AiSdkWorkPaperToolsOptions) {
  const handlers = createWorkPaperToolHandlers(options)
  const defaultReadRange = options.defaultReadRange ?? 'Summary!A1:B5'

  return {
    readWorkPaperSummary: tool({
      description: 'Read computed WorkPaper values and serialized cells for a small proof range.',
      inputSchema: createReadRangeInputSchema(defaultReadRange),
      execute: async ({ range = defaultReadRange }: AiSdkWorkPaperReadRangeArgs = {}) => {
        return handlers.readWorkPaperSummary(range)
      },
    }),

    setWorkPaperInputCell: tool({
      description: 'Set one validated WorkPaper input cell and return before/after/restore formula readback.',
      inputSchema: setInputCellSchema,
      execute: async (args: AiSdkWorkPaperSetCellArgs) => handlers.setWorkPaperInputCell(args),
    }),
  }
}

function createReadRangeInputSchema(defaultReadRange: string) {
  return z.object({
    range: z.string().default(defaultReadRange).describe(`A small A1 range including the sheet name, for example ${defaultReadRange}.`),
  })
}

export function createWorkPaperToolHandlers(options: AiSdkWorkPaperToolsOptions) {
  const defaultReadRange = options.defaultReadRange ?? 'Summary!A1:B5'
  const proofRange = options.proofRange ?? defaultReadRange

  return {
    readWorkPaperSummary(range = defaultReadRange): AiSdkWorkPaperReadResult {
      return readWorkPaperRange(options.workpaper, range)
    },

    setWorkPaperInputCell(args: AiSdkWorkPaperSetCellArgs): AiSdkWorkPaperWriteResult {
      const parsedArgs = setInputCellSchema.parse(args)
      assertWritableSheet(parsedArgs.sheetName, options.writableSheets)

      const address = requireCellAddress(options.workpaper, parsedArgs.sheetName, parsedArgs.address)
      const before = readWorkPaperRange(options.workpaper, proofRange)
      const previousValue = options.workpaper.getCellSerialized(address)

      options.workpaper.setCellContents(address, parsedArgs.value)

      const after = readWorkPaperRange(options.workpaper, proofRange)
      const saved = serializeWorkPaperDocument(
        exportWorkPaperDocument(options.workpaper, {
          includeConfig: true,
        }),
      )
      const restored = createWorkPaperFromDocument(parseWorkPaperDocument(saved))
      try {
        const restoredReadback = readWorkPaperRange(restored, proofRange)

        return {
          editedCell: options.workpaper.simpleCellAddressToString(address, {
            includeSheetName: true,
          }),
          before,
          after,
          restored: restoredReadback,
          checks: {
            previousValue,
            newValue: options.workpaper.getCellSerialized(address),
            formulasPersisted: sameJson(after.serialized, restoredReadback.serialized),
            restoredMatchesAfter: sameJson(after.values, restoredReadback.values),
            proofRangeChanged: !sameJson(before.values, after.values),
            serializedBytes: Buffer.byteLength(saved, 'utf8'),
          },
          ...(options.includeSerializedDocument ? { serializedDocument: saved } : {}),
        }
      } finally {
        restored.dispose()
      }
    },
  }
}

function readWorkPaperRange(workpaper: WorkPaperInstance, range: string): AiSdkWorkPaperReadResult {
  const parsedRange = workpaper.simpleCellRangeFromString(range, resolveDefaultSheet(workpaper))
  if (parsedRange === undefined) {
    throw new Error(`Invalid readable WorkPaper range: ${range}`)
  }

  return {
    range,
    values: workpaper.getRangeValues(parsedRange),
    serialized: workpaper.getRangeSerialized(parsedRange),
  }
}

function requireCellAddress(workpaper: WorkPaperInstance, sheetName: string, address: string): CellAddress {
  const sheetId = requireSheet(workpaper, sheetName)
  const parsedAddress = workpaper.simpleCellAddressFromString(address, sheetId)

  if (parsedAddress === undefined || parsedAddress.sheet !== sheetId) {
    throw new Error(`Invalid WorkPaper address: ${sheetName}!${address}`)
  }

  return parsedAddress
}

function requireSheet(workpaper: WorkPaperInstance, sheetName: string): number {
  const sheetId = workpaper.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error(`Expected sheet "${sheetName}" to exist`)
  }
  return sheetId
}

function resolveDefaultSheet(workpaper: WorkPaperInstance): number {
  const firstSheetName = workpaper.getSheetNames()[0]
  if (firstSheetName === undefined) {
    throw new Error('Expected WorkPaper to contain at least one sheet')
  }
  return requireSheet(workpaper, firstSheetName)
}

function assertWritableSheet(sheetName: string, writableSheets: readonly string[] | undefined): void {
  if (writableSheets !== undefined && !writableSheets.includes(sheetName)) {
    throw new Error(`Sheet "${sheetName}" is not writable for this AI SDK tool`)
  }
}

function sameJson(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}
