import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/workpaper'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSheets>

const workbook = WorkPaper.buildFromSheets(
  {
    Inputs: [
      ['Metric', 'Value'],
      ['Committed MRR', 36000],
      ['Expansion MRR', 18000],
      ['Pipeline MRR', 15000],
    ],
    Summary: [
      ['Metric', 'Value'],
      ['Total MRR', '=SUM(Inputs!B2:B4)'],
      ['Committed Share', '=Inputs!B2/B2'],
    ],
  },
  {
    maxRows: 1000,
    maxColumns: 64,
    useColumnIndex: true,
  },
)

const serialized = serializeWorkPaperDocument(
  exportWorkPaperDocument(workbook, {
    includeConfig: true,
  }),
)
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serialized))
const summarySheet = requireSheet(restored, 'Summary')

const output = {
  verified: true,
  restoredSheets: restored.getSheetNames(),
  lookup: {
    query: 'Summary',
    sheetId: summarySheet,
    sheetName: restored.getSheetName(summarySheet),
    dimensions: restored.getSheetDimensions(summarySheet),
  },
}

assertOutput(output)
console.log(JSON.stringify(output, null, 2))

function requireSheet(workpaper: WorkPaperInstance, sheetName: string): number {
  const sheetId = workpaper.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error(`Expected sheet "${sheetName}" to exist`)
  }
  return sheetId
}

function assertOutput(actual: typeof output): void {
  const expected = {
    verified: true,
    restoredSheets: ['Inputs', 'Summary'],
    lookup: {
      query: 'Summary',
      sheetId: 2,
      sheetName: 'Summary',
      dimensions: {
        width: 2,
        height: 3,
      },
    },
  }

  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(`Unexpected sheet inspection result: ${JSON.stringify(actual)}`)
  }
}
