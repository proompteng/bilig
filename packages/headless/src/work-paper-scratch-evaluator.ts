import type { CellValue } from "@bilig/protocol";
import type {
  RawCellContent,
  SerializedWorkPaperNamedExpression,
  WorkPaperCellAddress,
  WorkPaperCellRange,
  WorkPaperConfig,
  WorkPaperSheet,
  WorkPaperSheets,
} from "./work-paper-types.js";

export interface WorkPaperScratchWorkbook {
  readonly engine: {
    createSheet(sheetName: string): void;
    getSpillRanges(): Array<{
      sheetName: string;
      address: string;
      rows: number;
      cols: number;
    }>;
  };
  registerNamedExpression(expression: SerializedWorkPaperNamedExpression): void;
  requireSheetId(sheetName: string): number;
  replaceSheetContent(sheetId: number, sheet: WorkPaperSheet): void;
  clearHistoryStacks(): void;
  applyRawContent(address: WorkPaperCellAddress, content: RawCellContent): void;
  getRangeValues(range: WorkPaperCellRange): CellValue[][];
  getCellValue(address: WorkPaperCellAddress): CellValue;
  dispose(): void;
}

export function calculateWorkPaperFormulaInScratchWorkbook(args: {
  createWorkbook: (config: WorkPaperConfig) => WorkPaperScratchWorkbook;
  config: WorkPaperConfig;
  serializedSheets: WorkPaperSheets;
  namedExpressions: readonly SerializedWorkPaperNamedExpression[];
  formula: string;
  scope?: number;
}): CellValue | CellValue[][] {
  const temporaryWorkbook = args.createWorkbook(args.config);
  Object.keys(args.serializedSheets).forEach((sheetName) => {
    temporaryWorkbook.engine.createSheet(sheetName);
  });
  args.namedExpressions.forEach((expression) => {
    temporaryWorkbook.registerNamedExpression(expression);
  });
  Object.entries(args.serializedSheets).forEach(([sheetName, sheet]) => {
    const sheetId = temporaryWorkbook.requireSheetId(sheetName);
    temporaryWorkbook.replaceSheetContent(sheetId, sheet);
  });
  temporaryWorkbook.clearHistoryStacks();
  const scratchSheetName =
    args.scope !== undefined ? `__WORKPAPER_CALC_${args.scope}__` : "__WORKPAPER_CALC__";
  temporaryWorkbook.engine.createSheet(scratchSheetName);
  const scratchSheetId = temporaryWorkbook.requireSheetId(scratchSheetName);
  temporaryWorkbook.applyRawContent(
    { sheet: scratchSheetId, row: 0, col: 0 },
    args.formula.trim().startsWith("=") ? args.formula : `=${args.formula}`,
  );
  const spill = temporaryWorkbook.engine
    .getSpillRanges()
    .find((candidate) => candidate.sheetName === scratchSheetName && candidate.address === "A1");
  const value = spill
    ? temporaryWorkbook.getRangeValues({
        start: { sheet: scratchSheetId, row: 0, col: 0 },
        end: { sheet: scratchSheetId, row: spill.rows - 1, col: spill.cols - 1 },
      })
    : temporaryWorkbook.getCellValue({ sheet: scratchSheetId, row: 0, col: 0 });
  temporaryWorkbook.dispose();
  return value;
}
