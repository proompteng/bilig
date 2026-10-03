import type { CellValue } from '@bilig/protocol'
import {
  editWorkPaperAxisIntervals,
  moveWorkPaperAxis,
  setWorkPaperAxisOrder,
  swapWorkPaperAxisIndexes,
} from './work-paper-axis-helpers.js'
import {
  setWorkPaperCellContents,
  setWorkPaperCellValues,
  setWorkPaperSheetCellValues,
  setWorkPaperSheetRangeValues,
} from './work-paper-cell-content-setter.js'
import type { WorkPaperSheetFormulas, WorkPaperSheetValues } from './work-paper-sheet-read.js'
import type { WorkPaperRuntimeAdapters } from './work-paper-runtime-adapters.js'
import type {
  RawCellContent,
  SerializedWorkPaperNamedExpression,
  WorkPaperAddressFormatOptions,
  WorkPaperAddressLike,
  WorkPaperAxisInterval,
  WorkPaperAxisSwapMapping,
  WorkPaperCellAddress,
  WorkPaperCellRange,
  WorkPaperCellValueUpdate,
  WorkPaperCellType,
  WorkPaperCellValueDetailedType,
  WorkPaperCellValueType,
  WorkPaperChange,
  WorkPaperNamedExpression,
  WorkPaperRangeValueBlock,
  WorkPaperSheet,
  WorkPaperSheetCellValueUpdate,
  WorkPaperSheetDimensions,
  WorkPaperSheetRangeValues,
} from './work-paper-types.js'

export abstract class WorkPaperCapabilitySurface {
  protected abstract readonly runtimeAdapters: WorkPaperRuntimeAdapters

  isItPossibleToSetCellContents(address: WorkPaperCellAddress, content?: RawCellContent | WorkPaperSheet): boolean
  isItPossibleToSetCellContents(range: WorkPaperCellRange): boolean
  isItPossibleToSetCellContents(addressOrRange: WorkPaperAddressLike, content?: RawCellContent | WorkPaperSheet): boolean {
    return this.runtimeAdapters.capabilityOperations.isItPossibleToSetCellContents(addressOrRange, content)
  }

  isItPossibleToSwapRowIndexes(sheetId: number, rowA: number, rowB: number): boolean
  isItPossibleToSwapRowIndexes(sheetId: number, rowMappings: readonly WorkPaperAxisSwapMapping[]): boolean
  isItPossibleToSwapRowIndexes(sheetId: number, rowAOrMappings: number | readonly WorkPaperAxisSwapMapping[], rowB?: number): boolean {
    return this.runtimeAdapters.capabilityOperations.isItPossibleToSwapRowIndexes(sheetId, rowAOrMappings, rowB)
  }

  isItPossibleToSetRowOrder(sheetId: number, rowOrder: readonly number[]): boolean {
    return this.runtimeAdapters.capabilityOperations.isItPossibleToSetRowOrder(sheetId, rowOrder)
  }

  isItPossibleToSwapColumnIndexes(sheetId: number, columnA: number, columnB: number): boolean
  isItPossibleToSwapColumnIndexes(sheetId: number, columnMappings: readonly WorkPaperAxisSwapMapping[]): boolean
  isItPossibleToSwapColumnIndexes(
    sheetId: number,
    columnAOrMappings: number | readonly WorkPaperAxisSwapMapping[],
    columnB?: number,
  ): boolean {
    return this.runtimeAdapters.capabilityOperations.isItPossibleToSwapColumnIndexes(sheetId, columnAOrMappings, columnB)
  }

  isItPossibleToSetColumnOrder(sheetId: number, columnOrder: readonly number[]): boolean {
    return this.runtimeAdapters.capabilityOperations.isItPossibleToSetColumnOrder(sheetId, columnOrder)
  }

  isItPossibleToAddRows(sheetId: number, start: number, count?: number): boolean
  isItPossibleToAddRows(sheetId: number, ...indexes: readonly WorkPaperAxisInterval[]): boolean
  isItPossibleToAddRows(
    sheetId: number,
    startOrInterval: number | WorkPaperAxisInterval,
    countOrInterval?: number | WorkPaperAxisInterval,
    ...restIntervals: readonly WorkPaperAxisInterval[]
  ): boolean {
    return this.runtimeAdapters.capabilityOperations.isItPossibleToAddRows(sheetId, startOrInterval, countOrInterval, restIntervals)
  }

  isItPossibleToRemoveRows(sheetId: number, start: number, count?: number): boolean
  isItPossibleToRemoveRows(sheetId: number, ...indexes: readonly WorkPaperAxisInterval[]): boolean
  isItPossibleToRemoveRows(
    sheetId: number,
    startOrInterval: number | WorkPaperAxisInterval,
    countOrInterval?: number | WorkPaperAxisInterval,
    ...restIntervals: readonly WorkPaperAxisInterval[]
  ): boolean {
    return this.runtimeAdapters.capabilityOperations.isItPossibleToRemoveRows(sheetId, startOrInterval, countOrInterval, restIntervals)
  }

  isItPossibleToAddColumns(sheetId: number, start: number, count?: number): boolean
  isItPossibleToAddColumns(sheetId: number, ...indexes: readonly WorkPaperAxisInterval[]): boolean
  isItPossibleToAddColumns(
    sheetId: number,
    startOrInterval: number | WorkPaperAxisInterval,
    countOrInterval?: number | WorkPaperAxisInterval,
    ...restIntervals: readonly WorkPaperAxisInterval[]
  ): boolean {
    return this.runtimeAdapters.capabilityOperations.isItPossibleToAddColumns(sheetId, startOrInterval, countOrInterval, restIntervals)
  }

  isItPossibleToRemoveColumns(sheetId: number, start: number, count?: number): boolean
  isItPossibleToRemoveColumns(sheetId: number, ...indexes: readonly WorkPaperAxisInterval[]): boolean
  isItPossibleToRemoveColumns(
    sheetId: number,
    startOrInterval: number | WorkPaperAxisInterval,
    countOrInterval?: number | WorkPaperAxisInterval,
    ...restIntervals: readonly WorkPaperAxisInterval[]
  ): boolean {
    return this.runtimeAdapters.capabilityOperations.isItPossibleToRemoveColumns(sheetId, startOrInterval, countOrInterval, restIntervals)
  }

  isItPossibleToMoveCells(source: WorkPaperCellRange, target: WorkPaperCellAddress): boolean {
    return this.runtimeAdapters.capabilityOperations.isItPossibleToMoveCells(source, target)
  }

  isItPossibleToMoveRows(sheetId: number, start: number, count: number, target: number): boolean {
    return this.runtimeAdapters.capabilityOperations.isItPossibleToMoveRows(sheetId, start, count, target)
  }

  isItPossibleToMoveColumns(sheetId: number, start: number, count: number, target: number): boolean {
    return this.runtimeAdapters.capabilityOperations.isItPossibleToMoveColumns(sheetId, start, count, target)
  }

  isItPossibleToAddSheet(sheetName: string): boolean {
    return this.runtimeAdapters.capabilityOperations.isItPossibleToAddSheet(sheetName)
  }

  isItPossibleToRemoveSheet(sheetId: number): boolean {
    return this.runtimeAdapters.capabilityOperations.isItPossibleToRemoveSheet(sheetId)
  }

  isItPossibleToClearSheet(sheetId: number): boolean {
    return this.runtimeAdapters.capabilityOperations.isItPossibleToClearSheet(sheetId)
  }

  isItPossibleToReplaceSheetContent(sheetId: number, content: WorkPaperSheet): boolean {
    return this.runtimeAdapters.capabilityOperations.isItPossibleToReplaceSheetContent(sheetId, content)
  }

  isItPossibleToRenameSheet(sheetId: number, nextName: string): boolean {
    return this.runtimeAdapters.capabilityOperations.isItPossibleToRenameSheet(sheetId, nextName)
  }

  isItPossibleToAddNamedExpression(expressionName: string, expression: RawCellContent, scope?: number): boolean {
    return this.runtimeAdapters.capabilityOperations.isItPossibleToAddNamedExpression(expressionName, expression, scope)
  }

  isItPossibleToChangeNamedExpression(expressionName: string, expression: RawCellContent, scope?: number): boolean {
    return this.runtimeAdapters.capabilityOperations.isItPossibleToChangeNamedExpression(expressionName, expression, scope)
  }

  isItPossibleToRemoveNamedExpression(expressionName: string, scope?: number): boolean {
    return this.runtimeAdapters.capabilityOperations.isItPossibleToRemoveNamedExpression(expressionName, scope)
  }

  copy(range: WorkPaperCellRange): CellValue[][] {
    return this.runtimeAdapters.clipboardOperations.copy(range)
  }

  cut(range: WorkPaperCellRange): CellValue[][] {
    return this.runtimeAdapters.clipboardOperations.cut(range)
  }

  paste(targetLeftCorner: WorkPaperCellAddress): WorkPaperChange[] {
    return this.runtimeAdapters.clipboardOperations.paste(targetLeftCorner)
  }

  isClipboardEmpty(): boolean {
    return this.runtimeAdapters.clipboardOperations.isClipboardEmpty()
  }

  clearClipboard(): void {
    this.runtimeAdapters.clipboardOperations.clearClipboard()
  }

  getFillRangeData(source: WorkPaperCellRange, target: WorkPaperCellRange, offsetsFromTarget = false): RawCellContent[][] {
    return this.runtimeAdapters.clipboardOperations.getFillRangeData(source, target, offsetsFromTarget)
  }

  getCellValue(address: WorkPaperCellAddress): CellValue {
    return this.runtimeAdapters.readOperations.getCellValue(address)
  }

  getCellFormula(address: WorkPaperCellAddress): string | undefined {
    return this.runtimeAdapters.readOperations.getCellFormula(address)
  }

  getCellHyperlink(address: WorkPaperCellAddress): string | undefined {
    return this.runtimeAdapters.readOperations.getCellHyperlink(address)
  }

  getCellSerialized(address: WorkPaperCellAddress): RawCellContent {
    return this.runtimeAdapters.readOperations.getCellSerialized(address)
  }

  getRangeValues(range: WorkPaperCellRange): CellValue[][] {
    return this.runtimeAdapters.readOperations.getRangeValues(range)
  }

  getRangeValueBlock(range: WorkPaperCellRange): WorkPaperRangeValueBlock {
    return this.runtimeAdapters.readOperations.getRangeValueBlock(range)
  }

  getRangeFormulas(range: WorkPaperCellRange): WorkPaperSheetFormulas {
    return this.runtimeAdapters.readOperations.getRangeFormulas(range)
  }

  getRangeSerialized(range: WorkPaperCellRange): RawCellContent[][] {
    return this.runtimeAdapters.readOperations.getRangeSerialized(range)
  }

  getSheetValues(sheetId: number): WorkPaperSheetValues {
    return this.runtimeAdapters.readOperations.getSheetValues(sheetId)
  }

  getSheetFormulas(sheetId: number): WorkPaperSheetFormulas {
    return this.runtimeAdapters.readOperations.getSheetFormulas(sheetId)
  }

  getSheetSerialized(sheetId: number): RawCellContent[][] {
    return this.runtimeAdapters.readOperations.getSheetSerialized(sheetId)
  }

  getAllSheetsValues(): Record<string, CellValue[][]> {
    return this.runtimeAdapters.readOperations.getAllSheetsValues()
  }

  getAllSheetsFormulas(): Record<string, WorkPaperSheetFormulas> {
    return this.runtimeAdapters.readOperations.getAllSheetsFormulas()
  }

  getAllSheetsSerialized(): Record<string, RawCellContent[][]> {
    return this.runtimeAdapters.readOperations.getAllSheetsSerialized()
  }

  getAllSheetsDimensions(): Record<string, WorkPaperSheetDimensions> {
    return this.runtimeAdapters.readOperations.getAllSheetsDimensions()
  }

  getSheetDimensions(sheetId: number): WorkPaperSheetDimensions {
    return this.runtimeAdapters.readOperations.getSheetDimensions(sheetId)
  }

  simpleCellAddressFromString(value: string, defaultSheetId?: number): WorkPaperCellAddress | undefined {
    return this.runtimeAdapters.readOperations.simpleCellAddressFromString(value, defaultSheetId)
  }

  simpleCellRangeFromString(value: string, defaultSheetId?: number): WorkPaperCellRange | undefined {
    return this.runtimeAdapters.readOperations.simpleCellRangeFromString(value, defaultSheetId)
  }

  simpleCellAddressToString(address: WorkPaperCellAddress, optionsOrContextSheetId: WorkPaperAddressFormatOptions | number = {}): string {
    return this.runtimeAdapters.readOperations.simpleCellAddressToString(address, optionsOrContextSheetId)
  }

  simpleCellRangeToString(range: WorkPaperCellRange, optionsOrContextSheetId: WorkPaperAddressFormatOptions | number = {}): string {
    return this.runtimeAdapters.readOperations.simpleCellRangeToString(range, optionsOrContextSheetId)
  }

  getCellType(address: WorkPaperCellAddress): WorkPaperCellType {
    return this.runtimeAdapters.readOperations.getCellType(address)
  }

  doesCellHaveSimpleValue(address: WorkPaperCellAddress): boolean {
    return this.runtimeAdapters.readOperations.doesCellHaveSimpleValue(address)
  }

  doesCellHaveFormula(address: WorkPaperCellAddress): boolean {
    return this.runtimeAdapters.readOperations.doesCellHaveFormula(address)
  }

  isCellEmpty(address: WorkPaperCellAddress): boolean {
    return this.runtimeAdapters.readOperations.isCellEmpty(address)
  }

  isCellPartOfArray(address: WorkPaperCellAddress): boolean {
    return this.runtimeAdapters.readOperations.isCellPartOfArray(address)
  }

  getCellValueType(address: WorkPaperCellAddress): WorkPaperCellValueType {
    return this.runtimeAdapters.readOperations.getCellValueType(address)
  }

  getCellValueDetailedType(address: WorkPaperCellAddress): WorkPaperCellValueDetailedType {
    return this.runtimeAdapters.readOperations.getCellValueDetailedType(address)
  }

  getCellValueFormat(address: WorkPaperCellAddress): string | undefined {
    return this.runtimeAdapters.readOperations.getCellValueFormat(address)
  }

  getNamedExpressionValue(name: string, scope?: number): CellValue | CellValue[][] | undefined {
    return this.runtimeAdapters.namedExpressionOperations.getNamedExpressionValue(name, scope)
  }

  getNamedExpressionFormula(name: string, scope?: number): string | undefined {
    return this.runtimeAdapters.namedExpressionOperations.getNamedExpressionFormula(name, scope)
  }

  getNamedExpression(name: string, scope?: number): WorkPaperNamedExpression | undefined {
    return this.runtimeAdapters.namedExpressionOperations.getNamedExpression(name, scope)
  }

  addNamedExpression(
    expressionName: string,
    expression: RawCellContent,
    scope?: number,
    options?: Record<string, string | number | boolean>,
  ): WorkPaperChange[] {
    return this.runtimeAdapters.namedExpressionOperations.addNamedExpression(expressionName, expression, scope, options)
  }

  changeNamedExpression(
    expressionName: string,
    expression: RawCellContent,
    scope?: number,
    options?: Record<string, string | number | boolean>,
  ): WorkPaperChange[] {
    return this.runtimeAdapters.namedExpressionOperations.changeNamedExpression(expressionName, expression, scope, options)
  }

  removeNamedExpression(expressionName: string, scope?: number): WorkPaperChange[] {
    return this.runtimeAdapters.namedExpressionOperations.removeNamedExpression(expressionName, scope)
  }

  listNamedExpressions(scope?: number): string[] {
    return this.runtimeAdapters.namedExpressionOperations.listNamedExpressions(scope)
  }

  getAllNamedExpressionsSerialized(): SerializedWorkPaperNamedExpression[] {
    return this.runtimeAdapters.namedExpressionOperations.getAllNamedExpressionsSerialized()
  }

  setCellContents(address: WorkPaperCellAddress, content: RawCellContent | WorkPaperSheet): WorkPaperChange[] {
    return setWorkPaperCellContents(this.runtimeAdapters.setCellContentsRuntime, address, content)
  }

  setCellValues(updates: readonly WorkPaperCellValueUpdate[]): WorkPaperChange[] {
    return setWorkPaperCellValues(this.runtimeAdapters.setCellContentsRuntime, updates)
  }

  setSheetCellValues(sheetId: number, updates: readonly WorkPaperSheetCellValueUpdate[]): WorkPaperChange[] {
    return setWorkPaperSheetCellValues(this.runtimeAdapters.setCellContentsRuntime, sheetId, updates)
  }

  setSheetRangeValues(sheetId: number, startRow: number, startCol: number, values: WorkPaperSheetRangeValues): WorkPaperChange[] {
    return setWorkPaperSheetRangeValues(this.runtimeAdapters.setCellContentsRuntime, sheetId, startRow, startCol, values)
  }

  swapRowIndexes(sheetId: number, rowA: number, rowB: number): WorkPaperChange[]
  swapRowIndexes(sheetId: number, rowMappings: readonly WorkPaperAxisSwapMapping[]): WorkPaperChange[]
  swapRowIndexes(sheetId: number, rowAOrMappings: number | readonly WorkPaperAxisSwapMapping[], rowB?: number): WorkPaperChange[] {
    return swapWorkPaperAxisIndexes(this.runtimeAdapters.axisEditRuntime, 'row', sheetId, rowAOrMappings, rowB)
  }

  setRowOrder(sheetId: number, rowOrder: readonly number[]): WorkPaperChange[] {
    return setWorkPaperAxisOrder(this.runtimeAdapters.axisEditRuntime, 'row', sheetId, rowOrder)
  }

  swapColumnIndexes(sheetId: number, columnA: number, columnB: number): WorkPaperChange[]
  swapColumnIndexes(sheetId: number, columnMappings: readonly WorkPaperAxisSwapMapping[]): WorkPaperChange[]
  swapColumnIndexes(sheetId: number, columnAOrMappings: number | readonly WorkPaperAxisSwapMapping[], columnB?: number): WorkPaperChange[] {
    return swapWorkPaperAxisIndexes(this.runtimeAdapters.axisEditRuntime, 'column', sheetId, columnAOrMappings, columnB)
  }

  setColumnOrder(sheetId: number, columnOrder: readonly number[]): WorkPaperChange[] {
    return setWorkPaperAxisOrder(this.runtimeAdapters.axisEditRuntime, 'column', sheetId, columnOrder)
  }

  addRows(sheetId: number, start: number, count?: number): WorkPaperChange[]
  addRows(sheetId: number, ...indexes: readonly WorkPaperAxisInterval[]): WorkPaperChange[]
  addRows(
    sheetId: number,
    startOrInterval: number | WorkPaperAxisInterval,
    countOrInterval?: number | WorkPaperAxisInterval,
    ...restIntervals: readonly WorkPaperAxisInterval[]
  ): WorkPaperChange[] {
    return editWorkPaperAxisIntervals(
      this.runtimeAdapters.axisEditRuntime,
      'row',
      'add',
      sheetId,
      startOrInterval,
      countOrInterval,
      restIntervals,
    )
  }

  removeRows(sheetId: number, start: number, count?: number): WorkPaperChange[]
  removeRows(sheetId: number, ...indexes: readonly WorkPaperAxisInterval[]): WorkPaperChange[]
  removeRows(
    sheetId: number,
    startOrInterval: number | WorkPaperAxisInterval,
    countOrInterval?: number | WorkPaperAxisInterval,
    ...restIntervals: readonly WorkPaperAxisInterval[]
  ): WorkPaperChange[] {
    return editWorkPaperAxisIntervals(
      this.runtimeAdapters.axisEditRuntime,
      'row',
      'remove',
      sheetId,
      startOrInterval,
      countOrInterval,
      restIntervals,
    )
  }

  addColumns(sheetId: number, start: number, count?: number): WorkPaperChange[]
  addColumns(sheetId: number, ...indexes: readonly WorkPaperAxisInterval[]): WorkPaperChange[]
  addColumns(
    sheetId: number,
    startOrInterval: number | WorkPaperAxisInterval,
    countOrInterval?: number | WorkPaperAxisInterval,
    ...restIntervals: readonly WorkPaperAxisInterval[]
  ): WorkPaperChange[] {
    return editWorkPaperAxisIntervals(
      this.runtimeAdapters.axisEditRuntime,
      'column',
      'add',
      sheetId,
      startOrInterval,
      countOrInterval,
      restIntervals,
    )
  }

  removeColumns(sheetId: number, start: number, count?: number): WorkPaperChange[]
  removeColumns(sheetId: number, ...indexes: readonly WorkPaperAxisInterval[]): WorkPaperChange[]
  removeColumns(
    sheetId: number,
    startOrInterval: number | WorkPaperAxisInterval,
    countOrInterval?: number | WorkPaperAxisInterval,
    ...restIntervals: readonly WorkPaperAxisInterval[]
  ): WorkPaperChange[] {
    return editWorkPaperAxisIntervals(
      this.runtimeAdapters.axisEditRuntime,
      'column',
      'remove',
      sheetId,
      startOrInterval,
      countOrInterval,
      restIntervals,
    )
  }

  moveRows(sheetId: number, start: number, count: number, target: number): WorkPaperChange[] {
    return moveWorkPaperAxis(this.runtimeAdapters.axisEditRuntime, 'row', sheetId, start, count, target)
  }

  moveColumns(sheetId: number, start: number, count: number, target: number): WorkPaperChange[] {
    return moveWorkPaperAxis(this.runtimeAdapters.axisEditRuntime, 'column', sheetId, start, count, target)
  }

  addSheet(sheetName?: string): string {
    return this.runtimeAdapters.sheetOperations.addSheet(sheetName)
  }

  moveSheet(sheetId: number, order: number): WorkPaperChange[] {
    return this.runtimeAdapters.sheetOperations.moveSheet(sheetId, order)
  }

  removeSheet(sheetId: number): WorkPaperChange[] {
    return this.runtimeAdapters.sheetOperations.removeSheet(sheetId)
  }

  clearSheet(sheetId: number): WorkPaperChange[] {
    return this.runtimeAdapters.sheetOperations.clearSheet(sheetId)
  }

  setSheetContent(sheetId: number, content: WorkPaperSheet): WorkPaperChange[] {
    return this.runtimeAdapters.sheetOperations.setSheetContent(sheetId, content)
  }

  renameSheet(sheetId: number, nextName: string): WorkPaperChange[] {
    return this.runtimeAdapters.sheetOperations.renameSheet(sheetId, nextName)
  }
}
