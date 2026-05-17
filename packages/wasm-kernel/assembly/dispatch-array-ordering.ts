import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import { compareScalarValues } from './comparison'
import { copyInputCellToSpill } from './array-materialize'
import { inputCellScalarValue, inputCellTag, inputColsFromSlot, inputRowsFromSlot } from './operands'
import { coerceInteger } from './numeric-core'
import { scalarErrorAt } from './builtin-args'
import { STACK_KIND_ARRAY, STACK_KIND_RANGE, STACK_KIND_SCALAR, writeArrayResult, writeResult } from './result-io'
import { allocateSpillArrayResult } from './vm'

function coerceBoolean(tag: u8, value: f64): i32 {
  if (tag == ValueTag.Boolean || tag == ValueTag.Number) {
    return value != 0 ? 1 : 0
  }
  if (tag == ValueTag.Empty) {
    return 0
  }
  return -1
}

function resolveOneBasedIndex(rawIndex: i32, length: i32): i32 {
  if (rawIndex == i32.MIN_VALUE || rawIndex == 0 || length <= 0) {
    return i32.MIN_VALUE
  }
  const resolved = rawIndex < 0 ? length + rawIndex : rawIndex - 1
  return resolved >= 0 && resolved < length ? resolved : i32.MIN_VALUE
}

function writeValueError(
  base: i32,
  code: f64,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, code, rangeIndexStack, valueStack, tagStack, kindStack)
}

function compareInputCells(
  leftSlot: i32,
  leftRow: i32,
  leftCol: i32,
  rightSlot: i32,
  rightRow: i32,
  rightCol: i32,
  kindStack: Uint8Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): i32 {
  const leftTag = inputCellTag(
    leftSlot,
    leftRow,
    leftCol,
    kindStack,
    valueStack,
    tagStack,
    rangeIndexStack,
    rangeOffsets,
    rangeLengths,
    rangeRowCounts,
    rangeColCounts,
    rangeMembers,
    cellTags,
    cellNumbers,
  )
  const leftValue = inputCellScalarValue(
    leftSlot,
    leftRow,
    leftCol,
    kindStack,
    valueStack,
    tagStack,
    rangeIndexStack,
    rangeOffsets,
    rangeLengths,
    rangeRowCounts,
    rangeColCounts,
    rangeMembers,
    cellTags,
    cellNumbers,
    cellStringIds,
    cellErrors,
  )
  const rightTag = inputCellTag(
    rightSlot,
    rightRow,
    rightCol,
    kindStack,
    valueStack,
    tagStack,
    rangeIndexStack,
    rangeOffsets,
    rangeLengths,
    rangeRowCounts,
    rangeColCounts,
    rangeMembers,
    cellTags,
    cellNumbers,
  )
  const rightValue = inputCellScalarValue(
    rightSlot,
    rightRow,
    rightCol,
    kindStack,
    valueStack,
    tagStack,
    rangeIndexStack,
    rangeOffsets,
    rangeLengths,
    rangeRowCounts,
    rangeColCounts,
    rangeMembers,
    cellTags,
    cellNumbers,
    cellStringIds,
    cellErrors,
  )
  return compareScalarValues(
    leftTag,
    leftValue,
    rightTag,
    rightValue,
    null,
    stringOffsets,
    stringLengths,
    stringData,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  )
}

function copyOrderingCellToSpill(
  arrayIndex: u32,
  outputOffset: i32,
  slot: i32,
  row: i32,
  col: i32,
  kindStack: Uint8Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
): i32 {
  return copyInputCellToSpill(
    arrayIndex,
    outputOffset,
    slot,
    row,
    col,
    kindStack,
    valueStack,
    tagStack,
    rangeIndexStack,
    rangeOffsets,
    rangeLengths,
    rangeRowCounts,
    rangeColCounts,
    rangeMembers,
    cellTags,
    cellNumbers,
    cellStringIds,
    cellErrors,
  )
}

export function tryApplyArrayOrderingBuiltin(
  builtinId: i32,
  argc: i32,
  base: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): i32 {
  if (builtinId == BuiltinId.Choosecols && argc >= 2) {
    const sourceKind = kindStack[base]
    if (sourceKind != STACK_KIND_SCALAR && sourceKind != STACK_KIND_RANGE && sourceKind != STACK_KIND_ARRAY) {
      return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts)
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts)
    if (sourceRows <= 0 || sourceCols <= 0 || sourceRows == i32.MIN_VALUE || sourceCols == i32.MIN_VALUE) {
      return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, 1, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeValueError(base, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const selectedCols = new Array<i32>()
    for (let arg = 1; arg < argc; arg++) {
      const argumentSlot = base + arg
      if (kindStack[argumentSlot] != STACK_KIND_SCALAR) {
        return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      const selectedCol = resolveOneBasedIndex(coerceInteger(tagStack[argumentSlot], valueStack[argumentSlot]), sourceCols)
      if (selectedCol == i32.MIN_VALUE) {
        return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      selectedCols.push(selectedCol)
    }
    if (selectedCols.length == 0) {
      return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const outputCols = <i32>selectedCols.length
    const arrayIndex = allocateSpillArrayResult(sourceRows, outputCols)
    let outputOffset = 0
    for (let row = 0; row < sourceRows; row++) {
      for (let selectedColIndex = 0; selectedColIndex < outputCols; selectedColIndex++) {
        const selectedCol = selectedCols[selectedColIndex]
        const copyError = copyOrderingCellToSpill(
          arrayIndex,
          outputOffset,
          base,
          row,
          selectedCol,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
          cellStringIds,
          cellErrors,
        )
        if (copyError != ErrorCode.None) {
          return writeValueError(base, copyError, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        outputOffset += 1
      }
    }
    return writeArrayResult(base, arrayIndex, sourceRows, outputCols, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Chooserows && argc >= 2) {
    const sourceKind = kindStack[base]
    if (sourceKind != STACK_KIND_SCALAR && sourceKind != STACK_KIND_RANGE && sourceKind != STACK_KIND_ARRAY) {
      return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts)
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts)
    if (sourceRows <= 0 || sourceCols <= 0 || sourceRows == i32.MIN_VALUE || sourceCols == i32.MIN_VALUE) {
      return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, 1, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeValueError(base, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const selectedRows = new Array<i32>()
    for (let arg = 1; arg < argc; arg++) {
      const argumentSlot = base + arg
      if (kindStack[argumentSlot] != STACK_KIND_SCALAR) {
        return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      const selectedRow = resolveOneBasedIndex(coerceInteger(tagStack[argumentSlot], valueStack[argumentSlot]), sourceRows)
      if (selectedRow == i32.MIN_VALUE) {
        return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      selectedRows.push(selectedRow)
    }
    if (selectedRows.length == 0) {
      return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const outputRows = <i32>selectedRows.length
    const arrayIndex = allocateSpillArrayResult(outputRows, sourceCols)
    let outputOffset = 0
    for (let selectedRowIndex = 0; selectedRowIndex < outputRows; selectedRowIndex++) {
      const selectedRow = selectedRows[selectedRowIndex]
      for (let col = 0; col < sourceCols; col++) {
        const copyError = copyOrderingCellToSpill(
          arrayIndex,
          outputOffset,
          base,
          selectedRow,
          col,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
          cellStringIds,
          cellErrors,
        )
        if (copyError != ErrorCode.None) {
          return writeValueError(base, copyError, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        outputOffset += 1
      }
    }
    return writeArrayResult(base, arrayIndex, outputRows, sourceCols, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Sort && argc >= 1 && argc <= 4) {
    const sourceKind = kindStack[base]
    if (sourceKind != STACK_KIND_SCALAR && sourceKind != STACK_KIND_RANGE && sourceKind != STACK_KIND_ARRAY) {
      return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts)
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts)
    if (sourceRows <= 0 || sourceCols <= 0 || sourceRows == i32.MIN_VALUE || sourceCols == i32.MIN_VALUE) {
      return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeValueError(base, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const sortIndex = argc >= 2 ? coerceInteger(tagStack[base + 1], valueStack[base + 1]) : 1
    const sortOrder = argc >= 3 ? coerceInteger(tagStack[base + 2], valueStack[base + 2]) : 1
    const sortByColBoolean = argc >= 4 ? coerceBoolean(tagStack[base + 3], valueStack[base + 3]) : 0
    if (sortIndex == i32.MIN_VALUE || sortIndex < 1 || (sortOrder != 1 && sortOrder != -1) || sortByColBoolean < 0) {
      return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const sortByCol = sortByColBoolean != 0

    if (sourceRows == 1 || sourceCols == 1) {
      const length = sourceRows * sourceCols
      const order = new Array<i32>()
      for (let index = 0; index < length; index++) {
        order.push(index)
      }
      for (let index = 1; index < length; index++) {
        const current = order[index]
        const currentRow = current / sourceCols
        const currentCol = current - currentRow * sourceCols
        let cursor = index
        while (cursor > 0) {
          const previous = order[cursor - 1]
          const previousRow = previous / sourceCols
          const previousCol = previous - previousRow * sourceCols
          const comparison = compareInputCells(
            base,
            currentRow,
            currentCol,
            base,
            previousRow,
            previousCol,
            kindStack,
            valueStack,
            tagStack,
            rangeIndexStack,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
            cellTags,
            cellNumbers,
            cellStringIds,
            cellErrors,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
          if (comparison == i32.MIN_VALUE) {
            return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
          }
          if (comparison * sortOrder < 0) {
            order[cursor] = previous
            cursor -= 1
            continue
          }
          break
        }
        order[cursor] = current
      }
      const arrayIndex = allocateSpillArrayResult(sourceRows, sourceCols)
      for (let index = 0; index < length; index++) {
        const sourceOffset = order[index]
        const sourceRow = sourceOffset / sourceCols
        const sourceCol = sourceOffset - sourceRow * sourceCols
        const copyError = copyOrderingCellToSpill(
          arrayIndex,
          index,
          base,
          sourceRow,
          sourceCol,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
          cellStringIds,
          cellErrors,
        )
        if (copyError != ErrorCode.None) {
          return writeValueError(base, copyError, rangeIndexStack, valueStack, tagStack, kindStack)
        }
      }
      return writeArrayResult(base, arrayIndex, sourceRows, sourceCols, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    if (!sortByCol) {
      if (sortIndex > sourceCols) {
        return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      const rowSort = new Array<i32>()
      for (let row = 0; row < sourceRows; row++) {
        rowSort.push(row)
      }
      const sortCol = sortIndex - 1
      for (let cursor = 1; cursor < sourceRows; cursor++) {
        const currentRow = rowSort[cursor]
        let position = cursor
        while (position > 0) {
          const previousRow = rowSort[position - 1]
          const comparison = compareInputCells(
            base,
            currentRow,
            sortCol,
            base,
            previousRow,
            sortCol,
            kindStack,
            valueStack,
            tagStack,
            rangeIndexStack,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
            cellTags,
            cellNumbers,
            cellStringIds,
            cellErrors,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
          if (comparison == i32.MIN_VALUE) {
            return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
          }
          if (comparison * sortOrder < 0) {
            rowSort[position] = previousRow
            position -= 1
            continue
          }
          break
        }
        rowSort[position] = currentRow
      }
      const arrayIndex = allocateSpillArrayResult(sourceRows, sourceCols)
      let outputOffset = 0
      for (let sortedRow = 0; sortedRow < sourceRows; sortedRow++) {
        const sourceRow = rowSort[sortedRow]
        for (let col = 0; col < sourceCols; col++) {
          const copyError = copyOrderingCellToSpill(
            arrayIndex,
            outputOffset,
            base,
            sourceRow,
            col,
            kindStack,
            valueStack,
            tagStack,
            rangeIndexStack,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
            cellTags,
            cellNumbers,
            cellStringIds,
            cellErrors,
          )
          if (copyError != ErrorCode.None) {
            return writeValueError(base, copyError, rangeIndexStack, valueStack, tagStack, kindStack)
          }
          outputOffset += 1
        }
      }
      return writeArrayResult(base, arrayIndex, sourceRows, sourceCols, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    if (sortIndex > sourceRows) {
      return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const colSort = new Array<i32>()
    for (let col = 0; col < sourceCols; col++) {
      colSort.push(col)
    }
    const sortRow = sortIndex - 1
    for (let cursor = 1; cursor < sourceCols; cursor++) {
      const currentCol = colSort[cursor]
      let position = cursor
      while (position > 0) {
        const previousCol = colSort[position - 1]
        const comparison = compareInputCells(
          base,
          sortRow,
          currentCol,
          base,
          sortRow,
          previousCol,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
          cellStringIds,
          cellErrors,
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        )
        if (comparison == i32.MIN_VALUE) {
          return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        if (comparison * sortOrder < 0) {
          colSort[position] = previousCol
          position -= 1
          continue
        }
        break
      }
      colSort[position] = currentCol
    }
    const arrayIndex = allocateSpillArrayResult(sourceRows, sourceCols)
    let outputOffset = 0
    for (let row = 0; row < sourceRows; row++) {
      for (let sortedCol = 0; sortedCol < sourceCols; sortedCol++) {
        const sourceCol = colSort[sortedCol]
        const copyError = copyOrderingCellToSpill(
          arrayIndex,
          outputOffset,
          base,
          row,
          sourceCol,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
          cellStringIds,
          cellErrors,
        )
        if (copyError != ErrorCode.None) {
          return writeValueError(base, copyError, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        outputOffset += 1
      }
    }
    return writeArrayResult(base, arrayIndex, sourceRows, sourceCols, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Sortby && argc >= 2) {
    const sourceKind = kindStack[base]
    if (sourceKind != STACK_KIND_SCALAR && sourceKind != STACK_KIND_RANGE && sourceKind != STACK_KIND_ARRAY) {
      return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts)
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts)
    if (sourceRows <= 0 || sourceCols <= 0 || sourceRows == i32.MIN_VALUE || sourceCols == i32.MIN_VALUE) {
      return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const sourceLength = sourceRows * sourceCols
    if (sourceRows > 1 && sourceCols > 1) {
      return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeValueError(base, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const sortBySlots = new Array<i32>()
    const sortByLengths = new Array<i32>()
    const sortByCols = new Array<i32>()
    const sortByOrders = new Array<i32>()
    let arg = 1
    while (arg < argc) {
      const criterionSlot = base + arg
      const criterionKind = kindStack[criterionSlot]
      if (criterionKind != STACK_KIND_SCALAR && criterionKind != STACK_KIND_RANGE && criterionKind != STACK_KIND_ARRAY) {
        return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }

      const criterionRows = inputRowsFromSlot(criterionSlot, kindStack, rangeIndexStack, rangeRowCounts)
      const criterionCols = inputColsFromSlot(criterionSlot, kindStack, rangeIndexStack, rangeColCounts)
      if (criterionRows <= 0 || criterionCols <= 0 || criterionRows == i32.MIN_VALUE || criterionCols == i32.MIN_VALUE) {
        return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      const criterionLength = criterionRows * criterionCols
      if (criterionLength != 1 && criterionLength != sourceLength) {
        return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }

      sortBySlots.push(criterionSlot)
      sortByLengths.push(criterionLength)
      sortByCols.push(criterionCols)

      const nextSlot = criterionSlot + 1
      if (nextSlot < base + argc) {
        if (kindStack[nextSlot] == STACK_KIND_RANGE || kindStack[nextSlot] == STACK_KIND_ARRAY) {
          return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        if (kindStack[nextSlot] == STACK_KIND_SCALAR) {
          const requestedOrder = coerceInteger(tagStack[nextSlot], valueStack[nextSlot])
          if (requestedOrder != 1 && requestedOrder != -1) {
            return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
          }
          sortByOrders.push(requestedOrder)
          arg += 2
          continue
        }
      }

      sortByOrders.push(1)
      arg += 1
    }

    if (sortBySlots.length == 0) {
      return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const sourceIndexes = new Array<i32>()
    for (let offset = 0; offset < sourceLength; offset++) {
      sourceIndexes.push(offset)
    }

    for (let cursor = 1; cursor < sourceLength; cursor++) {
      const currentOffset = sourceIndexes[cursor]
      let position = cursor
      while (position > 0) {
        const previousOffset = sourceIndexes[position - 1]
        let comparison = 0
        for (let sortByIndex = 0; sortByIndex < sortBySlots.length; sortByIndex++) {
          const slot = sortBySlots[sortByIndex]
          const slotLength = sortByLengths[sortByIndex]
          const slotCols = sortByCols[sortByIndex]
          const slotOrder = sortByOrders[sortByIndex]

          const currentCriterionOffset = slotLength == 1 ? 0 : currentOffset
          const previousCriterionOffset = slotLength == 1 ? 0 : previousOffset
          const currentCriterionRow = slotLength == 1 ? 0 : currentCriterionOffset / slotCols
          const currentCriterionCol = slotLength == 1 ? 0 : currentCriterionOffset - currentCriterionRow * slotCols
          const previousCriterionRow = slotLength == 1 ? 0 : previousCriterionOffset / slotCols
          const previousCriterionCol = slotLength == 1 ? 0 : previousCriterionOffset - previousCriterionRow * slotCols

          const currentCriterionTag = inputCellTag(
            slot,
            currentCriterionRow,
            currentCriterionCol,
            kindStack,
            valueStack,
            tagStack,
            rangeIndexStack,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
            cellTags,
            cellNumbers,
          )
          const currentCriterionValue = inputCellScalarValue(
            slot,
            currentCriterionRow,
            currentCriterionCol,
            kindStack,
            valueStack,
            tagStack,
            rangeIndexStack,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
            cellTags,
            cellNumbers,
            cellStringIds,
            cellErrors,
          )
          const previousCriterionTag = inputCellTag(
            slot,
            previousCriterionRow,
            previousCriterionCol,
            kindStack,
            valueStack,
            tagStack,
            rangeIndexStack,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
            cellTags,
            cellNumbers,
          )
          const previousCriterionValue = inputCellScalarValue(
            slot,
            previousCriterionRow,
            previousCriterionCol,
            kindStack,
            valueStack,
            tagStack,
            rangeIndexStack,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
            cellTags,
            cellNumbers,
            cellStringIds,
            cellErrors,
          )
          const criterionComparison = compareScalarValues(
            currentCriterionTag,
            currentCriterionValue,
            previousCriterionTag,
            previousCriterionValue,
            null,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
          if (criterionComparison == i32.MIN_VALUE) {
            return writeValueError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
          }
          if (criterionComparison != 0) {
            comparison = criterionComparison * slotOrder
            break
          }
        }
        if (comparison < 0) {
          sourceIndexes[position] = previousOffset
          position -= 1
          continue
        }
        if (comparison > 0) {
          break
        }
        break
      }
      sourceIndexes[position] = currentOffset
    }

    const arrayIndex = allocateSpillArrayResult(sourceRows, sourceCols)
    let outputOffset = 0
    for (let index = 0; index < sourceLength; index++) {
      const sortedOffset = sourceIndexes[index]
      const sourceRow = sortedOffset / sourceCols
      const sourceCol = sortedOffset - sourceRow * sourceCols
      const copyError = copyOrderingCellToSpill(
        arrayIndex,
        outputOffset,
        base,
        sourceRow,
        sourceCol,
        kindStack,
        valueStack,
        tagStack,
        rangeIndexStack,
        rangeOffsets,
        rangeLengths,
        rangeRowCounts,
        rangeColCounts,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      )
      if (copyError != ErrorCode.None) {
        return writeValueError(base, copyError, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      outputOffset += 1
    }
    return writeArrayResult(base, arrayIndex, sourceRows, sourceCols, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  return -1
}
