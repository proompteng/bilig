import type { CellRangeRef, SheetFormatRangeSnapshot, SheetStyleRangeSnapshot } from '@bilig/protocol'
import {
  rewriteAddressForStructuralTransform,
  rewriteFormulaForStructuralTransform,
  rewriteRangeForStructuralTransform,
  type StructuralAxisTransform,
} from '@bilig/formula'
import { mapStructuralBoundary } from '../../engine-structural-utils.js'
import { normalizeDefinedName } from '../../workbook-store.js'
import type { CreateEngineStructureServiceArgs } from './structure-service-types.js'

type StructureMetadataRewriteArgs = Pick<CreateEngineStructureServiceArgs, 'state' | 'clearOwnedPivot'>

export function rewriteDefinedNamesForStructuralTransform(
  args: StructureMetadataRewriteArgs,
  sheetName: string,
  transform: StructuralAxisTransform,
): Set<string> {
  const changedNames = new Set<string>()
  args.state.workbook.listDefinedNames().forEach((record) => {
    if (typeof record.value === 'string' && record.value.startsWith('=')) {
      const nextFormula = rewriteDefinedNameFormulaOrNull(record.value.slice(1), sheetName, transform)
      if (nextFormula === null) {
        return
      }
      if (`=${nextFormula}` !== record.value) {
        args.state.workbook.setDefinedName(record.name, `=${nextFormula}`)
      }
      return
    }
    if (typeof record.value !== 'object' || !record.value) {
      return
    }
    switch (record.value.kind) {
      case 'formula': {
        const nextFormula = rewriteDefinedNameFormulaOrNull(
          record.value.formula.startsWith('=') ? record.value.formula.slice(1) : record.value.formula,
          sheetName,
          transform,
        )
        if (nextFormula === null) {
          return
        }
        const nextValue = {
          ...record.value,
          formula: record.value.formula.startsWith('=') ? `=${nextFormula}` : nextFormula,
        }
        if (nextValue.formula !== record.value.formula) {
          args.state.workbook.setDefinedName(record.name, nextValue)
          changedNames.add(normalizeDefinedName(record.name))
        }
        return
      }
      case 'cell-ref': {
        if (record.value.sheetName !== sheetName) {
          return
        }
        const nextAddress = rewriteAddressForStructuralTransform(record.value.address, transform)
        if (!nextAddress) {
          args.state.workbook.deleteDefinedName(record.name)
          changedNames.add(normalizeDefinedName(record.name))
          return
        }
        if (nextAddress !== record.value.address) {
          args.state.workbook.setDefinedName(record.name, {
            ...record.value,
            address: nextAddress,
          })
          changedNames.add(normalizeDefinedName(record.name))
        }
        return
      }
      case 'range-ref': {
        if (record.value.sheetName !== sheetName) {
          return
        }
        const nextRange = rewriteRangeForStructuralTransform(record.value.startAddress, record.value.endAddress, transform)
        if (!nextRange) {
          args.state.workbook.deleteDefinedName(record.name)
          changedNames.add(normalizeDefinedName(record.name))
          return
        }
        if (nextRange.startAddress !== record.value.startAddress || nextRange.endAddress !== record.value.endAddress) {
          args.state.workbook.setDefinedName(record.name, {
            ...record.value,
            startAddress: nextRange.startAddress,
            endAddress: nextRange.endAddress,
          })
          changedNames.add(normalizeDefinedName(record.name))
        }
        return
      }
      case 'scalar':
      case 'structured-ref':
        return
    }
  })
  return changedNames
}

function rewriteDefinedNameFormulaOrNull(formula: string, sheetName: string, transform: StructuralAxisTransform): string | null {
  try {
    return rewriteFormulaForStructuralTransform(formula, sheetName, sheetName, transform)
  } catch {
    return null
  }
}

export function rewriteWorkbookMetadataForStructuralTransform(
  args: StructureMetadataRewriteArgs,
  sheetName: string,
  transform: StructuralAxisTransform,
): { changedTableNames: Set<string> } {
  const changedTableNames = new Set<string>()
  args.state.workbook
    .listTables()
    .filter((table) => table.sheetName === sheetName)
    .forEach((table) => {
      const range = rewriteRangeForStructuralTransform(table.startAddress, table.endAddress, transform)
      if (!range) {
        changedTableNames.add(table.name)
        args.state.workbook.deleteTable(table.name)
        return
      }
      changedTableNames.add(table.name)
      args.state.workbook.setTable({
        ...table,
        startAddress: range.startAddress,
        endAddress: range.endAddress,
      })
    })
  const mergeRanges = args.state.workbook.listMergeRanges(sheetName)
  const rewrittenMergeRanges: CellRangeRef[] = []
  mergeRanges.forEach((merge) => {
    const range = rewriteRangeForStructuralTransform(merge.startAddress, merge.endAddress, transform)
    if (!range) {
      return
    }
    rewrittenMergeRanges.push({
      ...merge,
      startAddress: range.startAddress,
      endAddress: range.endAddress,
    })
  })
  args.state.workbook.setMergeRanges(sheetName, rewrittenMergeRanges)
  args.state.workbook.listFilters(sheetName).forEach((filter) => {
    const range = rewriteRangeForStructuralTransform(filter.range.startAddress, filter.range.endAddress, transform)
    args.state.workbook.deleteFilter(sheetName, filter.range)
    if (range) {
      args.state.workbook.setFilter(sheetName, {
        ...filter.range,
        startAddress: range.startAddress,
        endAddress: range.endAddress,
      })
    }
  })
  args.state.workbook.listSorts(sheetName).forEach((sort) => {
    const range = rewriteRangeForStructuralTransform(sort.range.startAddress, sort.range.endAddress, transform)
    args.state.workbook.deleteSort(sheetName, sort.range)
    if (!range) {
      return
    }
    args.state.workbook.setSort(
      sheetName,
      { ...sort.range, startAddress: range.startAddress, endAddress: range.endAddress },
      sort.keys.map((key) => ({
        ...key,
        keyAddress: rewriteAddressForStructuralTransform(key.keyAddress, transform) ?? key.keyAddress,
      })),
    )
  })
  args.state.workbook.listDataValidations(sheetName).forEach((validation) => {
    const range = rewriteRangeForStructuralTransform(validation.range.startAddress, validation.range.endAddress, transform)
    args.state.workbook.deleteDataValidation(sheetName, validation.range)
    if (!range) {
      return
    }
    const nextValidation = structuredClone(validation)
    nextValidation.range = {
      ...validation.range,
      startAddress: range.startAddress,
      endAddress: range.endAddress,
    }
    if (nextValidation.rule.kind === 'list' && nextValidation.rule.source) {
      switch (nextValidation.rule.source.kind) {
        case 'cell-ref': {
          if (nextValidation.rule.source.sheetName !== sheetName) {
            break
          }
          const nextAddress = rewriteAddressForStructuralTransform(nextValidation.rule.source.address, transform)
          if (!nextAddress) {
            return
          }
          nextValidation.rule.source.address = nextAddress
          break
        }
        case 'range-ref': {
          if (nextValidation.rule.source.sheetName !== sheetName) {
            break
          }
          const nextSourceRange = rewriteRangeForStructuralTransform(
            nextValidation.rule.source.startAddress,
            nextValidation.rule.source.endAddress,
            transform,
          )
          if (!nextSourceRange) {
            return
          }
          nextValidation.rule.source.startAddress = nextSourceRange.startAddress
          nextValidation.rule.source.endAddress = nextSourceRange.endAddress
          break
        }
        case 'named-range':
        case 'structured-ref':
          break
      }
    }
    args.state.workbook.setDataValidation(nextValidation)
  })
  args.state.workbook.listConditionalFormats(sheetName).forEach((format) => {
    const range = rewriteRangeForStructuralTransform(format.range.startAddress, format.range.endAddress, transform)
    args.state.workbook.deleteConditionalFormat(format.id)
    if (!range) {
      return
    }
    args.state.workbook.setConditionalFormat({
      ...format,
      range: {
        ...format.range,
        startAddress: range.startAddress,
        endAddress: range.endAddress,
      },
    })
  })
  args.state.workbook.listRangeProtections(sheetName).forEach((protection) => {
    const range = rewriteRangeForStructuralTransform(protection.range.startAddress, protection.range.endAddress, transform)
    args.state.workbook.deleteRangeProtection(protection.id)
    if (!range) {
      return
    }
    args.state.workbook.setRangeProtection({
      ...protection,
      range: {
        ...protection.range,
        startAddress: range.startAddress,
        endAddress: range.endAddress,
      },
    })
  })
  args.state.workbook.listCommentThreads(sheetName).forEach((thread) => {
    const nextAddress = rewriteAddressForStructuralTransform(thread.address, transform)
    args.state.workbook.deleteCommentThread(sheetName, thread.address)
    if (!nextAddress) {
      return
    }
    args.state.workbook.setCommentThread({
      ...thread,
      address: nextAddress,
    })
  })
  args.state.workbook.listNotes(sheetName).forEach((note) => {
    const nextAddress = rewriteAddressForStructuralTransform(note.address, transform)
    args.state.workbook.deleteNote(sheetName, note.address)
    if (!nextAddress) {
      return
    }
    args.state.workbook.setNote({
      ...note,
      address: nextAddress,
    })
  })
  const rewrittenStyleRanges: SheetStyleRangeSnapshot[] = []
  const rewrittenFormatRanges: SheetFormatRangeSnapshot[] = []
  args.state.workbook.listStyleRanges(sheetName).forEach((record) => {
    const range = rewriteRangeForStructuralTransform(record.range.startAddress, record.range.endAddress, transform)
    if (!range) {
      return
    }
    rewrittenStyleRanges.push({
      range: {
        ...record.range,
        startAddress: range.startAddress,
        endAddress: range.endAddress,
      },
      styleId: record.styleId,
    })
  })
  args.state.workbook.setStyleRanges(sheetName, rewrittenStyleRanges)
  args.state.workbook.listFormatRanges(sheetName).forEach((record) => {
    const range = rewriteRangeForStructuralTransform(record.range.startAddress, record.range.endAddress, transform)
    if (!range) {
      return
    }
    rewrittenFormatRanges.push({
      range: {
        ...record.range,
        startAddress: range.startAddress,
        endAddress: range.endAddress,
      },
      formatId: record.formatId,
    })
  })
  args.state.workbook.setFormatRanges(sheetName, rewrittenFormatRanges)
  const freezePane = args.state.workbook.getFreezePane(sheetName)
  if (freezePane) {
    const nextRows = transform.axis === 'row' ? mapStructuralBoundary(freezePane.rows, transform) : freezePane.rows
    const nextCols = transform.axis === 'column' ? mapStructuralBoundary(freezePane.cols, transform) : freezePane.cols
    if (nextRows <= 0 && nextCols <= 0) {
      args.state.workbook.clearFreezePane(sheetName)
    } else {
      args.state.workbook.setFreezePane(sheetName, nextRows, nextCols)
    }
  }
  args.state.workbook.listPivots().forEach((pivot) => {
    const nextAddress = pivot.sheetName === sheetName ? rewriteAddressForStructuralTransform(pivot.address, transform) : pivot.address
    const nextSource =
      pivot.source.sheetName === sheetName
        ? rewriteRangeForStructuralTransform(pivot.source.startAddress, pivot.source.endAddress, transform)
        : { startAddress: pivot.source.startAddress, endAddress: pivot.source.endAddress }
    if (!nextAddress || !nextSource) {
      args.clearOwnedPivot(pivot)
      args.state.workbook.deletePivot(pivot.sheetName, pivot.address)
      return
    }
    if (nextAddress !== pivot.address) {
      args.clearOwnedPivot(pivot)
      args.state.workbook.deletePivot(pivot.sheetName, pivot.address)
    }
    args.state.workbook.setPivot({
      ...pivot,
      address: nextAddress,
      source: {
        ...pivot.source,
        startAddress: nextSource.startAddress,
        endAddress: nextSource.endAddress,
      },
    })
  })
  args.state.workbook.listCharts().forEach((chart) => {
    const nextAddress = chart.sheetName === sheetName ? rewriteAddressForStructuralTransform(chart.address, transform) : chart.address
    const nextSource =
      chart.source.sheetName === sheetName
        ? rewriteRangeForStructuralTransform(chart.source.startAddress, chart.source.endAddress, transform)
        : { startAddress: chart.source.startAddress, endAddress: chart.source.endAddress }
    if (!nextAddress || !nextSource) {
      args.state.workbook.deleteChart(chart.id)
      return
    }
    args.state.workbook.setChart({
      ...chart,
      address: nextAddress,
      source: {
        ...chart.source,
        startAddress: nextSource.startAddress,
        endAddress: nextSource.endAddress,
      },
    })
  })
  args.state.workbook.listImages().forEach((image) => {
    if (image.sheetName !== sheetName) {
      return
    }
    const nextAddress = rewriteAddressForStructuralTransform(image.address, transform)
    if (!nextAddress) {
      args.state.workbook.deleteImage(image.id)
      return
    }
    args.state.workbook.setImage({
      ...image,
      address: nextAddress,
    })
  })
  args.state.workbook.listShapes().forEach((shape) => {
    if (shape.sheetName !== sheetName) {
      return
    }
    const nextAddress = rewriteAddressForStructuralTransform(shape.address, transform)
    if (!nextAddress) {
      args.state.workbook.deleteShape(shape.id)
      return
    }
    args.state.workbook.setShape({
      ...shape,
      address: nextAddress,
    })
  })
  return { changedTableNames }
}
