import { Effect } from 'effect'
import { ValueTag, type CellSnapshot, type WorkbookSnapshot } from '@bilig/protocol'
import type { EngineCellMutationRef } from '../../cell-mutations-at.js'
import { CellFlags } from '../../cell-store.js'
import { cloneCellStyleRecord } from '../../engine-style-utils.js'
import { exportSheetMetadata } from '../../engine-snapshot-utils.js'
import type { HydratedPreparedFormulaInitializationRef, PreparedFormulaInitializationRef } from './formula-initialization-service.js'
import type { FormulaInstanceSnapshot } from '../../formula/formula-instance-table.js'
import type { FormulaTemplateResolution, FormulaTemplateSnapshot } from '../../formula/template-bank.js'
import { exportReplicaSnapshot as exportReplicaStateSnapshot, hydrateReplicaState } from '../../replica-state.js'
import { attachRuntimeImage, readRuntimeImage } from '../../snapshot/runtime-image-codec.js'
import { restoreWorkbookFromRuntimeImage, restoreWorkbookFromSnapshot, type RuntimeImage } from '../../snapshot/runtime-image.js'
import type { EngineRuntimeState, EngineReplicaSnapshot } from '../runtime-state.js'
import { EngineSnapshotError } from '../errors.js'
import type { WorkbookPivotRecord } from '../../workbook-store.js'

export interface EngineSnapshotService {
  readonly exportWorkbook: () => Effect.Effect<WorkbookSnapshot, EngineSnapshotError>
  readonly importWorkbook: (snapshot: WorkbookSnapshot) => Effect.Effect<void, EngineSnapshotError>
  readonly exportReplica: () => Effect.Effect<EngineReplicaSnapshot, EngineSnapshotError>
  readonly importReplica: (snapshot: EngineReplicaSnapshot) => Effect.Effect<void, EngineSnapshotError>
}

export function createEngineSnapshotService(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook' | 'strings' | 'formulas' | 'replicaState' | 'entityVersions' | 'sheetDeleteVersions'>
  readonly getCellByIndex: (cellIndex: number) => CellSnapshot
  readonly resetWorkbook: (workbookName?: string) => void
  readonly exportTemplateBank?: () => readonly FormulaTemplateSnapshot[]
  readonly exportFormulaInstances?: () => readonly FormulaInstanceSnapshot[]
  readonly hydrateTemplateBank?: (templates: readonly FormulaTemplateSnapshot[]) => void
  readonly resolveTemplateById?: (templateId: number, source: string, row: number, col: number) => FormulaTemplateResolution | undefined
  readonly initializeCellFormulasAt: (refs: readonly EngineCellMutationRef[], potentialNewCells?: number) => void
  readonly initializePreparedCellFormulasAt?: (refs: readonly PreparedFormulaInitializationRef[], potentialNewCells?: number) => void
  readonly initializeHydratedPreparedCellFormulasAt?: (
    refs: readonly HydratedPreparedFormulaInitializationRef[],
    potentialNewCells?: number,
  ) => void
  readonly materializePivot?: (pivot: WorkbookPivotRecord) => number[]
  readonly emitFullInvalidation?: (options: { readonly incrementMetrics: boolean }) => void
}): EngineSnapshotService {
  return {
    exportWorkbook() {
      return Effect.try({
        try: () => {
          const workbook: WorkbookSnapshot['workbook'] = {
            name: args.state.workbook.workbookName,
          }
          const properties = args.state.workbook.listWorkbookProperties().map(({ key, value }) => ({ key, value }))
          const definedNames = args.state.workbook.listDefinedNames().map(({ name, value }) => ({ name, value }))
          const calculationSettings = args.state.workbook.getCalculationSettings()
          const volatileContext = args.state.workbook.getVolatileContext()
          const tables = args.state.workbook.listTables().map((table) => ({
            name: table.name,
            sheetName: table.sheetName,
            startAddress: table.startAddress,
            endAddress: table.endAddress,
            columnNames: [...table.columnNames],
            headerRow: table.headerRow,
            totalsRow: table.totalsRow,
          }))
          const spills = args.state.workbook.listSpills().map(({ sheetName, address, rows, cols }) => ({ sheetName, address, rows, cols }))
          const referencedStyleIds = new Set<string>()
          const referencedFormatIds = new Set<string>()
          args.state.workbook.sheetsByName.forEach((sheet) => {
            sheet.styleRanges.forEach((record) => referencedStyleIds.add(record.styleId))
            sheet.formatRanges.forEach((record) => referencedFormatIds.add(record.formatId))
            sheet.grid.forEachCell((cellIndex) => {
              const explicitFormat = args.state.workbook.getCellFormat(cellIndex)
              if (explicitFormat !== undefined) {
                referencedFormatIds.add(args.state.workbook.internCellNumberFormat(explicitFormat).id)
              }
            })
          })
          const styles = args.state.workbook
            .listCellStyles()
            .filter((style) => referencedStyleIds.has(style.id))
            .map((style) => cloneCellStyleRecord(style))
          const formats = args.state.workbook
            .listCellNumberFormats()
            .filter((format) => referencedFormatIds.has(format.id))
            .map((format) => Object.assign({}, format))
          const pivots = args.state.workbook.listPivots().map((pivot) => ({
            name: pivot.name,
            sheetName: pivot.sheetName,
            address: pivot.address,
            source: { ...pivot.source },
            groupBy: [...pivot.groupBy],
            values: pivot.values.map((value) => Object.assign({}, value)),
            rows: pivot.rows,
            cols: pivot.cols,
          }))
          const charts = args.state.workbook.listCharts().map((chart) => structuredClone(chart))
          const images = args.state.workbook.listImages().map((image) => structuredClone(image))
          const shapes = args.state.workbook.listShapes().map((shape) => structuredClone(shape))
          if (
            properties.length > 0 ||
            definedNames.length > 0 ||
            tables.length > 0 ||
            spills.length > 0 ||
            pivots.length > 0 ||
            charts.length > 0 ||
            images.length > 0 ||
            shapes.length > 0 ||
            styles.length > 0 ||
            formats.length > 0 ||
            calculationSettings.mode !== 'automatic' ||
            calculationSettings.compatibilityMode !== 'excel-modern' ||
            volatileContext.recalcEpoch !== 0
          ) {
            workbook.metadata = {}
            if (properties.length > 0) {
              workbook.metadata.properties = properties
            }
            if (definedNames.length > 0) {
              workbook.metadata.definedNames = definedNames
            }
            if (tables.length > 0) {
              workbook.metadata.tables = tables
            }
            if (spills.length > 0) {
              workbook.metadata.spills = spills
            }
            if (pivots.length > 0) {
              workbook.metadata.pivots = pivots
            }
            if (charts.length > 0) {
              workbook.metadata.charts = charts
            }
            if (images.length > 0) {
              workbook.metadata.images = images
            }
            if (shapes.length > 0) {
              workbook.metadata.shapes = shapes
            }
            if (styles.length > 0) {
              workbook.metadata.styles = styles
            }
            if (formats.length > 0) {
              workbook.metadata.formats = formats
            }
            if (calculationSettings.mode !== 'automatic' || calculationSettings.compatibilityMode !== 'excel-modern') {
              workbook.metadata.calculationSettings = calculationSettings
            }
            if (volatileContext.recalcEpoch !== 0) {
              workbook.metadata.volatileContext = volatileContext
            }
          }

          const runtimeImageSheetCells: Array<{
            sheetName: string
            coords: Array<{ row: number; col: number }>
            dimensions: { width: number; height: number }
            cellCount: number
          }> = []
          const workbookSnapshot: WorkbookSnapshot = {
            version: 1,
            workbook,
            sheets: [...args.state.workbook.sheetsByName.values()]
              .toSorted((left, right) => left.order - right.order)
              .map((sheet) => {
                const metadata = exportSheetMetadata(args.state.workbook, sheet.name)
                const cells: WorkbookSnapshot['sheets'][number]['cells'] = []
                const sheetCellCoords: Array<{ row: number; col: number }> = []
                let materializedWidth = 0
                let materializedHeight = 0
                sheet.grid.forEachCellEntry((cellIndex, row, col) => {
                  const cellSnapshot = args.getCellByIndex(cellIndex)
                  const explicitFormat = args.state.workbook.getCellFormat(cellIndex)
                  if ((cellSnapshot.flags & (CellFlags.SpillChild | CellFlags.PivotOutput)) !== 0) {
                    return
                  }
                  if (
                    cellSnapshot.formula === undefined &&
                    explicitFormat === undefined &&
                    (cellSnapshot.flags & CellFlags.AuthoredBlank) === 0 &&
                    (cellSnapshot.value.tag === ValueTag.Empty || cellSnapshot.value.tag === ValueTag.Error)
                  ) {
                    return
                  }
                  const cell: WorkbookSnapshot['sheets'][number]['cells'][number] = {
                    address: cellSnapshot.address,
                  }
                  if (explicitFormat !== undefined) {
                    cell.format = explicitFormat
                  }
                  if (cellSnapshot.formula) {
                    cell.formula = cellSnapshot.formula
                  } else if (cellSnapshot.value.tag === ValueTag.Number) {
                    cell.value = cellSnapshot.value.value
                  } else if (cellSnapshot.value.tag === ValueTag.Boolean) {
                    cell.value = cellSnapshot.value.value
                  } else if (cellSnapshot.value.tag === ValueTag.String) {
                    cell.value = cellSnapshot.value.value
                  } else {
                    cell.value = null
                  }
                  cells.push(cell)
                  sheetCellCoords.push({ row, col })
                  materializedHeight = Math.max(materializedHeight, row + 1)
                  materializedWidth = Math.max(materializedWidth, col + 1)
                })
                runtimeImageSheetCells.push({
                  sheetName: sheet.name,
                  coords: sheetCellCoords,
                  dimensions: { width: materializedWidth, height: materializedHeight },
                  cellCount: sheetCellCoords.length,
                })
                return metadata
                  ? { id: sheet.id, name: sheet.name, order: sheet.order, metadata, cells }
                  : { id: sheet.id, name: sheet.name, order: sheet.order, cells }
              }),
          }
          if (args.exportTemplateBank && args.exportFormulaInstances) {
            const formulaInstances = args.exportFormulaInstances()
            attachRuntimeImage(workbookSnapshot, {
              version: 1,
              templateBank: args.exportTemplateBank(),
              formulaInstances,
              formulaValues: formulaInstances.map((record) => ({
                sheetName: record.sheetName,
                row: record.row,
                col: record.col,
                value: args.getCellByIndex(record.cellIndex).value,
              })),
              sheetCells: runtimeImageSheetCells,
            } satisfies RuntimeImage)
          }
          return workbookSnapshot
        },
        catch: (cause) =>
          new EngineSnapshotError({
            message: 'Failed to export workbook snapshot',
            cause,
          }),
      })
    },
    importWorkbook(snapshot) {
      return Effect.try({
        try: () => {
          const materializeImportedPivots = (): void => {
            if (!args.materializePivot) {
              return
            }
            args.state.workbook.listPivots().forEach((pivot) => {
              args.materializePivot!(pivot)
            })
            snapshot.workbook.metadata?.pivots?.forEach((pivot) => {
              args.state.workbook.setPivot(pivot)
            })
          }

          const runtimeImage = readRuntimeImage(snapshot)
          if (runtimeImage && args.hydrateTemplateBank) {
            const restoreResult = restoreWorkbookFromRuntimeImage({
              snapshot,
              runtimeImage,
              workbook: args.state.workbook,
              strings: args.state.strings,
              resetWorkbook: args.resetWorkbook,
              hydrateTemplateBank: args.hydrateTemplateBank,
              initializeCellFormulasAt: args.initializeCellFormulasAt,
              ...(args.resolveTemplateById ? { resolveTemplateById: args.resolveTemplateById } : {}),
              ...(args.initializePreparedCellFormulasAt ? { initializePreparedCellFormulasAt: args.initializePreparedCellFormulasAt } : {}),
              ...(args.initializeHydratedPreparedCellFormulasAt
                ? { initializeHydratedPreparedCellFormulasAt: args.initializeHydratedPreparedCellFormulasAt }
                : {}),
            })
            materializeImportedPivots()
            args.emitFullInvalidation?.({ incrementMetrics: restoreResult.formulaCount === 0 })
            return
          }
          const restoreResult = restoreWorkbookFromSnapshot({
            snapshot,
            workbook: args.state.workbook,
            strings: args.state.strings,
            resetWorkbook: args.resetWorkbook,
            initializeCellFormulasAt: args.initializeCellFormulasAt,
          })
          materializeImportedPivots()
          args.emitFullInvalidation?.({ incrementMetrics: restoreResult.formulaCount === 0 })
        },
        catch: (cause) =>
          new EngineSnapshotError({
            message: 'Failed to import workbook snapshot',
            cause,
          }),
      })
    },
    exportReplica() {
      return Effect.try({
        try: () => ({
          replica: exportReplicaStateSnapshot(args.state.replicaState),
          entityVersions: [...args.state.entityVersions.entries()].map(([entityKey, order]) => ({
            entityKey,
            order,
          })),
          sheetDeleteVersions: [...args.state.sheetDeleteVersions.entries()].map(([sheetName, order]) => ({
            sheetName,
            order,
          })),
        }),
        catch: (cause) =>
          new EngineSnapshotError({
            message: 'Failed to export replica snapshot',
            cause,
          }),
      })
    },
    importReplica(snapshot) {
      return Effect.try({
        try: () => {
          hydrateReplicaState(args.state.replicaState, snapshot.replica)
          args.state.entityVersions.clear()
          snapshot.entityVersions.forEach(({ entityKey, order }) => {
            args.state.entityVersions.set(entityKey, order)
          })
          args.state.sheetDeleteVersions.clear()
          snapshot.sheetDeleteVersions.forEach(({ sheetName, order }) => {
            args.state.sheetDeleteVersions.set(sheetName, order)
          })
        },
        catch: (cause) =>
          new EngineSnapshotError({
            message: 'Failed to import replica snapshot',
            cause,
          }),
      })
    },
  }
}
