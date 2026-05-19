import {
  readRuntimeImage,
  readRuntimeSnapshot,
  type EngineFormulaSourceRef,
  type EngineFormulaSourceRefs,
  type EngineFormulaSourceRefTable,
  type SpreadsheetEngine,
} from '@bilig/core/headless-runtime'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { loadInitialLiteralSheet, prepareInitialMixedSheetLoad } from './initial-sheet-load.js'
import { normalizeConfiguredWorkPaperCalculationSettings } from './work-paper-config.js'
import {
  inspectRuntimeSnapshotSheetDimensionsWithinLimits,
  inspectSheetWithinLimits,
  runtimeSnapshotMatchesSheetEntries,
  workbookSnapshotSheetHasDynamicSpillFormula,
  type WorkPaperSheetInspection,
} from './work-paper-sheet-inspection.js'
import type {
  SerializedWorkPaperNamedExpression,
  WorkPaperConfig,
  WorkPaperSheet,
  WorkPaperSheetDimensions,
  WorkPaperSheets,
} from './work-paper-types.js'

type RuntimeSnapshot = NonNullable<ReturnType<typeof readRuntimeSnapshot>>
type RuntimeImage = NonNullable<ReturnType<typeof readRuntimeImage>>

export function initializeWorkPaperFromSheets(args: {
  readonly engine: SpreadsheetEngine
  readonly config: WorkPaperConfig
  readonly sheets: WorkPaperSheets
  readonly namedExpressions: readonly SerializedWorkPaperNamedExpression[]
  readonly hasNamedExpressions: () => boolean
  readonly hasFunctionAliases: () => boolean
  readonly withEngineEventCaptureDisabled: (callback: () => void) => void
  readonly upsertNamedExpression: (expression: SerializedWorkPaperNamedExpression, options: { duringInitialization: boolean }) => void
  readonly rewriteFormulaForStorage: (formula: string, sheetId: number) => string
  readonly requireSheetId: (name: string) => number
  readonly cacheInitializedSheetDimensions: (
    sheetId: number,
    dimensions: WorkPaperSheetDimensions,
    options?: { readonly mayResizeDynamically?: boolean },
  ) => void
  readonly clearHistoryStacks: () => void
  readonly resetChangeTrackingCaches: () => void
}): void {
  initializeWorkPaperFromSheetEntries({
    ...args,
    sheetEntries: Object.entries(args.sheets),
    runtimeSnapshot: args.namedExpressions.length === 0 && !args.hasFunctionAliases() ? readRuntimeSnapshot(args.sheets) : undefined,
  })
}

export function initializeWorkPaperFromSheetEntries(args: {
  readonly engine: SpreadsheetEngine
  readonly config: WorkPaperConfig
  readonly sheetEntries: readonly (readonly [string, WorkPaperSheet])[]
  readonly runtimeSnapshot?: ReturnType<typeof readRuntimeSnapshot>
  readonly namedExpressions: readonly SerializedWorkPaperNamedExpression[]
  readonly hasNamedExpressions: () => boolean
  readonly hasFunctionAliases: () => boolean
  readonly withEngineEventCaptureDisabled: (callback: () => void) => void
  readonly upsertNamedExpression: (expression: SerializedWorkPaperNamedExpression, options: { duringInitialization: boolean }) => void
  readonly rewriteFormulaForStorage: (formula: string, sheetId: number) => string
  readonly requireSheetId: (name: string) => number
  readonly cacheInitializedSheetDimensions: (
    sheetId: number,
    dimensions: WorkPaperSheetDimensions,
    options?: { readonly mayResizeDynamically?: boolean },
  ) => void
  readonly clearHistoryStacks: () => void
  readonly resetChangeTrackingCaches: () => void
}): void {
  const sheetEntries = args.sheetEntries
  const runtimeSnapshot = args.runtimeSnapshot
  const runtimeSnapshotMatchesSheets = runtimeSnapshot !== undefined && runtimeSnapshotMatchesSheetEntries(sheetEntries, runtimeSnapshot)
  const runtimeSnapshotSheetsByName = runtimeSnapshotMatchesSheets
    ? new Map(runtimeSnapshot.sheets.map((sheet) => [sheet.name, sheet] as const))
    : undefined
  const runtimeImage = runtimeSnapshotMatchesSheets && runtimeSnapshot ? readRuntimeImage(runtimeSnapshot) : undefined
  const runtimeImageSheetCellsByName = runtimeImage
    ? new Map((runtimeImage.sheetCells ?? []).map((sheet) => [sheet.sheetName, sheet] as const))
    : undefined
  const runtimeImageDynamicSpillSheetsByName = collectRuntimeImageDynamicSpillSheets(runtimeImage)
  const inspectedSheets = inspectWorkPaperInitialSheets({
    sheetEntries,
    config: args.config,
    runtimeSnapshotSheetsByName,
    runtimeImageSheetCellsByName,
    runtimeImageDynamicSpillSheetsByName,
  })

  args.withEngineEventCaptureDisabled(() => {
    if (runtimeSnapshot && runtimeSnapshotMatchesSheets) {
      args.engine.importSnapshot(runtimeSnapshot)
      reapplyConfiguredCalculationSettings(args.engine, args.config)
    } else {
      const sheetIds = sheetEntries.map(([sheetName]) => args.engine.createSheetForInitialization(sheetName))
      args.namedExpressions.forEach((expression) => {
        args.upsertNamedExpression(expression, { duringInitialization: true })
      })
      let initialFormulaRefs: EngineFormulaSourceRefs | undefined
      let initialFormulaPotentialNewCells = 0
      for (let index = 0; index < sheetEntries.length; index += 1) {
        const [, sheet] = sheetEntries[index]!
        const sheetId = sheetIds[index]!
        const inspected = inspectedSheets[index]!
        if (!inspected.hasFormula) {
          loadInitialLiteralSheet(args.engine, sheetId, sheet, inspected)
          continue
        }
        const rewriteInitialFormula =
          !args.hasNamedExpressions() && !args.hasFunctionAliases()
            ? (formula: string) => formula
            : (formula: string) => args.rewriteFormulaForStorage(formula, sheetId)
        const prepared = prepareInitialMixedSheetLoad({
          engine: args.engine,
          sheetId,
          content: sheet,
          rewriteFormula: rewriteInitialFormula,
          inspection: inspected,
        })
        if (prepared.formulaRefs.length > 0) {
          initialFormulaRefs = appendInitialFormulaRefs(initialFormulaRefs, prepared.formulaRefs)
        }
        initialFormulaPotentialNewCells += prepared.potentialNewCells
      }
      if (initialFormulaRefs !== undefined && initialFormulaRefs.length > 0) {
        args.engine.initializeFormulaSourcesAtNow(initialFormulaRefs, initialFormulaPotentialNewCells)
      }
    }
    for (let index = 0; index < sheetEntries.length; index += 1) {
      const sheetId = args.requireSheetId(sheetEntries[index]![0])
      const inspected = inspectedSheets[index]
      if (inspected !== undefined) {
        args.cacheInitializedSheetDimensions(sheetId, inspected.dimensions, {
          mayResizeDynamically: inspected.hasDynamicSpillFormula,
        })
      }
    }
  })
  args.clearHistoryStacks()
  args.resetChangeTrackingCaches()
}

export function initializeWorkPaperFromSnapshot(args: {
  readonly engine: SpreadsheetEngine
  readonly config: WorkPaperConfig
  readonly snapshot: WorkbookSnapshot
  readonly withEngineEventCaptureDisabled: (callback: () => void) => void
  readonly requireSheetId: (name: string) => number
  readonly cacheInitializedSheetDimensions: (
    sheetId: number,
    dimensions: WorkPaperSheetDimensions,
    options?: { readonly mayResizeDynamically?: boolean },
  ) => void
  readonly clearHistoryStacks: () => void
  readonly resetChangeTrackingCaches: () => void
}): void {
  const runtimeImage = readRuntimeImage(args.snapshot)
  const runtimeImageSheetCellsByName = new Map((runtimeImage?.sheetCells ?? []).map((sheet) => [sheet.sheetName, sheet] as const))
  const runtimeImageDynamicSpillSheetsByName = collectRuntimeImageDynamicSpillSheets(runtimeImage)
  const inspectedSheets = args.snapshot.sheets.map((snapshotSheet) =>
    inspectRuntimeSnapshotSheetDimensionsWithinLimits({
      sheetName: snapshotSheet.name,
      snapshotSheet,
      ...(runtimeImageSheetCellsByName.get(snapshotSheet.name) !== undefined
        ? { runtimeSheetCells: runtimeImageSheetCellsByName.get(snapshotSheet.name)! }
        : {}),
      config: args.config,
    }),
  )

  args.withEngineEventCaptureDisabled(() => {
    args.engine.importSnapshot(args.snapshot)
    reapplyConfiguredCalculationSettings(args.engine, args.config)
    for (let index = 0; index < args.snapshot.sheets.length; index += 1) {
      const snapshotSheet = args.snapshot.sheets[index]!
      const sheetId = args.requireSheetId(snapshotSheet.name)
      args.cacheInitializedSheetDimensions(sheetId, inspectedSheets[index]!, {
        mayResizeDynamically:
          runtimeImageDynamicSpillSheetsByName?.get(snapshotSheet.name) ?? workbookSnapshotSheetHasDynamicSpillFormula(snapshotSheet),
      })
    }
  })
  args.clearHistoryStacks()
  args.resetChangeTrackingCaches()
}

function reapplyConfiguredCalculationSettings(engine: SpreadsheetEngine, config: WorkPaperConfig): void {
  const calculationSettings = normalizeConfiguredWorkPaperCalculationSettings(config.calculationSettings, engine.getCalculationSettings())
  if (calculationSettings !== undefined) {
    engine.setCalculationSettings(calculationSettings)
  }
}

function inspectWorkPaperInitialSheets(args: {
  readonly sheetEntries: readonly (readonly [string, WorkPaperSheets[string]])[]
  readonly config: WorkPaperConfig
  readonly runtimeSnapshotSheetsByName: ReadonlyMap<string, RuntimeSnapshot['sheets'][number]> | undefined
  readonly runtimeImageSheetCellsByName: ReadonlyMap<string, NonNullable<RuntimeImage['sheetCells']>[number]> | undefined
  readonly runtimeImageDynamicSpillSheetsByName: ReadonlyMap<string, boolean> | undefined
}): WorkPaperSheetInspection[] {
  const inspectedSheets: WorkPaperSheetInspection[] = []
  for (let index = 0; index < args.sheetEntries.length; index += 1) {
    const [sheetName, sheet] = args.sheetEntries[index]!
    const snapshotSheet = args.runtimeSnapshotSheetsByName?.get(sheetName)
    inspectedSheets[index] = snapshotSheet
      ? (() => {
          const runtimeSheetCells = args.runtimeImageSheetCellsByName?.get(sheetName)
          const dimensions = inspectRuntimeSnapshotSheetDimensionsWithinLimits({
            sheetName,
            snapshotSheet,
            ...(runtimeSheetCells !== undefined ? { runtimeSheetCells } : {}),
            config: args.config,
          })
          return {
            hasFormula: false,
            hasDynamicSpillFormula:
              args.runtimeImageDynamicSpillSheetsByName?.get(sheetName) ?? workbookSnapshotSheetHasDynamicSpillFormula(snapshotSheet),
            dimensions,
            materializedCellCount: runtimeSheetCells?.cellCount ?? 0,
            maxColumnCount: dimensions.width,
            formulaCellCount: 0,
            allMaterializedCellsAreNumbers: false,
          }
        })()
      : inspectSheetWithinLimits(sheetName, sheet, args.config)
  }
  return inspectedSheets
}

function collectRuntimeImageDynamicSpillSheets(runtimeImage: RuntimeImage | undefined): ReadonlyMap<string, boolean> | undefined {
  if (!runtimeImage) {
    return undefined
  }
  const templateProducesSpillById = new Map(runtimeImage.templateBank.map((template) => [template.id, template.compiled.producesSpill]))
  const sheetHasDynamicSpill = new Map<string, boolean>()
  for (const formula of runtimeImage.formulaInstances) {
    if (formula.templateId === undefined) {
      return undefined
    }
    const producesSpill = templateProducesSpillById.get(formula.templateId)
    if (producesSpill === undefined) {
      return undefined
    }
    if (producesSpill) {
      sheetHasDynamicSpill.set(formula.sheetName, true)
    } else if (!sheetHasDynamicSpill.has(formula.sheetName)) {
      sheetHasDynamicSpill.set(formula.sheetName, false)
    }
  }
  return sheetHasDynamicSpill
}

class CombinedInitialFormulaSourceRefs implements EngineFormulaSourceRefTable {
  private readonly chunks: EngineFormulaSourceRefs[]
  private readonly starts: number[]
  private lastChunkIndex = 0

  length: number

  constructor(first: EngineFormulaSourceRefs, second: EngineFormulaSourceRefs) {
    this.chunks = [first, second]
    this.starts = [0, first.length]
    this.length = first.length + second.length
  }

  append(next: EngineFormulaSourceRefs): void {
    this.starts.push(this.length)
    this.chunks.push(next)
    this.length += next.length
  }

  at(index: number) {
    if (index < 0 || index >= this.length) {
      throw new RangeError(`Initial formula ref index out of bounds: ${index.toString()}`)
    }
    const cached = this.tryReadFromChunk(this.lastChunkIndex, index)
    if (cached !== undefined) {
      return cached
    }
    for (
      let chunkIndex = index >= this.starts[this.lastChunkIndex]! ? this.lastChunkIndex + 1 : 0;
      chunkIndex < this.chunks.length;
      chunkIndex += 1
    ) {
      const ref = this.tryReadFromChunk(chunkIndex, index)
      if (ref !== undefined) {
        return ref
      }
    }
    throw new RangeError(`Initial formula ref index out of bounds: ${index.toString()}`)
  }

  private tryReadFromChunk(chunkIndex: number, index: number) {
    const chunk = this.chunks[chunkIndex]
    const start = this.starts[chunkIndex]
    if (chunk === undefined || start === undefined || index < start || index >= start + chunk.length) {
      return undefined
    }
    this.lastChunkIndex = chunkIndex
    return readInitialFormulaRef(chunk, index - start)
  }
}

function appendInitialFormulaRefs(existing: EngineFormulaSourceRefs | undefined, next: EngineFormulaSourceRefs): EngineFormulaSourceRefs {
  if (existing === undefined) {
    return next
  }
  if (existing instanceof CombinedInitialFormulaSourceRefs) {
    existing.append(next)
    return existing
  }
  return new CombinedInitialFormulaSourceRefs(existing, next)
}

function readInitialFormulaRef(refs: EngineFormulaSourceRefs, index: number): EngineFormulaSourceRef {
  if (Array.isArray(refs)) {
    return refs[index]!
  }
  return refs.at(index)!
}
