import * as fc from 'fast-check'
import { SpreadsheetEngine } from '../engine.js'

export type MetadataStructuralAction =
  | { kind: 'insertRows'; sheetName: string; start: number; count: number }
  | { kind: 'deleteRows'; sheetName: string; start: number; count: number }
  | { kind: 'insertColumns'; sheetName: string; start: number; count: number }
  | { kind: 'deleteColumns'; sheetName: string; start: number; count: number }

export function metadataStructuralActionArbitrary(sheetNames: readonly string[]): fc.Arbitrary<MetadataStructuralAction> {
  return fc.oneof<MetadataStructuralAction>(
    structuralArbitrary('insertRows', sheetNames),
    structuralArbitrary('deleteRows', sheetNames),
    structuralArbitrary('insertColumns', sheetNames),
    structuralArbitrary('deleteColumns', sheetNames),
  )
}

export function applyMetadataStructuralAction(engine: SpreadsheetEngine, action: MetadataStructuralAction): void {
  switch (action.kind) {
    case 'insertRows':
      engine.insertRows(action.sheetName, action.start, action.count)
      return
    case 'deleteRows':
      engine.deleteRows(action.sheetName, action.start, action.count)
      return
    case 'insertColumns':
      engine.insertColumns(action.sheetName, action.start, action.count)
      return
    case 'deleteColumns':
      engine.deleteColumns(action.sheetName, action.start, action.count)
      return
  }
}

export async function restoreMetadataSnapshot(engine: SpreadsheetEngine, workbookName: string): Promise<SpreadsheetEngine> {
  const restored = new SpreadsheetEngine({
    workbookName,
    replicaId: workbookName,
  })
  await restored.ready()
  restored.importSnapshot(engine.exportSnapshot())
  return restored
}

function structuralArbitrary(
  kind: MetadataStructuralAction['kind'],
  sheetNames: readonly string[],
): fc.Arbitrary<MetadataStructuralAction> {
  return fc
    .record({
      sheetName: fc.constantFrom(...sheetNames),
      start: fc.integer({ min: 0, max: 3 }),
      count: fc.integer({ min: 1, max: 1 }),
    })
    .map((action) => Object.assign({ kind }, action) as MetadataStructuralAction)
}
