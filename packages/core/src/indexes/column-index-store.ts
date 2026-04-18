import type { EngineRuntimeState } from '../engine/runtime-state.js'
import type { EngineRuntimeColumnStoreService } from '../engine/services/runtime-column-store-service.js'
import {
  applyLookupColumnOwnerLiteralWrite,
  buildLookupColumnOwner,
  type LookupColumnOwner,
  type LookupColumnOwnerWrite,
} from '../engine/services/lookup-column-owner.js'

export interface ColumnIndexStore {
  readonly getLookupColumnOwner: (request: { sheetName: string; col: number }) => LookupColumnOwner | undefined
  readonly invalidateColumn: (request: { sheetName: string; col: number }) => void
  readonly recordLiteralWrite: (request: LookupColumnOwnerWrite & { sheetName: string; col: number }) => void
}

function registryKey(sheetName: string, col: number): string {
  return `${sheetName}\t${col}`
}

function getCurrentColumnVersions(
  state: Pick<EngineRuntimeState, 'workbook'>,
  sheetName: string,
  col: number,
): {
  columnVersion: number
  structureVersion: number
  sheetColumnVersions: Uint32Array
} {
  const emptyColumnVersions = new Uint32Array(0)
  const sheet = state.workbook.getSheet(sheetName)
  const sheetColumnVersions = sheet?.columnVersions ?? emptyColumnVersions
  return {
    columnVersion: sheetColumnVersions[col] ?? 0,
    structureVersion: sheet?.structureVersion ?? 0,
    sheetColumnVersions,
  }
}

export function createColumnIndexStore(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook' | 'strings'>
  readonly runtimeColumnStore: EngineRuntimeColumnStoreService
}): ColumnIndexStore {
  const ownerIndices = new Map<string, LookupColumnOwner>()

  const ensureLookupColumnOwner = (sheetName: string, col: number): LookupColumnOwner | undefined => {
    const key = registryKey(sheetName, col)
    const currentVersions = getCurrentColumnVersions(args.state, sheetName, col)
    const existing = ownerIndices.get(key)
    if (
      existing &&
      existing.columnVersion === currentVersions.columnVersion &&
      existing.structureVersion === currentVersions.structureVersion &&
      existing.sheetColumnVersions === currentVersions.sheetColumnVersions
    ) {
      return existing
    }

    const owner = buildLookupColumnOwner({
      owner: args.runtimeColumnStore.getColumnOwner({ sheetName, col }),
      normalizeStringId: args.runtimeColumnStore.normalizeStringId,
    })
    if (owner) {
      ownerIndices.set(key, owner)
      return owner
    }
    ownerIndices.delete(key)
    return undefined
  }

  return {
    getLookupColumnOwner(request) {
      return ensureLookupColumnOwner(request.sheetName, request.col)
    },
    invalidateColumn(request) {
      ownerIndices.delete(registryKey(request.sheetName, request.col))
    },
    recordLiteralWrite(request) {
      const key = registryKey(request.sheetName, request.col)
      const owner = ownerIndices.get(key)
      if (!owner) {
        return
      }
      const currentVersions = getCurrentColumnVersions(args.state, request.sheetName, request.col)
      owner.columnVersion = currentVersions.columnVersion
      owner.structureVersion = currentVersions.structureVersion
      owner.sheetColumnVersions = currentVersions.sheetColumnVersions
      if (
        !applyLookupColumnOwnerLiteralWrite({
          owner,
          write: request,
          normalizeStringId: args.runtimeColumnStore.normalizeStringId,
        })
      ) {
        ownerIndices.delete(key)
      }
    },
  }
}
