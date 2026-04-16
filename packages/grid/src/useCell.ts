import { useSyncExternalStore } from 'react'
import { selectors, type SpreadsheetEngine } from '@bilig/core'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'

function snapshotSignature(snapshot: CellSnapshot): string {
  let valueKey = `${snapshot.value.tag}`
  switch (snapshot.value.tag) {
    case ValueTag.Number:
      valueKey = `n:${snapshot.value.value}`
      break
    case ValueTag.Boolean:
      valueKey = `b:${snapshot.value.value ? 1 : 0}`
      break
    case ValueTag.String:
      valueKey = `s:${snapshot.value.stringId}:${snapshot.value.value}`
      break
    case ValueTag.Error:
      valueKey = `e:${snapshot.value.code}`
      break
    case ValueTag.Empty:
      valueKey = 'empty'
      break
  }

  return [
    snapshot.version,
    snapshot.flags,
    snapshot.formula ?? '',
    snapshot.format ?? '',
    snapshot.styleId ?? '',
    snapshot.input ?? '',
    valueKey,
  ].join('|')
}

export function useCell(engine: SpreadsheetEngine, sheetName: string, addr: string) {
  const getRevision = () => {
    const snapshot = selectors.selectCellSnapshot(engine, sheetName, addr)
    return snapshotSignature(snapshot)
  }

  useSyncExternalStore((listener) => engine.subscribeCell(sheetName, addr, listener), getRevision, getRevision)

  return selectors.selectCellSnapshot(engine, sheetName, addr)
}
