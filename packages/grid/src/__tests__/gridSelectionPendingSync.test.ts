import { describe, expect, test } from 'vitest'
import { resolveGridSelectionPendingSync } from '../gridSelectionPendingSync.js'
import type { GridSelectionSnapshot } from '../gridTypes.js'

function snapshot(address: string, endAddress = address): GridSelectionSnapshot {
  return {
    address,
    kind: address === endAddress ? 'cell' : 'range',
    range: {
      startAddress: address,
      endAddress,
    },
    sheetName: 'Sheet1',
  }
}

function columnSnapshot(address: string, startAddress: string, endAddress: string): GridSelectionSnapshot {
  return {
    address,
    kind: 'column',
    range: {
      startAddress,
      endAddress,
    },
    sheetName: 'Sheet1',
  }
}

function rowSnapshot(address: string, startAddress: string, endAddress: string): GridSelectionSnapshot {
  return {
    address,
    kind: 'row',
    range: {
      startAddress,
      endAddress,
    },
    sheetName: 'Sheet1',
  }
}

describe('resolveGridSelectionPendingSync', () => {
  test('keeps an optimistic local selection while the external snapshot is still the base selection', () => {
    const base = snapshot('C4')
    const pending = snapshot('B2', 'C3')

    expect(
      resolveGridSelectionPendingSync({
        currentSnapshot: pending,
        externalSnapshot: base,
        pendingBaseSnapshot: base,
        pendingLocalSnapshot: pending,
        sheetChanged: false,
      }),
    ).toEqual({
      keepCurrentSelection: true,
      pendingBaseSnapshot: base,
      pendingLocalSnapshot: pending,
    })
  })

  test('clears a confirmed optimistic selection when the external snapshot catches up', () => {
    const base = snapshot('C4')
    const pending = snapshot('B2', 'C3')

    expect(
      resolveGridSelectionPendingSync({
        currentSnapshot: pending,
        externalSnapshot: pending,
        pendingBaseSnapshot: base,
        pendingLocalSnapshot: pending,
        sheetChanged: false,
      }),
    ).toEqual({
      keepCurrentSelection: true,
      pendingBaseSnapshot: null,
      pendingLocalSnapshot: null,
    })
  })

  test('syncs to a confirmed pending selection when current state has not applied it yet', () => {
    const base = snapshot('C4')
    const pending = snapshot('B2', 'C3')

    expect(
      resolveGridSelectionPendingSync({
        currentSnapshot: snapshot('B2'),
        externalSnapshot: pending,
        pendingBaseSnapshot: base,
        pendingLocalSnapshot: pending,
        sheetChanged: false,
      }),
    ).toEqual({
      keepCurrentSelection: false,
      pendingBaseSnapshot: null,
      pendingLocalSnapshot: null,
    })
  })

  test('does not let stale pending state block a newer external selection', () => {
    const stalePending = snapshot('C4')
    const external = snapshot('B2', 'C3')

    expect(
      resolveGridSelectionPendingSync({
        currentSnapshot: stalePending,
        externalSnapshot: external,
        pendingBaseSnapshot: stalePending,
        pendingLocalSnapshot: stalePending,
        sheetChanged: false,
      }),
    ).toEqual({
      keepCurrentSelection: false,
      pendingBaseSnapshot: null,
      pendingLocalSnapshot: null,
    })
  })

  test('drops a pending range when a newer external cell selection arrives at the same active address', () => {
    const base = snapshot('C4')
    const pending = snapshot('B2', 'C3')
    const externalCell = snapshot('B2')

    expect(
      resolveGridSelectionPendingSync({
        currentSnapshot: pending,
        externalSnapshot: externalCell,
        pendingBaseSnapshot: base,
        pendingLocalSnapshot: pending,
        sheetChanged: false,
      }),
    ).toEqual({
      keepCurrentSelection: false,
      pendingBaseSnapshot: null,
      pendingLocalSnapshot: null,
    })
  })

  test('drops pending state on sheet changes', () => {
    const current = snapshot('B2')

    expect(
      resolveGridSelectionPendingSync({
        currentSnapshot: current,
        externalSnapshot: { ...snapshot('A1'), sheetName: 'Sheet2' },
        pendingBaseSnapshot: snapshot('A1'),
        pendingLocalSnapshot: current,
        sheetChanged: true,
      }),
    ).toEqual({
      keepCurrentSelection: false,
      pendingBaseSnapshot: null,
      pendingLocalSnapshot: null,
    })
  })

  test('preserves a pending row selection while the external selection is still the base cell', () => {
    const base = snapshot('E8')
    const pending = rowSnapshot('E2', 'A2', 'XFD4')

    expect(
      resolveGridSelectionPendingSync({
        currentSnapshot: pending,
        externalSnapshot: base,
        pendingBaseSnapshot: base,
        pendingLocalSnapshot: pending,
        sheetChanged: false,
      }),
    ).toEqual({
      keepCurrentSelection: true,
      pendingBaseSnapshot: base,
      pendingLocalSnapshot: pending,
    })
  })

  test('preserves a pending column selection while the external selection is still the base cell', () => {
    const base = snapshot('E8')
    const pending = columnSnapshot('B8', 'B1', 'D1048576')

    expect(
      resolveGridSelectionPendingSync({
        currentSnapshot: pending,
        externalSnapshot: base,
        pendingBaseSnapshot: base,
        pendingLocalSnapshot: pending,
        sheetChanged: false,
      }),
    ).toEqual({
      keepCurrentSelection: true,
      pendingBaseSnapshot: base,
      pendingLocalSnapshot: pending,
    })
  })

  test('clears pending row selection once the external snapshot catches up', () => {
    const base = snapshot('E8')
    const pending = rowSnapshot('E2', 'A2', 'XFD4')

    expect(
      resolveGridSelectionPendingSync({
        currentSnapshot: pending,
        externalSnapshot: pending,
        pendingBaseSnapshot: base,
        pendingLocalSnapshot: pending,
        sheetChanged: false,
      }),
    ).toEqual({
      keepCurrentSelection: true,
      pendingBaseSnapshot: null,
      pendingLocalSnapshot: null,
    })
  })

  test('clears pending column selection once the external snapshot catches up', () => {
    const base = snapshot('E8')
    const pending = columnSnapshot('B8', 'B1', 'D1048576')

    expect(
      resolveGridSelectionPendingSync({
        currentSnapshot: pending,
        externalSnapshot: pending,
        pendingBaseSnapshot: base,
        pendingLocalSnapshot: pending,
        sheetChanged: false,
      }),
    ).toEqual({
      keepCurrentSelection: true,
      pendingBaseSnapshot: null,
      pendingLocalSnapshot: null,
    })
  })
})
