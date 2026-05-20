import { describe, expect, it } from 'vitest'
import { ValueTag, type CellValue } from '@bilig/protocol'
import { makeCellEntity } from '../entity-ids.js'
import { DirectFormulaIndexCollection } from '../engine/services/direct-formula-index-collection.js'
import {
  createOperationDirectPostRecalcMarkers,
  initialDirectScalarLinearDeltaClosureCapacity,
} from '../engine/services/operation-direct-post-recalc-markers.js'
import type { RuntimeDirectScalarDescriptor } from '../engine/runtime-state.js'

type MarkerArgs = Parameters<typeof createOperationDirectPostRecalcMarkers>[0]

const numberValue = (value: number): CellValue => ({ tag: ValueTag.Number, value })

function formula(directScalar: RuntimeDirectScalarDescriptor) {
  return {
    compiled: {
      deps: [],
      volatile: false,
      producesSpill: false,
    },
    directAggregate: undefined,
    directCriteria: undefined,
    directLookup: undefined,
    directScalar,
  }
}

function binaryScalar(
  operator: '+' | '-' | '*' | '/',
  leftCellIndex: number,
  right: { kind: 'cell'; cellIndex: number } | { kind: 'literal-number'; value: number },
): RuntimeDirectScalarDescriptor {
  return binaryScalarWithOperands(operator, { kind: 'cell', cellIndex: leftCellIndex }, right)
}

function binaryScalarWithOperands(
  operator: '+' | '-' | '*' | '/',
  left: { kind: 'cell'; cellIndex: number } | { kind: 'literal-number'; value: number },
  right: { kind: 'cell'; cellIndex: number } | { kind: 'literal-number'; value: number },
): RuntimeDirectScalarDescriptor {
  return {
    kind: 'binary',
    operator,
    left,
    right,
  }
}

function createMarkers(input: {
  readonly formulas: Map<number, ReturnType<typeof formula>>
  readonly numbers: Map<number, number>
  readonly singleDependents: Map<number, number>
  readonly entityDependents?: Map<number, Uint32Array>
  readonly canSkipAllDirectFormulaColumnVersions?: () => boolean
  readonly canSkipDirectFormulaColumnVersion?: (cellIndex: number) => boolean
}) {
  const tags: ValueTag[] = []
  input.numbers.forEach((_value, cellIndex) => {
    tags[cellIndex] = ValueTag.Number
  })
  const state: MarkerArgs['state'] = {
    workbook: {
      cellStore: {
        flags: [],
        tags,
        getValue: (cellIndex: number) => numberValue(input.numbers.get(cellIndex) ?? 0),
      },
    },
    formulas: {
      get: (cellIndex: number) => input.formulas.get(cellIndex),
    },
    strings: {
      get: (id: number) => `s${id}`,
    },
  }

  return createOperationDirectPostRecalcMarkers({
    state,
    getSingleEntityDependent: (entityId) => input.singleDependents.get(entityId) ?? -1,
    getEntityDependents: (entityId) => input.entityDependents?.get(entityId) ?? new Uint32Array(),
    hasNoCellDependents: () => true,
    canSkipAllDirectFormulaColumnVersions: input.canSkipAllDirectFormulaColumnVersions,
    canSkipDirectFormulaColumnVersion: input.canSkipDirectFormulaColumnVersion ?? (() => true),
    readDirectScalarCellNumber: (cellIndex) => input.numbers.get(cellIndex) ?? 0,
    directScalarCellNumericValue: (cellIndex) => input.numbers.get(cellIndex),
    directScalarCurrentResultMatchesCell: () => false,
    lookupCurrent: {
      canEvaluateDirectUniformLookupCurrentResultFromNumeric: () => false,
      tryDirectExactLookupCurrentResult: () => undefined,
      tryDirectUniformLookupCurrentResult: () => undefined,
      tryDirectUniformLookupCurrentResultFromNumeric: () => undefined,
    },
    scalarDeltaClosureLimit: 32,
  })
}

describe('operation direct post-recalc markers', () => {
  it('preallocates linear scalar closure buffers to the configured large-chain limit', () => {
    expect(initialDirectScalarLinearDeltaClosureCapacity(4096)).toBe(4097)
    expect(initialDirectScalarLinearDeltaClosureCapacity(32)).toBe(33)
    expect(initialDirectScalarLinearDeltaClosureCapacity(0)).toBe(1)
  })

  it('marks a linear scalar closure as validated constant deltas', () => {
    const formulas = new Map([
      [20, formula(binaryScalar('+', 10, { kind: 'literal-number', value: 1 }))],
      [30, formula(binaryScalar('+', 20, { kind: 'literal-number', value: 2 }))],
    ])
    const markers = createMarkers({
      formulas,
      numbers: new Map([
        [20, 3],
        [30, 5],
      ]),
      singleDependents: new Map([
        [makeCellEntity(10), 20],
        [makeCellEntity(20), 30],
        [makeCellEntity(30), -1],
      ]),
    })
    const collection = new DirectFormulaIndexCollection()

    expect(markers.tryMarkDirectScalarLinearDeltaClosure(10, numberValue(2), numberValue(5), collection)).toBe(true)

    expect(collection.size).toBe(2)
    expect(collection.getCellIndicesForRead()).toBeInstanceOf(Uint32Array)
    expect(collection.getDelta(20)).toBe(3)
    expect(collection.getDelta(30)).toBe(3)
    expect(collection.hasCompleteScalarDeltas()).toBe(true)
    expect(collection.hasValidatedScalarDeltaCells()).toBe(true)
    expect(collection.hasTrustedDirectScalarDeltaCells()).toBe(true)
  })

  it('uses the all-column-version skip gate for linear scalar closures', () => {
    let perCellSkipCalls = 0
    let allSkipCalls = 0
    const formulas = new Map([
      [20, formula(binaryScalar('+', 10, { kind: 'literal-number', value: 1 }))],
      [30, formula(binaryScalar('+', 20, { kind: 'literal-number', value: 2 }))],
      [40, formula(binaryScalar('+', 30, { kind: 'literal-number', value: 3 }))],
    ])
    const markers = createMarkers({
      formulas,
      numbers: new Map([
        [20, 3],
        [30, 5],
        [40, 8],
      ]),
      singleDependents: new Map([
        [makeCellEntity(10), 20],
        [makeCellEntity(20), 30],
        [makeCellEntity(30), 40],
        [makeCellEntity(40), -1],
      ]),
      canSkipAllDirectFormulaColumnVersions: () => {
        allSkipCalls += 1
        return true
      },
      canSkipDirectFormulaColumnVersion: () => {
        perCellSkipCalls += 1
        return false
      },
    })
    const collection = new DirectFormulaIndexCollection()

    expect(markers.tryMarkDirectScalarLinearDeltaClosure(10, numberValue(2), numberValue(5), collection)).toBe(true)

    expect(allSkipCalls).toBe(1)
    expect(perCellSkipCalls).toBe(0)
    expect(collection.hasValidatedScalarDeltaCells()).toBe(true)
    expect(collection.hasTrustedDirectScalarDeltaCells()).toBe(true)
  })

  it('marks right-input affine scalar closures without falling back to graph traversal', () => {
    const formulas = new Map([
      [20, formula(binaryScalarWithOperands('-', { kind: 'literal-number', value: 10 }, { kind: 'cell', cellIndex: 10 }))],
      [30, formula(binaryScalar('*', 20, { kind: 'literal-number', value: 2 }))],
    ])
    const markers = createMarkers({
      formulas,
      numbers: new Map([
        [20, 8],
        [30, 16],
      ]),
      singleDependents: new Map([
        [makeCellEntity(10), 20],
        [makeCellEntity(20), 30],
        [makeCellEntity(30), -1],
      ]),
    })
    const collection = new DirectFormulaIndexCollection()

    expect(markers.tryMarkDirectScalarLinearDeltaClosure(10, numberValue(2), numberValue(5), collection)).toBe(true)

    expect(collection.size).toBe(2)
    expect(collection.getDelta(20)).toBe(-3)
    expect(collection.getDelta(30)).toBe(-6)
    expect(collection.hasCompleteScalarDeltas()).toBe(true)
    expect(collection.hasValidatedScalarDeltaCells()).toBe(true)
    expect(collection.hasTrustedDirectScalarDeltaCells()).toBe(true)
  })

  it('falls back to graph scalar closure when the dependent path branches', () => {
    const formulas = new Map([
      [20, formula(binaryScalar('+', 10, { kind: 'literal-number', value: 1 }))],
      [30, formula(binaryScalar('+', 20, { kind: 'literal-number', value: 2 }))],
      [40, formula(binaryScalar('*', 20, { kind: 'literal-number', value: 2 }))],
    ])
    const markers = createMarkers({
      formulas,
      numbers: new Map([
        [20, 3],
        [30, 5],
        [40, 6],
      ]),
      singleDependents: new Map([
        [makeCellEntity(10), 20],
        [makeCellEntity(20), -2],
      ]),
      entityDependents: new Map([
        [makeCellEntity(10), new Uint32Array([20])],
        [makeCellEntity(20), new Uint32Array([30, 40])],
      ]),
    })
    const collection = new DirectFormulaIndexCollection()

    markers.markDirectScalarDeltaClosure(10, numberValue(2), numberValue(5), collection)

    expect(collection.size).toBe(3)
    expect(collection.getDelta(20)).toBe(3)
    expect(collection.getDelta(30)).toBe(3)
    expect(collection.getDelta(40)).toBe(6)
  })
})
