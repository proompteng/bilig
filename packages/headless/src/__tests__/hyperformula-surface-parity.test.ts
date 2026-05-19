import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import {
  extractClassSurface,
  extractInterfaceKeys,
  parseHyperFormulaSurfaceSnapshot,
} from '../../../../scripts/workpaper-surface-contract.js'

const ALLOWED_BILIG_INSTANCE_METHODS = [
  'calculateScalarFormula',
  'compileScalarFormula',
  'dispose',
  'exportSnapshot',
  'getCalculationSettings',
  'getCellDisplayValue',
  'getCellFormulaDiagnostics',
  'getPerformanceCounters',
  'getRangeValueBlock',
  'offDetailed',
  'onDetailed',
  'onceDetailed',
  'resetPerformanceCounters',
  'setCalculationSettings',
  'setCellValues',
  'setSheetCellValues',
  'setSheetRangeValues',
  'transaction',
] as const
const ALLOWED_BILIG_INSTANCE_ACCESSORS = ['internals'] as const
const ALLOWED_BILIG_STATIC_METHODS = ['buildFromSheetEntries', 'buildFromSnapshot'] as const
const ALLOWED_BILIG_CONFIG_KEYS = ['calculationSettings', 'evaluationTimeoutMs'] as const

describe('WorkPaper HyperFormula snapshot parity', () => {
  it('matches the checked-in HyperFormula class surface snapshot', () => {
    const snapshot = loadSnapshot()
    const currentSurface = extractClassSurface(
      [
        readFileSync(new URL('../work-paper-runtime.ts', import.meta.url), 'utf8'),
        readFileSync(new URL('../work-paper-runtime-fast-path-base.ts', import.meta.url), 'utf8'),
        readFileSync(new URL('../work-paper-runtime-surface.ts', import.meta.url), 'utf8'),
        readFileSync(new URL('../work-paper-runtime-metadata-surface.ts', import.meta.url), 'utf8'),
        readFileSync(new URL('../work-paper-public-surface.ts', import.meta.url), 'utf8'),
        readFileSync(new URL('../work-paper-capability-surface.ts', import.meta.url), 'utf8'),
      ].join('\n'),
      'WorkPaper',
    )

    expect(currentSurface.staticMembers).toEqual(snapshot.classSurface.staticMembers)
    expect(currentSurface.staticMethods).toEqual(
      [...new Set([...snapshot.classSurface.staticMethods, ...ALLOWED_BILIG_STATIC_METHODS])].toSorted(),
    )
    expect(currentSurface.instanceAccessors).toEqual(
      [...new Set([...snapshot.classSurface.instanceAccessors, ...ALLOWED_BILIG_INSTANCE_ACCESSORS])].toSorted(),
    )
    expect(currentSurface.instanceMethods).toEqual(
      [...new Set([...snapshot.classSurface.instanceMethods, ...ALLOWED_BILIG_INSTANCE_METHODS])].toSorted(),
    )
  })

  it('matches the checked-in HyperFormula config-key snapshot', () => {
    const snapshot = loadSnapshot()
    const currentConfigKeys = extractInterfaceKeys(
      readFileSync(new URL('../work-paper-types.ts', import.meta.url), 'utf8'),
      'WorkPaperConfig',
    )

    expect(currentConfigKeys).toEqual([...new Set([...snapshot.configKeys, ...ALLOWED_BILIG_CONFIG_KEYS])].toSorted())
  })
})

function loadSnapshot() {
  return parseHyperFormulaSurfaceSnapshot(readFileSync(new URL('./fixtures/hyperformula-surface.json', import.meta.url), 'utf8'))
}
