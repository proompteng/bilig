import { describe, expect, it } from 'vitest'
import { ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import { SpreadsheetEngine } from '../index.js'

describe('engine OFFSET dynamic dependencies', () => {
  it('evaluates imported OFFSET formulas after MATCH-selected formula targets', async () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'Offset Import' },
      sheets: [
        {
          id: 1,
          name: 'Forecast',
          order: 0,
          cells: [
            { address: 'C24', value: 'Downside' },
            { address: 'B28', value: 'Upside' },
            { address: 'B29', value: 'Base' },
            { address: 'B30', value: 'Downside' },
            { address: 'I19', formula: 'OFFSET(I27,MATCH($C$24,$B$28:$B$30,0),0)' },
            { address: 'J19', formula: 'OFFSET(J27,MATCH($C$24,$B$28:$B$30,0),0)' },
            { address: 'K19', formula: 'OFFSET(K27,MATCH($C$24,$B$28:$B$30,0),0)' },
            { address: 'I28', formula: '10+1' },
            { address: 'I29', formula: '20+2' },
            { address: 'I30', formula: '40+4' },
            { address: 'J28', formula: '100+1' },
            { address: 'J29', formula: '200+2' },
            { address: 'J30', formula: '400+4' },
            { address: 'K28', formula: '1000+1' },
            { address: 'K29', formula: '2000+2' },
            { address: 'K30', formula: '4000+4' },
          ],
        },
      ],
    }

    const engine = new SpreadsheetEngine({ workbookName: 'offset-import' })
    await engine.ready()
    engine.importSnapshot(snapshot)

    expect(engine.getCellValue('Forecast', 'I19')).toMatchObject({ tag: ValueTag.Number, value: 44 })
    expect(engine.getCellValue('Forecast', 'J19')).toMatchObject({ tag: ValueTag.Number, value: 404 })
    expect(engine.getCellValue('Forecast', 'K19')).toMatchObject({ tag: ValueTag.Number, value: 4004 })
  })

  it('evaluates imported OFFSET aggregates after formula targets selected by cell-valued dimensions', async () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'Offset Dimension Import',
        metadata: {
          definedNames: [{ name: 'Window', value: { kind: 'cell-ref', sheetName: 'Analysis', address: 'V10' } }],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Analysis',
          order: 0,
          cells: [
            { address: 'V10', value: 3 },
            { address: 'W7', formula: 'STDEV.S(OFFSET(W$20,-2*Window-1,0,Window))' },
            { address: 'W13', formula: '10/10' },
            { address: 'W14', formula: '1+1' },
            { address: 'W15', formula: '6/2' },
          ],
        },
      ],
    }

    const engine = new SpreadsheetEngine({ workbookName: 'offset-dimension-import' })
    await engine.ready()
    engine.importSnapshot(snapshot)

    expect(engine.getCellValue('Analysis', 'W7')).toMatchObject({ tag: ValueTag.Number, value: 1 })
  })
})
