import { describe, expect, it, vi } from 'vitest'

import { exportWorkPaperDocument } from '../persistence.js'
import { buildN8nForecastWorkPaper, createN8nForecastProof } from '../n8n-forecast-proof.js'
import { n8nForecastServerHelpText, parseN8nForecastServerCliArgs } from '../n8n-forecast-server-cli.js'
import { createN8nWorkPaperEvaluationProof } from '../n8n-workpaper-evaluation-proof.js'
import { WorkPaper } from '../work-paper.js'

describe('n8n formula readback proof', () => {
  it('returns formula readback and restore proof for one input edit', () => {
    expect(
      createN8nForecastProof({
        address: 'B3',
        value: 0.4,
      }),
    ).toMatchObject({
      verified: true,
      editedCell: 'Inputs!B3',
      before: {
        expectedCustomers: 5,
        expectedArr: 60000,
        expansionArr: 66000,
        targetGap: -34000,
      },
      after: {
        expectedCustomers: 8,
        expectedArr: 96000,
        expansionArr: 105600,
        targetGap: 5600,
      },
      formulaContracts: {
        expectedCustomers: '=Inputs!B2*Inputs!B3',
        expectedArr: '=B2*Inputs!B4',
        expansionArr: '=B3*Inputs!B5',
        targetGap: '=B4-100000',
      },
      checks: {
        previousValue: 0.25,
        newValue: 0.4,
        formulasPersisted: true,
        restoredMatchesAfter: true,
        computedOutputChanged: true,
      },
    })
  })

  it('rejects edits outside the demo input cells', () => {
    expect(() =>
      createN8nForecastProof({
        address: 'C9',
        value: 0.4,
      }),
    ).toThrow('Editable input address must be one of B2, B3, B4, B5')
  })

  it('parses the local server CLI options', () => {
    expect(parseN8nForecastServerCliArgs([], {})).toEqual({
      help: false,
      host: '127.0.0.1',
      port: 4321,
    })
    expect(
      parseN8nForecastServerCliArgs(['--host', '0.0.0.0', '--port', '8787'], {
        BILIG_N8N_PORT: '9999',
      }),
    ).toEqual({
      help: false,
      host: '0.0.0.0',
      port: 8787,
    })
    expect(n8nForecastServerHelpText()).toContain('bilig-n8n-formula-server')
    expect(n8nForecastServerHelpText()).toContain('/api/workpaper/n8n/evaluate')
    expect(() => parseN8nForecastServerCliArgs(['--bad'], {})).toThrow('Unknown bilig-n8n-formula-server argument')
  })

  it('evaluates a caller-provided WorkPaper document for n8n workflows', () => {
    const workbook = buildN8nForecastWorkPaper()
    const document = exportWorkPaperDocument(workbook, { includeConfig: true })
    workbook.dispose()

    const proof = createN8nWorkPaperEvaluationProof({
      document,
      edits: [
        {
          cell: 'Inputs!B3',
          value: 0.4,
        },
      ],
      readCells: ['Summary!B3'],
    })

    expect(proof).toMatchObject({
      verified: true,
      editedCells: [
        {
          cell: 'Inputs!B3',
          previousValue: 0.25,
          newValue: 0.4,
        },
      ],
      readback: {
        before: [
          {
            cell: 'Summary!B3',
            displayValue: '60000',
          },
        ],
        after: [
          {
            cell: 'Summary!B3',
            displayValue: '96000',
          },
        ],
        restored: [
          {
            cell: 'Summary!B3',
            displayValue: '96000',
          },
        ],
      },
      checks: {
        restoredMatchesAfter: true,
        formulasPersisted: true,
        computedOutputChanged: true,
      },
      updatedDocument: {
        format: 'bilig.headless.work-paper.document.v1',
      },
    })
  })

  it('disposes temporary workbooks after both n8n proof flows', () => {
    const workbook = buildN8nForecastWorkPaper()
    const document = exportWorkPaperDocument(workbook, { includeConfig: true })
    workbook.dispose()
    const dispose = vi.spyOn(WorkPaper.prototype, 'dispose')

    try {
      createN8nForecastProof({ address: 'B3', value: 0.4 })
      createN8nWorkPaperEvaluationProof({
        document,
        edits: [{ cell: 'Inputs!B3', value: 0.4 }],
      })

      expect(dispose).toHaveBeenCalledTimes(4)
    } finally {
      dispose.mockRestore()
    }
  })

  it('rejects generic n8n evaluation without a WorkPaper document', () => {
    expect(() =>
      createN8nWorkPaperEvaluationProof({
        edits: [
          {
            cell: 'Inputs!B3',
            value: 0.4,
          },
        ],
      }),
    ).toThrow('document is required')
  })
})
