import { ValueTag } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'

type SheetCell = string | number | boolean | null

function numberValue(workbook: WorkPaper, ref: string): number {
  const address = workbook.simpleCellAddressFromString(ref)
  if (!address) {
    throw new Error(`Expected ${ref} to resolve`)
  }
  const value = workbook.getCellValue(address)
  if (value.tag !== ValueTag.Number) {
    throw new Error(`Expected ${ref} to be numeric, received ${JSON.stringify(value)}`)
  }
  return value.value
}

function expectedBalance(rows: number): number {
  let total = 0
  for (let row = 2; row <= rows; row += 1) {
    total += row % 17
  }
  return total
}

function expectedFilteredBalance(rows: number): number {
  let total = 0
  for (let row = 2; row <= Math.min(rows, 2_000); row += 1) {
    if (row % 2 === 0 && row % 3 === 0) {
      total += row % 17
    }
  }
  return -total
}

function buildGiftCardSheets(rows: number): Record<string, SheetCell[][]> {
  const summary = Array.from({ length: 6 }, () => Array<SheetCell>(2).fill(null))
  summary[0][0] = 'period'
  summary[0][1] = 46053
  summary[3][0] = 'dynamic full-column balance'
  summary[3][1] =
    "=SUMPRODUCT(IFERROR(1*(INDEX('Gift Cards'!A:Z,2,MATCH(\"Current Balance\",'Gift Cards'!A1:Z1,0)):INDEX('Gift Cards'!A:Z,ROWS('Gift Cards'!A:A),MATCH(\"Current Balance\",'Gift Cards'!A1:Z1,0))),0))"
  summary[5][0] = 'bounded filtered balance'
  summary[5][1] =
    "=-SUMPRODUCT(IFERROR(1*(INDEX('Gift Cards'!A2:Z2000,0,MATCH(\"Expired?\",'Gift Cards'!A1:Z1,0))=TRUE),0),IFERROR(1*(INDEX('Gift Cards'!A2:Z2000,0,MATCH(\"Enabled?\",'Gift Cards'!A1:Z1,0))=TRUE),0),IFERROR(1*INDEX('Gift Cards'!A2:Z2000,0,MATCH(\"Current Balance\",'Gift Cards'!A1:Z1,0)),0))"

  const giftCards = Array.from({ length: rows }, () => Array<SheetCell>(26).fill(null))
  giftCards[0][0] = 'Current Balance'
  giftCards[0][1] = 'Expired?'
  giftCards[0][2] = 'Enabled?'

  for (let row = 2; row <= rows; row += 1) {
    giftCards[row - 1][0] = row % 17
    giftCards[row - 1][1] = row % 2 === 0
    giftCards[row - 1][2] = row % 3 === 0
  }

  return {
    Summary: summary,
    'Gift Cards': giftCards,
  }
}

describe('dynamic full-column SUMPRODUCT formulas', () => {
  it('evaluates dynamic full-column and bounded INDEX ranges together within a bounded budget', () => {
    const rows = 30_000
    const workbook = WorkPaper.buildFromSheets(buildGiftCardSheets(rows), {
      evaluationTimeoutMs: 10_000,
      useWildcards: true,
      useRegularExpressions: false,
    })

    try {
      expect(numberValue(workbook, 'Summary!B4')).toBe(expectedBalance(rows))
      expect(numberValue(workbook, 'Summary!B6')).toBe(expectedFilteredBalance(rows))
    } finally {
      workbook.dispose()
    }
  }, 10_000)
})
