import { describe, expect, it, vi } from 'vitest'
import { formatAddress } from '@bilig/formula'
import { ValueTag } from '@bilig/protocol'
import {
  autofitProjectedColumnWidth,
  normalizeProjectedColumnWidth,
  normalizeProjectedRowHeight,
  WorkerRuntimeProjectionCommands,
} from '../worker-runtime-projection-commands.js'

describe('worker runtime projection commands', () => {
  it('normalizes row heights to positive rounded values', () => {
    expect(normalizeProjectedRowHeight(null)).toBeNull()
    expect(normalizeProjectedRowHeight(41.6)).toBe(42)
    expect(normalizeProjectedRowHeight(0.2)).toBe(1)
  })

  it('normalizes column widths into the allowed range', () => {
    expect(normalizeProjectedColumnWidth(null, 24, 480)).toBeNull()
    expect(normalizeProjectedColumnWidth(12.2, 24, 480)).toBe(24)
    expect(normalizeProjectedColumnWidth(128.6, 24, 480)).toBe(129)
    expect(normalizeProjectedColumnWidth(999, 24, 480)).toBe(480)
  })

  it('computes autofit widths from formatted display values and column labels', () => {
    const width = autofitProjectedColumnWidth({
      columnIndex: 2,
      charWidth: 8,
      padding: 12,
      sheet: {
        grid: {
          forEachCellEntry(listener) {
            listener(0, 0, 2)
            listener(1, 1, 2)
            listener(2, 1, 3)
          },
        },
      },
      getCellDisplayValue(row, col) {
        if (col === 2 && row === 0) {
          return 'short'
        }
        if (col === 2 && row === 1) {
          return 'the longest value'
        }
        return 'ignored'
      },
    })

    expect(width).toBe('the longest value'.length * 8 + 12)
  })

  it('marks the projection as diverged before mutating the engine', async () => {
    const calls: string[] = []
    const engine = {
      setCellValue(sheetName: string, address: string, value: unknown) {
        calls.push(`set:${sheetName}:${address}:${String(value)}`)
      },
      getCell(sheetName: string, address: string) {
        return { sheetName, address, value: { tag: 0 } }
      },
    }
    const commands = new WorkerRuntimeProjectionCommands({
      markProjectionDivergedFromLocalStore() {
        calls.push('mark')
      },
      async getProjectionEngine() {
        calls.push('get')
        return engine
      },
      getCell(sheetName, address) {
        calls.push(`read:${sheetName}:${address}`)
        return engine.getCell(sheetName, address)
      },
      minColumnWidth: 24,
      maxColumnWidth: 480,
      autofitCharWidth: 8,
      autofitPadding: 12,
      formatCellDisplayValue(value) {
        return value.tag === ValueTag.String ? value.value : ''
      },
    })

    await commands.setCellValue('Sheet1', 'A1', 7)

    expect(calls).toEqual(['mark', 'get', 'set:Sheet1:A1:7', 'read:Sheet1:A1'])
  })

  it('delegates autofit through normalized column width updates', async () => {
    const updateColumnMetadata = vi.fn()
    const commands = new WorkerRuntimeProjectionCommands({
      markProjectionDivergedFromLocalStore: vi.fn(),
      async getProjectionEngine() {
        return {
          workbook: {
            getSheet() {
              return {
                grid: {
                  forEachCellEntry(listener: (cellIndex: number, row: number, col: number) => void) {
                    listener(0, 0, 1)
                    listener(1, 1, 1)
                  },
                },
              }
            },
          },
          getCell(_sheetName: string, address: string) {
            return {
              address,
              format: undefined,
              value:
                address === formatAddress(1, 1)
                  ? { tag: ValueTag.String, value: 'wider value', stringId: 1 }
                  : { tag: ValueTag.String, value: 'x', stringId: 2 },
            }
          },
          updateColumnMetadata,
        }
      },
      getCell(sheetName, address) {
        return { sheetName, address, value: { tag: 0 } }
      },
      minColumnWidth: 24,
      maxColumnWidth: 480,
      autofitCharWidth: 8,
      autofitPadding: 12,
      formatCellDisplayValue(value) {
        return value.tag === ValueTag.String ? value.value : ''
      },
    })

    const width = await commands.autofitColumn('Sheet1', 1)

    expect(width).toBe('wider value'.length * 8 + 12)
    expect(updateColumnMetadata).toHaveBeenCalledWith('Sheet1', 1, 1, width, null)
  })
})
