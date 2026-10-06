// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { expect, it, vi } from 'vitest'
import { useWorkbookSelectionSummary } from '../use-workbook-selection-summary.js'

const response = (sum: number) => ({ nonEmptyCount: 1, numericCount: 1, sum, min: sum, max: sum })

it('ignores superseded ranges and preserves the current summary while refreshing', async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
  const pending: Array<(value: unknown) => void> = []
  const runtime = {
    invoke: vi.fn(
      () =>
        new Promise<unknown>((resolve) => {
          pending.push(resolve)
        }),
    ),
  }
  const onError = vi.fn()
  function Harness({ address, revision, kind = 'range' }: { address: string; revision: number; kind?: 'cell' | 'range' }) {
    const result = useWorkbookSelectionSummary({
      runtime,
      onError,
      runtimeState: revision,
      selection: { sheetName: 'Sheet1', address, kind, range: { startAddress: address, endAddress: 'B3' } },
    })
    return <output>{result?.sum ?? 'pending'}</output>
  }
  const host = document.createElement('div')
  document.body.appendChild(host)
  const root = createRoot(host)
  try {
    await act(async () => {
      root.render(<Harness address="A1" revision={1} />)
    })
    await act(async () => {
      root.render(<Harness address="B1" revision={1} />)
    })
    expect(host.textContent).toBe('pending')
    await act(async () => {
      pending[1]?.(response(20))
    })
    expect(host.textContent).toBe('20')
    await act(async () => {
      pending[0]?.(response(10))
    })
    expect(host.textContent).toBe('20')
    await act(async () => {
      root.render(<Harness address="B1" revision={2} />)
    })
    expect(host.textContent).toBe('20')
    await act(async () => {
      pending[2]?.(response(30))
    })
    expect(host.textContent).toBe('30')
    await act(async () => {
      root.render(<Harness address="C1" revision={3} kind="cell" />)
    })
    await act(async () => {
      root.render(<Harness address="B1" revision={3} />)
    })
    expect(host.textContent).toBe('pending')
    await act(async () => {
      pending[3]?.(response(40))
    })
    expect(host.textContent).toBe('40')
    expect(runtime.invoke).toHaveBeenCalledTimes(4)
    expect(onError).not.toHaveBeenCalled()
  } finally {
    await act(async () => {
      root.unmount()
    })
    host.remove()
    vi.unstubAllGlobals()
  }
})
