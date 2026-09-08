import { useEffect, useState } from 'react'
import type { ModelDocument } from './model-document.js'
import type { ModelWorkerResponse } from './model.worker.js'

type CalculationState = { readonly kind: 'calculating' } | ModelWorkerResponse

export function useModelCalculation(model: ModelDocument): CalculationState {
  const [result, setResult] = useState<{ readonly model: ModelDocument; readonly state: CalculationState } | null>(null)
  useEffect(() => {
    let worker: Worker | null = null
    let timeout: ReturnType<typeof setTimeout> | undefined
    const debounce = setTimeout(() => {
      try {
        worker = new Worker(new URL('./model.worker.ts', import.meta.url), { type: 'module' })
        timeout = setTimeout(() => {
          worker?.terminate()
          setResult({ model, state: { kind: 'error', message: 'Calculation took too long. Simplify the formula and try again.' } })
        }, 10_000)
        worker.addEventListener('message', (event: MessageEvent<ModelWorkerResponse>) => {
          clearTimeout(timeout)
          setResult({ model, state: event.data })
          worker?.terminate()
        })
        worker.addEventListener('error', () => {
          clearTimeout(timeout)
          setResult({
            model,
            state: { kind: 'error', message: 'The calculation worker could not start. Reload to retry.' },
          })
          worker?.terminate()
        })
        worker.postMessage(model, [])
      } catch (error) {
        setResult({ model, state: { kind: 'error', message: error instanceof Error ? error.message : 'Calculation could not start.' } })
      }
    }, 100)
    return () => {
      clearTimeout(debounce)
      clearTimeout(timeout)
      worker?.terminate()
    }
  }, [model])
  return result?.model === model ? result.state : { kind: 'calculating' }
}
