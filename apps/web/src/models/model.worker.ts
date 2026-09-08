/// <reference lib="webworker" />
import { calculateModel, type ModelCalculation } from './model-calculation.js'
import { parseModelDocument } from './model-document.js'

declare const self: DedicatedWorkerGlobalScope

export type ModelWorkerResponse =
  | { readonly kind: 'ready'; readonly calculation: ModelCalculation }
  | { readonly kind: 'error'; readonly message: string }

self.addEventListener('message', (event: MessageEvent<unknown>) => {
  let response: ModelWorkerResponse
  try {
    response = { kind: 'ready', calculation: calculateModel(parseModelDocument(event.data)) }
  } catch (error) {
    response = { kind: 'error', message: error instanceof Error ? error.message : 'The model could not be calculated.' }
  }
  self.postMessage(response, [])
})
