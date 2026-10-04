/// <reference lib="webworker" />
import { calculateModel, parseModelSensitivityRequest, type ModelCalculation, type ModelSensitivityRequest } from './model-calculation.js'
import { parseModelDocument } from './model-document.js'
import type { ModelDocument } from './model-document.js'

declare const self: DedicatedWorkerGlobalScope

export type ModelWorkerResponse =
  | { readonly kind: 'ready'; readonly calculation: ModelCalculation }
  | { readonly kind: 'error'; readonly message: string }

export interface ModelWorkerRequest {
  readonly model: ModelDocument
  readonly sensitivity: ModelSensitivityRequest | null
}

self.addEventListener('message', (event: MessageEvent<unknown>) => {
  let response: ModelWorkerResponse
  try {
    const request = event.data
    if (typeof request !== 'object' || !request || !('model' in request) || !('sensitivity' in request)) {
      throw new Error('Invalid model calculation request.')
    }
    response = {
      kind: 'ready',
      calculation: calculateModel(parseModelDocument(request.model), parseModelSensitivityRequest(request.sensitivity)),
    }
  } catch (error) {
    response = { kind: 'error', message: error instanceof Error ? error.message : 'The model could not be calculated.' }
  }
  self.postMessage(response, [])
})
