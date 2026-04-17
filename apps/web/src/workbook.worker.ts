/// <reference lib="webworker" />

import { createWorkerEngineHost, type MessagePortLike } from '@bilig/worker-transport'
import { WorkbookWorkerRuntime } from './worker-runtime.js'

declare const self: DedicatedWorkerGlobalScope

const scope = self
const runtime = new WorkbookWorkerRuntime()
type PortListener = Parameters<NonNullable<MessagePortLike['addEventListener']>>[1]
const listenerMap = new Map<PortListener, EventListener>()

createWorkerEngineHost(runtime, {
  postMessage(message: unknown) {
    scope.postMessage(message, [])
  },
  addEventListener(_type: 'message', listener: PortListener) {
    const wrapped: EventListener = (event) => {
      if (event instanceof MessageEvent) {
        listener(event)
      }
    }
    listenerMap.set(listener, wrapped)
    scope.addEventListener('message', wrapped)
  },
  removeEventListener(_type: 'message', listener: PortListener) {
    const wrapped = listenerMap.get(listener)
    if (!wrapped) {
      return
    }
    listenerMap.delete(listener)
    scope.removeEventListener('message', wrapped)
  },
})
