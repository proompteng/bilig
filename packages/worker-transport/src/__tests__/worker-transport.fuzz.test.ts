import { describe, expect, it } from 'vitest'
import fc from 'fast-check'
import type { Scheduler } from 'fast-check'
import { runScheduledProperty } from '@bilig/test-fuzz'
import { createWorkerEngineClient, createWorkerEngineHost, type MessagePortLike } from '../index.js'

type WorkerTransportAction = {
  method: 'add' | 'multiply'
  left: number
  right: number
}

describe('worker transport fuzz', () => {
  it('should route scheduled request and response traffic back to the matching client promises', async () => {
    await runScheduledProperty({
      suite: 'worker-transport/request-response-parity',
      arbitrary: fc.array(workerTransportActionArbitrary, { minLength: 4, maxLength: 16 }),
      predicate: async ({ scheduler, value: actions }) => {
        const deliveryPromises: Promise<void>[] = []
        const { hostPort, clientPort } = createScheduledPortPair(scheduler, deliveryPromises)
        const host = createWorkerEngineHost(
          {
            async ready() {
              return
            },
            async add(left: number, right: number) {
              await Promise.resolve()
              return left + right
            },
            async multiply(left: number, right: number) {
              await Promise.resolve()
              return left * right
            },
          },
          hostPort,
        )
        const client = createWorkerEngineClient({ port: clientPort })

        try {
          const readyPromise = client.ready()
          await expect(scheduler.waitFor(readyPromise)).resolves.toBeUndefined()
          const invokePromises = actions.map((action) => client.invoke(action.method, action.left, action.right))
          const actualResults = await scheduler.waitFor(Promise.all(invokePromises))
          expect(actualResults).toEqual(actions.map((action) => evaluateWorkerTransportAction(action)))
          await scheduler.waitIdle()
          await Promise.all(deliveryPromises)
        } finally {
          client.dispose()
          host.dispose()
          await Promise.allSettled(deliveryPromises)
        }
      },
    })
  })
})

// Helpers

const workerTransportActionArbitrary = fc
  .record({
    method: fc.constantFrom<'add' | 'multiply'>('add', 'multiply'),
    left: fc.integer({ min: -100, max: 100 }),
    right: fc.integer({ min: -100, max: 100 }),
  })
  .map((action) => ({
    method: action.method,
    left: action.left,
    right: action.right,
  }))

type PortMessage = Parameters<MessagePortLike['postMessage']>[0]
type PortListener = (event: MessageEvent<PortMessage>) => void

class ScheduledMessagePort implements MessagePortLike {
  peer: ScheduledMessagePort | null = null
  private readonly listeners = new Set<PortListener>()

  constructor(
    private readonly scheduleDelivery: (target: ScheduledMessagePort, message: PortMessage) => Promise<void>,
    private readonly deliveryPromises: Promise<void>[],
  ) {}

  postMessage(message: PortMessage): void {
    if (!this.peer) {
      throw new Error('Scheduled message port is not connected')
    }
    const delivery = this.scheduleDelivery(this.peer, structuredClone(message))
    this.deliveryPromises.push(delivery)
  }

  addEventListener(_type: 'message', listener: PortListener): void {
    this.listeners.add(listener)
  }

  removeEventListener(_type: 'message', listener: PortListener): void {
    this.listeners.delete(listener)
  }

  start(): void {}

  emit(message: PortMessage): void {
    const event = new MessageEvent<PortMessage>('message', { data: message })
    this.listeners.forEach((listener) => {
      listener(event)
    })
  }
}

function createScheduledPortPair(
  scheduler: Scheduler,
  deliveryPromises: Promise<void>[],
): {
  hostPort: MessagePortLike
  clientPort: MessagePortLike
} {
  const scheduleDelivery = scheduler.scheduleFunction(async (target: ScheduledMessagePort, message: PortMessage): Promise<void> => {
    target.emit(message)
  })
  const hostPort = new ScheduledMessagePort(scheduleDelivery, deliveryPromises)
  const clientPort = new ScheduledMessagePort(scheduleDelivery, deliveryPromises)
  hostPort.peer = clientPort
  clientPort.peer = hostPort
  return { hostPort, clientPort }
}

function evaluateWorkerTransportAction(action: WorkerTransportAction): number {
  return action.method === 'add' ? action.left + action.right : action.left * action.right
}
