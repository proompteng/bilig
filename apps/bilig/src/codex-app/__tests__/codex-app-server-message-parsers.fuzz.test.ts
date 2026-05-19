import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { fuzzJsonValueArbitrary, runProperty } from '@bilig/test-fuzz'
import { parseJsonRpcResponse, parseServerNotification, parseServerRequest } from '../codex-app-server-message-parsers.js'

describe('Codex app-server message parser fuzz', () => {
  it('should parse generated JSON-RPC responses without losing result or error identity', async () => {
    await runProperty({
      suite: 'bilig/codex-app/message-parsers/json-rpc-response',
      arbitrary: fc.oneof(jsonRpcResultResponseArbitrary, jsonRpcErrorResponseArbitrary),
      predicate: async (response) => {
        expect(parseJsonRpcResponse(response)).toEqual(response)
      },
      parameters: { numRuns: 100 },
    })
  })

  it('should parse generated server requests and reject non-json params', async () => {
    await runProperty({
      suite: 'bilig/codex-app/message-parsers/server-request',
      arbitrary: fc.record({
        request: serverRequestArbitrary,
        corruptParams: fc.oneof(
          fc.constant(() => null),
          fc.constant(Symbol('bad')),
        ),
      }),
      predicate: async ({ request, corruptParams }) => {
        expect(parseServerRequest(request)).toEqual(request)
        expect(parseServerRequest({ ...request, params: corruptParams })).toBeNull()
      },
      parameters: { numRuns: 100 },
    })
  })

  it('should parse generated notifications and reject unknown methods', async () => {
    await runProperty({
      suite: 'bilig/codex-app/message-parsers/notifications',
      arbitrary: notificationArbitrary,
      predicate: async (notification) => {
        expect(parseServerNotification(notification)).toEqual(notification)
        expect(parseServerNotification({ ...notification, method: 'unknown/event' })).toBeNull()
      },
      parameters: { numRuns: 100 },
    })
  })
})

// Helpers

const requestIdArbitrary = fc.oneof(fc.uuid(), fc.integer({ min: 0, max: 10_000 }))

const jsonRpcResultResponseArbitrary = fc.record({
  id: requestIdArbitrary,
  result: fuzzJsonValueArbitrary,
})

const jsonRpcErrorResponseArbitrary = fc.record({
  id: requestIdArbitrary,
  error: fc.record({
    code: fc.integer({ min: -32_768, max: 32_767 }),
    message: fc.string({ maxLength: 80 }),
    data: fc.option(fuzzJsonValueArbitrary, { nil: undefined }),
  }),
})

const toolCallRequestArbitrary = fc
  .record({
    id: requestIdArbitrary,
    threadId: fc.uuid(),
    turnId: fc.uuid(),
    callId: fc.uuid(),
    tool: fc.constantFrom('read_range', 'write_range', 'verify_invariants'),
    argumentsValue: fuzzJsonValueArbitrary,
    namespace: fc.option(fc.constantFrom('workbook', 'codex'), { nil: null }),
  })
  .map(({ id, threadId, turnId, callId, tool, argumentsValue, namespace }) => ({
    method: 'item/tool/call' as const,
    id,
    params: {
      threadId,
      turnId,
      callId,
      tool,
      arguments: argumentsValue,
      namespace,
    },
  }))

const serverRequestArbitrary = fc.oneof(
  toolCallRequestArbitrary,
  fc
    .record({
      id: requestIdArbitrary,
      method: fc.constantFrom('thread/start', 'thread/resume', 'turn/start'),
      params: fc.option(fuzzJsonValueArbitrary, { nil: undefined }),
    })
    .map(({ id, method, params }) => {
      const request: { method: string; id: string | number; params?: unknown } = {
        method,
        id,
      }
      if (params !== undefined) {
        request.params = params
      }
      return request
    }),
)

const notificationArbitrary = fc.oneof(
  fc
    .record({
      threadId: fc.uuid(),
      turnId: fc.uuid(),
      itemId: fc.uuid(),
      delta: fc.string({ maxLength: 80 }),
      method: fc.constantFrom(
        'item/agentMessage/delta',
        'item/plan/delta',
        'item/reasoning/delta',
        'item/reasoning/textDelta',
        'item/reasoning/summaryTextDelta',
        'item/commandExecution/outputDelta',
      ),
    })
    .map(({ method, threadId, turnId, itemId, delta }) => ({
      method,
      params: {
        threadId,
        turnId,
        itemId,
        delta,
      },
    })),
  fc
    .record({
      message: fc.option(fc.string({ maxLength: 80 }), { nil: undefined }),
      code: fc.option(fc.string({ maxLength: 24 }), { nil: undefined }),
    })
    .map((params) => ({
      method: 'error' as const,
      params,
    })),
)
