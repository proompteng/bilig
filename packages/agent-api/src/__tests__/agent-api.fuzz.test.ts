import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import {
  cloneJsonValue,
  fuzzCellRangeRefArbitrary,
  fuzzLiteralInputArbitrary,
  fuzzWorkbookSnapshotArbitrary,
  runProperty,
} from '@bilig/test-fuzz'
import type { AgentFrame, AgentRequest, AgentResponse } from '../index.js'
import {
  decodeAgentFrame,
  decodeStdioMessages,
  encodeAgentFrame,
  encodeStdioMessage,
  normalizeWorkbookImportContentType,
} from '../index.js'

describe('agent api frame fuzz', () => {
  it('should roundtrip generated frames through binary and stdio codecs', async () => {
    await runProperty({
      suite: 'agent-api/frame-codec/generated-roundtrip',
      arbitrary: fc.array(agentFrameArbitrary, { minLength: 1, maxLength: 5 }),
      predicate: async (frames) => {
        for (const frame of frames) {
          expect(decodeAgentFrame(encodeAgentFrame(frame))).toEqual(frame)
        }

        const buffer = concatBytes(frames.map((frame) => encodeStdioMessage(frame)))
        const splitAt = Math.floor(buffer.byteLength / 2)
        const first = decodeStdioMessages(buffer.subarray(0, splitAt))
        const second = decodeStdioMessages(concatBytes([first.remainder, buffer.subarray(splitAt)]))

        expect([...first.frames, ...second.frames]).toEqual(frames)
        expect(second.remainder.byteLength).toBe(0)
      },
      parameters: { numRuns: 100 },
    })
  })

  it('should reject generated frame corruptions instead of accepting unsafe payloads', async () => {
    await runProperty({
      suite: 'agent-api/frame-codec/reject-corruption',
      arbitrary: agentFrameArbitrary,
      predicate: async (frame) => {
        const encoded = encodeAgentFrame(frame)
        const wrongMagic = new Uint8Array(encoded)
        wrongMagic[0] = wrongMagic[0] === 0 ? 1 : 0
        expect(() => decodeAgentFrame(wrongMagic)).toThrow(/magic mismatch/u)

        const badPayload = new Uint8Array(encoded)
        badPayload[badPayload.byteLength - 1] = 0xff
        expect(() => decodeAgentFrame(badPayload)).toThrow()
      },
      parameters: { numRuns: 80 },
    })
  })

  it('should normalize workbook import content types with generated metadata casing and parameters', async () => {
    await runProperty({
      suite: 'agent-api/import-content-type/metadata-normalization',
      arbitrary: fc.record({
        base: fc.constantFrom(
          'text/csv',
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          'application/vnd.ms-excel.sheet.macroenabled.12',
          'application/vnd.ms-excel.sheet.binary.macroenabled.12',
          'application/vnd.ms-excel',
        ),
        prefixSpaces: fc.array(fc.constant(' '), { maxLength: 3 }).map((chars) => chars.join('')),
        suffixSpaces: fc.array(fc.constant(' '), { maxLength: 3 }).map((chars) => chars.join('')),
        withParameter: fc.boolean(),
      }),
      predicate: async ({ base, prefixSpaces, suffixSpaces, withParameter }) => {
        const cased = base
          .split('')
          .map((char, index) => (index % 2 === 0 ? char.toUpperCase() : char.toLowerCase()))
          .join('')
        const raw = `${prefixSpaces}${cased}${withParameter ? '; charset=utf-8' : ''}${suffixSpaces}`
        expect(normalizeWorkbookImportContentType(raw)).toBe(base)
      },
      parameters: { numRuns: 80 },
    })
  })
})

// Helpers

const requestArbitrary: fc.Arbitrary<AgentRequest> = fc.oneof(
  fc
    .record({
      id: fc.uuid(),
      documentId: fc.uuid(),
      replicaId: fc.uuid(),
    })
    .map(({ id, documentId, replicaId }) => ({
      kind: 'openWorkbookSession' as const,
      id,
      documentId,
      replicaId,
    })),
  fc
    .record({
      id: fc.uuid(),
      sessionId: fc.uuid(),
      range: fuzzCellRangeRefArbitrary,
      values: fc.array(fc.array(fuzzLiteralInputArbitrary, { minLength: 1, maxLength: 4 }), { minLength: 1, maxLength: 4 }),
    })
    .map(({ id, sessionId, range, values }) => ({
      kind: 'writeRange' as const,
      id,
      sessionId,
      range,
      values,
    })),
  fc
    .record({
      id: fc.uuid(),
      sessionId: fc.uuid(),
      snapshot: fuzzWorkbookSnapshotArbitrary,
    })
    .map(({ id, sessionId, snapshot }) => ({
      kind: 'importSnapshot' as const,
      id,
      sessionId,
      snapshot: cloneJsonValue(snapshot),
    })),
)

const responseArbitrary: fc.Arbitrary<AgentResponse> = fc.oneof(
  fc
    .record({
      id: fc.uuid(),
      sessionId: fc.option(fc.uuid(), { nil: undefined }),
      value: fc.option(fc.dictionary(fc.string({ maxLength: 8 }), fc.string({ maxLength: 16 })), { nil: undefined }),
    })
    .map(({ id, sessionId, value }) => {
      const response: Extract<AgentResponse, { kind: 'ok' }> = {
        kind: 'ok',
        id,
      }
      if (sessionId !== undefined) {
        response.sessionId = sessionId
      }
      if (value !== undefined) {
        response.value = value
      }
      return response
    }),
  fc
    .record({
      id: fc.uuid(),
      code: fc.constantFrom('BAD_REQUEST', 'NOT_FOUND', 'CONFLICT'),
      message: fc.string({ maxLength: 80 }),
      retryable: fc.boolean(),
    })
    .map((error) => ({
      kind: 'error' as const,
      id: error.id,
      code: error.code,
      message: error.message,
      retryable: error.retryable,
    })),
)

const agentFrameArbitrary: fc.Arbitrary<AgentFrame> = fc.oneof(
  requestArbitrary.map((request) => ({ kind: 'request' as const, request })),
  responseArbitrary.map((response) => ({ kind: 'response' as const, response })),
  fc
    .record({
      subscriptionId: fc.uuid(),
      range: fuzzCellRangeRefArbitrary,
      changedAddresses: fc.array(fc.constantFrom('A1', 'B2', 'C3'), { maxLength: 4 }),
    })
    .map(({ subscriptionId, range, changedAddresses }) => ({
      kind: 'event' as const,
      event: {
        kind: 'rangeChanged' as const,
        subscriptionId,
        range,
        changedAddresses,
      },
    })),
)

function concatBytes(chunks: readonly Uint8Array[]): Uint8Array {
  const output = new Uint8Array(chunks.reduce((total, chunk) => total + chunk.byteLength, 0))
  let offset = 0
  for (const chunk of chunks) {
    output.set(chunk, offset)
    offset += chunk.byteLength
  }
  return output
}
