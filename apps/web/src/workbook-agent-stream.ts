import { useCallback, useRef, type Dispatch, type SetStateAction } from 'react'
import { WorkbookAgentStreamEventSchema, decodeUnknownSync, type WorkbookAgentThreadSnapshot } from '@bilig/contracts'
import type { createWorkbookPerfSession } from './perf/workbook-perf.js'
import type { createWorkbookAgentClient } from './workbook-agent-client.js'
import { readMessageEventData } from './workbook-agent-pane-helpers.js'
import { updateSnapshotFromTextDelta, updateSnapshotFromToolOutputDelta } from './workbook-agent-stream-state.js'

export interface WorkbookAgentLiveSession {
  threadId: string
}

interface MutableRef<T> {
  current: T
}

interface UseWorkbookAgentStreamInput {
  readonly client: ReturnType<typeof createWorkbookAgentClient>
  readonly sessionRef: MutableRef<WorkbookAgentLiveSession | null>
  readonly perfSession: ReturnType<typeof createWorkbookPerfSession>
  readonly persistSessionSnapshot: (nextSnapshot: WorkbookAgentThreadSnapshot) => void
  readonly setError: Dispatch<SetStateAction<string | null>>
  readonly setIsLoading: Dispatch<SetStateAction<boolean>>
  readonly setSnapshot: Dispatch<SetStateAction<WorkbookAgentThreadSnapshot | null>>
}

export function decodeWorkbookAgentStreamEventPayload(payloadText: string) {
  let payload: unknown
  try {
    payload = JSON.parse(payloadText) as unknown
  } catch {
    throw new Error('Assistant stream returned malformed event data.')
  }
  try {
    return decodeUnknownSync(WorkbookAgentStreamEventSchema, payload)
  } catch {
    throw new Error('Assistant stream returned invalid event data.')
  }
}

export function useWorkbookAgentStream(input: UseWorkbookAgentStreamInput) {
  const { client, perfSession, persistSessionSnapshot, sessionRef, setError, setIsLoading, setSnapshot } = input
  const eventSourceRef = useRef<EventSource | null>(null)
  const recoveringStreamRef = useRef(false)
  const streamGenerationRef = useRef(0)

  const closeStream = useCallback(() => {
    streamGenerationRef.current += 1
    eventSourceRef.current?.close()
    eventSourceRef.current = null
  }, [])

  const connectStream = useCallback(
    (threadId: string) => {
      closeStream()
      recoveringStreamRef.current = false
      const streamGeneration = streamGenerationRef.current
      const source = new EventSource(client.threadEventsUrl(threadId))
      source.addEventListener('message', (message) => {
        if (eventSourceRef.current !== source || streamGenerationRef.current !== streamGeneration) {
          return
        }
        try {
          const payloadText = readMessageEventData(message)
          if (payloadText === null) {
            return
          }
          const event = decodeWorkbookAgentStreamEventPayload(payloadText)
          if (event.type === 'snapshot') {
            recoveringStreamRef.current = false
            persistSessionSnapshot(event.snapshot)
            setError(null)
            return
          }
          setSnapshot((current: WorkbookAgentThreadSnapshot | null) =>
            event.type === 'entryTextDelta'
              ? updateSnapshotFromTextDelta(current, event)
              : updateSnapshotFromToolOutputDelta(current, event),
          )
          if (event.type === 'entryTextDelta' && event.entryKind === 'assistant') {
            perfSession.markFirstAssistantDeltaVisible?.()
          }
        } catch (nextError) {
          setError(nextError instanceof Error ? nextError.message : String(nextError))
        }
      })
      source.addEventListener('error', () => {
        if (eventSourceRef.current !== source || streamGenerationRef.current !== streamGeneration) {
          return
        }
        source.close()
        eventSourceRef.current = null
        if (recoveringStreamRef.current) {
          return
        }
        const storedSession = sessionRef.current
        if (!storedSession) {
          setError('Assistant stream disconnected.')
          return
        }
        recoveringStreamRef.current = true
        setError(null)
        void (async () => {
          try {
            setIsLoading(true)
            const nextSnapshot = await client.loadThreadSnapshot(storedSession.threadId)
            if (eventSourceRef.current !== null || streamGenerationRef.current !== streamGeneration) {
              return
            }
            if (sessionRef.current?.threadId !== storedSession.threadId) {
              return
            }
            persistSessionSnapshot(nextSnapshot)
            connectStream(nextSnapshot.threadId)
            setIsLoading(false)
          } catch (nextError) {
            if (streamGenerationRef.current === streamGeneration && sessionRef.current?.threadId === storedSession.threadId) {
              recoveringStreamRef.current = false
              setError(nextError instanceof Error ? nextError.message : String(nextError))
            }
          } finally {
            if (streamGenerationRef.current === streamGeneration && sessionRef.current?.threadId === storedSession.threadId) {
              setIsLoading(false)
            }
          }
        })()
      })
      eventSourceRef.current = source
    },
    [client, closeStream, perfSession, persistSessionSnapshot, sessionRef, setError, setIsLoading, setSnapshot],
  )

  const resetRecoveringStream = useCallback(() => {
    recoveringStreamRef.current = false
  }, [])

  return {
    closeStream,
    connectStream,
    resetRecoveringStream,
  }
}
