import { useCallback, useRef, useState } from 'react'
import { deferInteractionPersistence } from './interaction-idle-scheduler.js'

interface DeferredEditCommitTask {
  ready: boolean
  runNow(): Promise<void>
}

export function useDeferredEditCommits() {
  const queueRef = useRef<DeferredEditCommitTask[]>([])
  const [isEditCommitPending, setIsEditCommitPending] = useState(false)

  const drain = useCallback(async (readyOnly: boolean): Promise<void> => {
    const drainTasks = async (): Promise<void> => {
      const nextTask = queueRef.current[0]
      if (!nextTask || (readyOnly && !nextTask.ready)) {
        return
      }
      nextTask.ready = true
      await nextTask.runNow()
      if (queueRef.current[0] === nextTask) {
        queueRef.current.shift()
        setIsEditCommitPending(queueRef.current.length > 0)
      }
      await drainTasks()
    }
    await drainTasks()
  }, [])

  const flushPendingEditCommit = useCallback(async (): Promise<void> => {
    await drain(false)
  }, [drain])

  const enqueueDeferredEditCommit = useCallback(
    (run: () => Promise<void>): void => {
      let taskPromise: Promise<void> | null = null
      const task: DeferredEditCommitTask = {
        ready: false,
        runNow() {
          task.ready = true
          taskPromise ??= run()
          return taskPromise
        },
      }
      queueRef.current.push(task)
      setIsEditCommitPending(true)
      void (async () => {
        await deferInteractionPersistence()
        task.ready = true
        await drain(true)
      })()
    },
    [drain],
  )

  return { enqueueDeferredEditCommit, flushPendingEditCommit, isEditCommitPending }
}
