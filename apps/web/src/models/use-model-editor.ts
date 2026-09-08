import { useCallback, useEffect, useRef, useState } from 'react'
import type { ModelDocument } from './model-document.js'
import { saveModel } from './model-repository.js'

type SaveState = { readonly kind: 'saved' | 'saving' } | { readonly kind: 'error'; readonly message: string }

export function useModelEditor(initial: ModelDocument, onSaved: (model: ModelDocument) => void) {
  const [model, setModel] = useState(initial)
  const [saveState, setSaveState] = useState<SaveState>({ kind: 'saved' })
  const current = useRef(initial)
  const revision = useRef(initial.revision)
  const queue = useRef(Promise.resolve())
  const failure = useRef(false)
  const pending = useRef(0)
  const mounted = useRef(true)
  useEffect(() => {
    mounted.current = true
    const beforeUnload = (event: BeforeUnloadEvent) => {
      if (pending.current > 0 || failure.current) event.preventDefault()
    }
    window.addEventListener('beforeunload', beforeUnload)
    return () => {
      mounted.current = false
      window.removeEventListener('beforeunload', beforeUnload)
    }
  }, [])

  const change = useCallback(
    (update: (model: ModelDocument) => ModelDocument) => {
      const next = update(current.current)
      current.current = next
      setModel(next)
      if (failure.current) return
      pending.current += 1
      setSaveState({ kind: 'saving' })
      const previous = queue.current
      queue.current = (async () => {
        await previous
        try {
          if (failure.current) return
          const saved = await saveModel(next, revision.current)
          revision.current = saved.revision
          onSaved(saved)
        } catch (error) {
          failure.current = true
          if (mounted.current)
            setSaveState({ kind: 'error', message: error instanceof Error ? error.message : 'Could not save this model.' })
        } finally {
          pending.current -= 1
          if (mounted.current && pending.current === 0 && !failure.current) setSaveState({ kind: 'saved' })
        }
      })()
    },
    [onSaved],
  )

  const waitForSave = useCallback(async () => {
    await queue.current
    return !failure.current
  }, [])

  const reloadSavedVersion = useCallback(() => {
    void (async () => {
      await queue.current
      failure.current = false
      window.location.reload()
    })()
  }, [])

  return { model, change, saveState, waitForSave, reloadSavedVersion }
}
