import { useCallback, useEffect, useRef, useState } from 'react'
import { Archive, FolderOpen, Grid2X2 } from 'lucide-react'
import { importModelBackup, MODEL_LIMITS, type ModelDocument } from './model-document.js'
import { listModels, saveModel } from './model-repository.js'
import { createModel } from './model-templates.js'
import { ModelLibrary } from './ModelLibrary.js'
import { ModelEditor } from './ModelEditor.js'

type LibraryState =
  | { readonly kind: 'loading' }
  | { readonly kind: 'ready'; readonly models: readonly ModelDocument[] }
  | { readonly kind: 'error'; readonly message: string }

function activeModelId(): string | null {
  return new URLSearchParams(window.location.search).get('model')
}

export function ModelsWorkspace() {
  const [library, setLibrary] = useState<LibraryState>({ kind: 'loading' })
  const [activeId, setActiveId] = useState(activeModelId)
  const [isArchived, setArchived] = useState(false)
  const [isBusy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [retry, setRetry] = useState(0)
  const navigationGuard = useRef<(() => Promise<boolean>) | null>(null)
  const onNavigationGuard = useCallback((guard: (() => Promise<boolean>) | null) => {
    navigationGuard.current = guard
  }, [])

  useEffect(() => {
    let active = true
    void (async () => {
      try {
        const models = await listModels()
        if (active) setLibrary({ kind: 'ready', models })
      } catch (failure) {
        if (active) setLibrary({ kind: 'error', message: failure instanceof Error ? failure.message : 'Could not read saved models.' })
      }
    })()
    return () => {
      active = false
    }
  }, [retry])

  useEffect(() => {
    const onPopState = () => {
      const destinationId = activeModelId()
      void (async () => {
        if (navigationGuard.current && !(await navigationGuard.current())) {
          const currentUrl = new URL('/models', window.location.origin)
          if (activeId) currentUrl.searchParams.set('model', activeId)
          window.history.pushState(null, '', `${currentUrl.pathname}${currentUrl.search}`)
          return
        }
        setActiveId(destinationId)
        setArchived(false)
      })()
    }
    window.addEventListener('popstate', onPopState)
    return () => window.removeEventListener('popstate', onPopState)
  }, [activeId])

  const navigate = useCallback((id: string | null, archived = false) => {
    void (async () => {
      if (navigationGuard.current && !(await navigationGuard.current())) return
      const url = new URL('/models', window.location.origin)
      if (id) url.searchParams.set('model', id)
      window.history.pushState(null, '', `${url.pathname}${url.search}`)
      setActiveId(id)
      setArchived(archived)
      setError(null)
    })()
  }, [])

  const onSaved = useCallback((model: ModelDocument) => {
    setLibrary((current) =>
      current.kind === 'ready'
        ? {
            kind: 'ready',
            models: [model, ...current.models.filter((entry) => entry.id !== model.id)],
          }
        : current,
    )
  }, [])

  const run = (action: () => Promise<void>) => {
    setBusy(true)
    setError(null)
    void (async () => {
      try {
        await action()
      } catch (failure) {
        setError(failure instanceof Error ? failure.message : 'The model could not be saved.')
      } finally {
        setBusy(false)
      }
    })()
  }

  const create = (templateId: string) =>
    run(async () => {
      const model = await saveModel(createModel(templateId), 0)
      onSaved(model)
      navigate(model.id)
    })

  const selected = library.kind === 'ready' ? library.models.find((model) => model.id === activeId) : undefined
  const isEditing = selected?.status === 'active'
  useEffect(() => {
    document.title = selected ? `${selected.title} · Bilig` : 'Models · Bilig'
  }, [selected])

  return (
    <div className="models-workspace">
      <a className="model-skip-link" href="#model-main">
        Skip to model
      </a>
      <aside className="model-sidebar" aria-label="Workspace navigation">
        <a className="model-brand" href="/models">
          <span className="model-brand-mark">b</span>
          <span>bilig</span>
        </a>
        <span className="model-sidebar-caption">Personal workspace</span>
        <nav>
          <button
            aria-current={!isArchived ? 'page' : undefined}
            onClick={() => {
              navigate(null)
            }}
          >
            <FolderOpen size={17} /> Models
          </button>
          <button
            aria-current={isArchived ? 'page' : undefined}
            onClick={() => {
              navigate(null, true)
            }}
          >
            <Archive size={17} /> Archive
          </button>
        </nav>
        <div className="model-sidebar-bottom">
          <a href="/workbook">
            <Grid2X2 size={16} /> Spreadsheet editor
          </a>
          <p>Models are stored on this device.</p>
        </div>
      </aside>
      <div className="model-main-area">
        {error ? (
          <div role="alert" className="model-error">
            {error}
            <button className="model-text-button" onClick={() => setError(null)}>
              Dismiss
            </button>
          </div>
        ) : null}
        {library.kind === 'loading' ? (
          <main id="model-main" className="model-loading" role="status">
            Opening your workspace…
          </main>
        ) : library.kind === 'error' ? (
          <main id="model-main" className="model-loading">
            <h1>Could not open your workspace</h1>
            <p role="alert">{library.message}</p>
            <button
              className="model-button"
              onClick={() => {
                setLibrary({ kind: 'loading' })
                setRetry((value) => value + 1)
              }}
            >
              Retry
            </button>
          </main>
        ) : selected && isEditing ? (
          <ModelEditor
            key={selected.id}
            initial={selected}
            onSaved={onSaved}
            onBack={() => navigate(null)}
            onNavigationGuard={onNavigationGuard}
          />
        ) : activeId && !selected ? (
          <main id="model-main" className="model-loading">
            <h1>Model not found in this browser</h1>
            <p className="model-muted">Import a backup on this device, or return to your models.</p>
            <button className="model-button" onClick={() => navigate(null)}>
              Back to models
            </button>
          </main>
        ) : (
          <ModelLibrary
            models={library.models}
            isArchived={isArchived}
            isBusy={isBusy}
            onCreate={create}
            onOpen={navigate}
            onRestore={(model) =>
              run(async () => {
                onSaved(await saveModel({ ...model, status: 'active' }, model.revision))
              })
            }
            onImport={(file) =>
              run(async () => {
                if (file.size > MODEL_LIMITS.fileBytes) throw new Error('Model backups must be smaller than 1 MB.')
                const model = await saveModel(importModelBackup(await file.text()), 0)
                onSaved(model)
                navigate(model.id)
              })
            }
          />
        )}
      </div>
    </div>
  )
}
