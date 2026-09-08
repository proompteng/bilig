import { useRef, useState } from 'react'
import { ArrowRight, FileUp, Plus, Search } from 'lucide-react'
import type { ModelDocument } from './model-document.js'
import { MODEL_TEMPLATES } from './model-templates.js'

export function ModelLibrary(props: {
  readonly models: readonly ModelDocument[]
  readonly isArchived: boolean
  readonly isBusy: boolean
  readonly onCreate: (templateId: string) => void
  readonly onOpen: (modelId: string) => void
  readonly onImport: (file: File) => void
  readonly onRestore: (model: ModelDocument) => void
}) {
  const [search, setSearch] = useState('')
  const input = useRef<HTMLInputElement>(null)
  const models = props.models.filter(
    (model) =>
      model.status === (props.isArchived ? 'archived' : 'active') &&
      `${model.title} ${model.description}`.toLocaleLowerCase().includes(search.toLocaleLowerCase()),
  )
  return (
    <main className="model-library" id="model-main">
      <div className="model-page-heading">
        <div>
          <p className="model-eyebrow">Workspace</p>
          <h1>{props.isArchived ? 'Archive' : 'Your models'}</h1>
          <p className="model-muted">
            {props.isArchived ? 'Archived models stay saved. Restore one to keep working.' : 'A place to work through the numbers.'}
          </p>
        </div>
        {!props.isArchived ? (
          <div className="model-actions">
            <input
              hidden
              ref={input}
              type="file"
              accept=".json,application/json"
              aria-label="Import model backup"
              className="model-file-input"
              onChange={(event) => {
                const file = event.target.files?.[0]
                event.target.value = ''
                if (file) props.onImport(file)
              }}
            />
            <button className="model-button" disabled={props.isBusy} onClick={() => input.current?.click()}>
              <FileUp size={16} /> Import backup
            </button>
            <button className="model-button model-primary" disabled={props.isBusy} onClick={() => props.onCreate('custom')}>
              <Plus size={16} /> New model
            </button>
          </div>
        ) : null}
      </div>
      {!props.isArchived ? (
        <section className="model-starters" aria-labelledby="model-starters-title">
          <div className="model-section-heading">
            <h2 id="model-starters-title">Start with a question</h2>
            <span className="model-muted">Editable models with working formulas</span>
          </div>
          <div className="model-template-list">
            {MODEL_TEMPLATES.filter((template) => template.id !== 'custom').map((template, index) => (
              <button key={template.id} className="model-template" disabled={props.isBusy} onClick={() => props.onCreate(template.id)}>
                <span className="model-template-index">0{index + 1}</span>
                <span>
                  <strong>{template.question}</strong>
                  <span className="model-muted">{template.description}</span>
                  <span className="model-template-title">
                    {template.title} <ArrowRight size={14} />
                  </span>
                </span>
              </button>
            ))}
          </div>
        </section>
      ) : null}
      <section aria-labelledby="model-list-title" className="model-saved-list">
        <div className="model-section-heading">
          <h2 id="model-list-title">
            {props.isArchived ? 'Archived models' : 'Saved models'} <span className="model-count">{models.length}</span>
          </h2>
          <label className="model-search">
            <Search size={16} />
            <input
              type="search"
              placeholder="Search models"
              aria-label="Search models"
              value={search}
              onChange={(event) => setSearch(event.target.value)}
            />
          </label>
        </div>
        {models.length === 0 ? (
          <div className="model-empty">
            <div className="model-empty-mark" aria-hidden="true">
              ƒ
            </div>
            <h3>{search ? 'No matching models' : props.isArchived ? 'Nothing archived' : 'Your first model starts here'}</h3>
            <p className="model-muted">
              {search
                ? 'Try another name or clear your search.'
                : props.isArchived
                  ? 'Models you archive will appear here.'
                  : 'Choose a question above, or build a model with your own assumptions.'}
            </p>
          </div>
        ) : (
          <div className="model-records">
            {models.map((model) => (
              <div className="model-record" key={model.id}>
                <button className="model-record-open" onClick={() => props.onOpen(model.id)} disabled={props.isArchived}>
                  <span className="model-record-icon" aria-hidden="true">
                    ƒ
                  </span>
                  <span>
                    <strong>{model.title}</strong>
                    <span className="model-muted">{model.description}</span>
                  </span>
                </button>
                <span className="model-record-meta">
                  {model.scenarios.length} {model.scenarios.length === 1 ? 'scenario' : 'scenarios'}
                  <time dateTime={model.updatedAt}>
                    {new Date(model.updatedAt).toLocaleDateString(undefined, { month: 'short', day: 'numeric' })}
                  </time>
                </span>
                {props.isArchived ? (
                  <button className="model-button" disabled={props.isBusy} onClick={() => props.onRestore(model)}>
                    Restore
                  </button>
                ) : (
                  <ArrowRight aria-hidden="true" size={16} />
                )}
              </div>
            ))}
          </div>
        )}
      </section>
      <p className="model-storage-note">Saved in this browser. Export a backup to keep a copy or move your work to another device.</p>
    </main>
  )
}
