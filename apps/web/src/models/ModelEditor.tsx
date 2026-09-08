import { useEffect, useState } from 'react'
import { Archive, ArrowLeft, Download, Plus } from 'lucide-react'
import { captureModelScenario, MODEL_LIMITS, type ModelDocument } from './model-document.js'
import { ModelField, ModelFormatSelect } from './ModelField.js'
import { ModelComparison } from './ModelComparison.js'
import { useModelCalculation } from './use-model-calculation.js'
import { useModelEditor } from './use-model-editor.js'
import { downloadModelFile } from './model-download.js'

export function ModelEditor(props: {
  readonly initial: ModelDocument
  readonly onSaved: (model: ModelDocument) => void
  readonly onBack: () => void
  readonly onNavigationGuard: (guard: (() => Promise<boolean>) | null) => void
}) {
  const { model, change, saveState, waitForSave, reloadSavedVersion } = useModelEditor(props.initial, props.onSaved)
  const { onNavigationGuard } = props
  useEffect(() => {
    onNavigationGuard(waitForSave)
    return () => onNavigationGuard(null)
  }, [onNavigationGuard, waitForSave])
  const calculationState = useModelCalculation(model)
  const calculation = calculationState.kind === 'ready' ? calculationState.calculation : null
  const [view, setView] = useState<'model' | 'compare'>('model')
  const [scenarioName, setScenarioName] = useState('')
  const [notice, setNotice] = useState('')
  const isSaved = saveState.kind === 'saved'
  const backup = () => downloadModelFile(model.title, 'bilig-model.json', JSON.stringify(model, null, 2))

  return (
    <main className="model-editor" id="model-main">
      <div className="model-editor-topbar">
        <button className="model-text-button" onClick={props.onBack} disabled={!isSaved}>
          <ArrowLeft size={15} /> Models
        </button>
        <span className="model-save-status" role="status">
          {saveState.kind === 'saved' ? 'Saved in this browser' : saveState.kind === 'saving' ? 'Saving…' : 'Unsaved changes'}
        </span>
        <div className="model-actions">
          <button className="model-button" onClick={backup}>
            <Download size={15} /> Export backup
          </button>
          <button
            className="model-button"
            disabled={!calculation}
            onClick={() => {
              if (calculation) downloadModelFile(model.title, 'workpaper.json', calculation.workpaperJson)
            }}
          >
            WorkPaper JSON
          </button>
        </div>
      </div>
      {saveState.kind === 'error' ? (
        <div className="model-error" role="alert">
          <p>{saveState.message}</p>
          <button className="model-button" onClick={backup}>
            Export unsaved edits
          </button>{' '}
          <button className="model-button" onClick={reloadSavedVersion}>
            Reload saved version
          </button>
        </div>
      ) : null}
      <div className="model-editor-heading">
        <p className="model-eyebrow">Model</p>
        <h1>
          <ModelField
            label="Model name"
            value={model.title}
            className="model-title-field"
            onCommit={(title) => change((current) => ({ ...current, title }))}
          />
        </h1>
        <ModelField
          label="Model description"
          value={model.description}
          className="model-description-field"
          maxLength={1000}
          onCommit={(description) => change((current) => ({ ...current, description }))}
        />
      </div>
      <div className="model-editor-controls">
        <div className="model-view-switch" aria-label="Model view">
          <button aria-pressed={view === 'model'} onClick={() => setView('model')}>
            Model
          </button>
          <button aria-pressed={view === 'compare'} onClick={() => setView('compare')}>
            Compare <span className="model-count">{model.scenarios.length}</span>
          </button>
        </div>
        <form
          className="model-scenario-form"
          onSubmit={(event) => {
            event.preventDefault()
            const name = scenarioName.trim()
            if (!name || model.scenarios.length >= MODEL_LIMITS.scenarios) return
            change((current) => captureModelScenario(current, name))
            setScenarioName('')
            setNotice(`Added scenario ${name}.`)
          }}
        >
          <input
            aria-label="Scenario name"
            placeholder="Name this scenario"
            value={scenarioName}
            maxLength={120}
            onChange={(event) => setScenarioName(event.target.value)}
          />
          <button
            className="model-button model-primary"
            disabled={!scenarioName.trim() || !calculation || model.scenarios.length >= MODEL_LIMITS.scenarios}
          >
            Save scenario
          </button>
        </form>
      </div>
      {notice ? (
        <p className="model-notice" role="status">
          {notice}
        </p>
      ) : null}
      {calculationState.kind === 'error' ? (
        <div className="model-error" role="alert">
          {calculationState.message}
        </div>
      ) : null}
      {view === 'compare' ? (
        <ModelComparison
          model={model}
          calculation={calculation}
          onRestore={(scenario) => {
            change((current) => ({ ...current, ...structuredClone({ inputs: scenario.inputs, outputs: scenario.outputs }) }))
            setNotice(`Restored ${scenario.name}. Saved scenarios are unchanged.`)
          }}
        />
      ) : (
        <div className="model-working-area">
          <section className="model-inputs" aria-labelledby="model-inputs-title">
            <div className="model-section-heading">
              <h2 id="model-inputs-title">Assumptions</h2>
              <span className="model-muted">Inputs</span>
            </div>
            <p className="model-muted model-section-description">Change a value to recalculate your results.</p>
            {model.inputs.map((input, index) => (
              <div className="model-input-row" key={input.id}>
                <div>
                  <ModelField
                    label={`Assumption ${index + 1} name`}
                    value={input.label}
                    className="model-label-field"
                    onCommit={(label) =>
                      change((current) => ({
                        ...current,
                        inputs: current.inputs.map((entry) => (entry.id === input.id ? { ...entry, label } : entry)),
                      }))
                    }
                  />
                  <code>Inputs!B{index + 2}</code>
                </div>
                <div className="model-input-value">
                  <ModelField
                    label={input.label}
                    kind="number"
                    value={String(input.format === 'percent' ? Number((input.value * 100).toPrecision(12)) : input.value)}
                    onCommit={(raw) =>
                      change((current) => ({
                        ...current,
                        inputs: current.inputs.map((entry) =>
                          entry.id === input.id ? { ...entry, value: Number(raw) / (input.format === 'percent' ? 100 : 1) } : entry,
                        ),
                      }))
                    }
                  />
                  <ModelFormatSelect
                    label={`${input.label} format`}
                    value={input.format}
                    onChange={(format) =>
                      change((current) => ({
                        ...current,
                        inputs: current.inputs.map((entry) => (entry.id === input.id ? { ...entry, format } : entry)),
                      }))
                    }
                  />
                </div>
              </div>
            ))}
            <button
              className="model-text-button model-add"
              disabled={model.inputs.length >= MODEL_LIMITS.inputs}
              onClick={() =>
                change((current) => ({
                  ...current,
                  inputs: [
                    ...current.inputs,
                    { id: crypto.randomUUID(), label: `Input ${current.inputs.length + 1}`, value: 0, format: 'number' },
                  ],
                }))
              }
            >
              <Plus size={15} /> Add assumption
            </button>
          </section>
          <section className="model-results" aria-labelledby="model-results-title">
            <div className="model-section-heading">
              <h2 id="model-results-title">Results</h2>
              <span className="model-calculation-status" role="status">
                {calculationState.kind === 'calculating' ? 'Calculating…' : calculation ? 'Calculated' : 'Needs attention'}
              </span>
            </div>
            <p className="model-muted model-section-description">Every result comes from an editable workbook formula.</p>
            {model.outputs.map((output, index) => {
              const result = calculation?.current.find((entry) => entry.id === output.id)
              return (
                <div className="model-output-row" key={output.id}>
                  <div className="model-output-heading">
                    <ModelField
                      label={`Result ${index + 1} name`}
                      value={output.label}
                      className="model-label-field"
                      onCommit={(label) =>
                        change((current) => ({
                          ...current,
                          outputs: current.outputs.map((entry) => (entry.id === output.id ? { ...entry, label } : entry)),
                        }))
                      }
                    />
                    <code>Results!B{index + 2}</code>
                  </div>
                  <div className="model-output-value-row">
                    <output
                      aria-label={output.label}
                      className={`model-output-value ${result?.kind === 'error' ? 'model-result-error' : ''}`}
                    >
                      {result?.text ?? '—'}
                    </output>
                    <ModelFormatSelect
                      label={`${output.label} format`}
                      value={output.format}
                      onChange={(format) =>
                        change((current) => ({
                          ...current,
                          outputs: current.outputs.map((entry) => (entry.id === output.id ? { ...entry, format } : entry)),
                        }))
                      }
                    />
                  </div>
                  <ModelField
                    label={`${output.label} formula`}
                    value={output.formula}
                    kind="formula"
                    maxLength={2048}
                    className="model-formula-field"
                    onCommit={(formula) =>
                      change((current) => ({
                        ...current,
                        outputs: current.outputs.map((entry) => (entry.id === output.id ? { ...entry, formula } : entry)),
                      }))
                    }
                  />
                </div>
              )
            })}
            <button
              className="model-text-button model-add"
              disabled={model.outputs.length >= MODEL_LIMITS.outputs}
              onClick={() =>
                change((current) => ({
                  ...current,
                  outputs: [
                    ...current.outputs,
                    { id: crypto.randomUUID(), label: `Result ${current.outputs.length + 1}`, formula: '=0', format: 'number' },
                  ],
                }))
              }
            >
              <Plus size={15} /> Add result
            </button>
          </section>
        </div>
      )}
      <footer className="model-editor-footer">
        <p className="model-muted">
          {calculation?.restoreVerified
            ? 'WorkPaper export restored with matching results.'
            : calculation
              ? 'Results changed after export restore. Check for volatile formulas before using this result.'
              : 'Inputs use the addresses shown. Percent inputs are stored as fractions.'}
        </p>
        <button className="model-text-button" disabled={!isSaved} onClick={() => change((current) => ({ ...current, status: 'archived' }))}>
          <Archive size={14} /> Archive model
        </button>
      </footer>
    </main>
  )
}
