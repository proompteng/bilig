import { ScrollArea } from '@base-ui/react/scroll-area'
import type { ModelCalculation } from './model-calculation.js'
import type { ModelDocument, ModelInput } from './model-document.js'
import { formatModelNumber } from './model-format.js'
import { ModelField } from './ModelField.js'

export function ModelSensitivity(props: {
  readonly model: ModelDocument
  readonly input: ModelInput | undefined
  readonly step: number
  readonly calculation: ModelCalculation | null
  readonly onInputChange: (inputId: string) => void
  readonly onStepChange: (step: number) => void
}) {
  const { model, input, step, calculation } = props
  const sensitivity = calculation?.sensitivity
  return (
    <section className="model-comparison" aria-labelledby="model-sensitivity-title">
      <div className="model-section-heading">
        <h2 id="model-sensitivity-title">What if an assumption changes?</h2>
        <span className="model-calculation-status" role="status">
          {calculation ? 'Calculated' : 'Calculating…'}
        </span>
      </div>
      <p className="model-muted">Compare five variations against your current model. Your saved assumptions stay unchanged.</p>
      <div className="model-sensitivity-controls">
        <label>
          Assumption
          <select aria-label="Assumption" value={input?.id ?? ''} onChange={(event) => props.onInputChange(event.target.value)}>
            {model.inputs.map((entry) => (
              <option key={entry.id} value={entry.id}>
                {entry.label}
              </option>
            ))}
          </select>
        </label>
        <div className="model-sensitivity-step">
          <span>Step size{input?.format === 'percent' ? ' (%)' : ''}</span>
          <ModelField
            label="Step size"
            kind="positive-number"
            value={String(Number((input?.format === 'percent' ? step * 100 : step).toPrecision(12)))}
            onCommit={(value) => props.onStepChange(Number(value) / (input?.format === 'percent' ? 100 : 1))}
          />
        </div>
      </div>
      {input ? (
        <ScrollArea.Root className="model-table-scroll">
          <ScrollArea.Viewport tabIndex={0} role="region" aria-label="What-if comparison table">
            <table>
              <caption className="model-sr-only">
                Results when {input.label} changes in steps of {formatModelNumber(step, input.format)}
              </caption>
              <thead>
                <tr>
                  <th scope="col">Result / assumption</th>
                  {[-2, -1, 0, 1, 2].map((offset) => (
                    <th scope="col" key={offset}>
                      {offset === 0 ? 'Current' : `${offset > 0 ? '+' : '−'}${Math.abs(offset)} step${Math.abs(offset) === 1 ? '' : 's'}`}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                <tr>
                  <th scope="row">{input.label}</th>
                  {[-2, -1, 0, 1, 2].map((offset) => (
                    <td key={offset}>
                      {sensitivity
                        ? formatModelNumber(sensitivity.rows.find((row) => row.offset === offset)?.inputValue ?? input.value, input.format)
                        : '…'}
                    </td>
                  ))}
                </tr>
                {model.outputs.map((output) => (
                  <tr key={output.id} className="model-comparison-result">
                    <th scope="row">{output.label}</th>
                    {[-2, -1, 0, 1, 2].map((offset) => (
                      <td key={offset}>
                        {sensitivity?.rows.find((row) => row.offset === offset)?.values.find((value) => value.id === output.id)?.text ??
                          '…'}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </ScrollArea.Viewport>
          <ScrollArea.Scrollbar orientation="horizontal" className="model-scrollbar">
            <ScrollArea.Thumb className="model-scrollbar-thumb" />
          </ScrollArea.Scrollbar>
        </ScrollArea.Root>
      ) : (
        <p className="model-muted">Add an assumption to explore variations.</p>
      )}
    </section>
  )
}
