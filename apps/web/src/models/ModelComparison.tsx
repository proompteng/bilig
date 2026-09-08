import type { ModelCalculation, ModelResult } from './model-calculation.js'
import { ScrollArea } from '@base-ui/react/scroll-area'
import type { ModelDocument, ModelScenario } from './model-document.js'
import { formatModelNumber } from './model-format.js'

function resultText(result: ModelResult | undefined): string {
  return result?.text ?? '—'
}

export function ModelComparison(props: {
  readonly model: ModelDocument
  readonly calculation: ModelCalculation | null
  readonly onRestore: (scenario: ModelScenario) => void
}) {
  const { model, calculation } = props
  return (
    <section className="model-comparison" aria-labelledby="model-comparison-title">
      <div className="model-section-heading">
        <h2 id="model-comparison-title">Scenario comparison</h2>
      </div>
      <p className="model-muted">
        Each saved scenario keeps its own assumptions and formulas. Restore one to use it as your working model.
      </p>
      {model.scenarios.length === 0 ? (
        <div className="model-empty">
          <h3>No scenarios yet</h3>
          <p className="model-muted">Save the current model as a scenario, then change an assumption to compare outcomes.</p>
        </div>
      ) : (
        <ScrollArea.Root className="model-table-scroll">
          <ScrollArea.Viewport tabIndex={0} role="region" aria-label="Scenario comparison table">
            <table>
              <caption className="model-sr-only">Current model compared with saved scenarios</caption>
              <thead>
                <tr>
                  <th scope="col">Result / assumption</th>
                  <th scope="col">Current</th>
                  {model.scenarios.map((scenario) => (
                    <th scope="col" key={scenario.id}>
                      {scenario.name}
                      <button className="model-text-button" onClick={() => props.onRestore(scenario)}>
                        Restore<span className="model-sr-only"> {scenario.name}</span>
                      </button>
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {model.outputs.map((output) => (
                  <tr key={output.id} className="model-comparison-result">
                    <th scope="row">{output.label}</th>
                    <td>{resultText(calculation?.current.find((value) => value.id === output.id))}</td>
                    {model.scenarios.map((scenario) => (
                      <td key={scenario.id}>
                        {resultText(
                          calculation?.scenarios.find((item) => item.id === scenario.id)?.values.find((value) => value.id === output.id),
                        )}
                      </td>
                    ))}
                  </tr>
                ))}
                {model.inputs.map((input) => (
                  <tr key={input.id}>
                    <th scope="row">{input.label}</th>
                    <td>{formatModelNumber(input.value, input.format)}</td>
                    {model.scenarios.map((scenario) => {
                      const saved = scenario.inputs.find((entry) => entry.id === input.id)
                      return <td key={scenario.id}>{saved ? formatModelNumber(saved.value, saved.format) : '—'}</td>
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </ScrollArea.Viewport>
          <ScrollArea.Scrollbar orientation="horizontal" className="model-scrollbar">
            <ScrollArea.Thumb className="model-scrollbar-thumb" />
          </ScrollArea.Scrollbar>
        </ScrollArea.Root>
      )}
    </section>
  )
}
