import { generateText, stepCountIs } from 'ai'
import { MockLanguageModelV3 } from 'ai/test'
import {
  assertAiSdkWorkPaperSmokeProof,
  createAiSdkWorkPaperTools,
  modelUsage,
  requireToolOutput,
  requireWorkPaperReadResult,
  requireWorkPaperWriteResult,
} from './ai-sdk-workpaper-tool-smoke-shared.js'

const aiSdkTools = createAiSdkWorkPaperTools()
const model = new MockLanguageModelV3({
  provider: 'ai-sdk-test',
  modelId: 'deterministic-workpaper-tool-caller',
  doGenerate: async () => {
    if (model.doGenerateCalls.length === 1) {
      return {
        content: [
          {
            type: 'tool-call',
            toolCallId: 'call_read_summary',
            toolName: 'readWorkPaperSummary',
            input: JSON.stringify({ range: 'Summary!A1:B5' }),
          },
          {
            type: 'tool-call',
            toolCallId: 'call_set_input_b3',
            toolName: 'setWorkPaperInputCell',
            input: JSON.stringify({
              sheetName: 'Inputs',
              address: 'B3',
              value: 0.4,
            }),
          },
        ],
        finishReason: {
          unified: 'tool-calls',
          raw: 'tool-calls',
        },
        usage: modelUsage,
        warnings: [],
      }
    }

    return {
      content: [
        {
          type: 'text',
          text: 'Edited Inputs!B3. Expected ARR moved from 60,000 to 96,000, and the restored WorkPaper matched the post-write value.',
        },
      ],
      finishReason: {
        unified: 'stop',
        raw: 'stop',
      },
      usage: modelUsage,
      warnings: [],
    }
  },
})

const result = await generateText({
  model,
  tools: aiSdkTools,
  stopWhen: stepCountIs(2),
  prompt: [
    'Read the WorkPaper summary range Summary!A1:B5.',
    'Then set Inputs!B3 to 0.4.',
    'Return the edited cell, before and after expected ARR, and the restore check.',
  ].join('\n'),
})
const allToolCalls = result.steps.flatMap((step) => step.toolCalls)
const allToolResults = result.steps.flatMap((step) => step.toolResults)

const readResult = requireToolOutput(allToolResults, 'call_read_summary', 'readWorkPaperSummary')
const writeResult = requireWorkPaperWriteResult(requireToolOutput(allToolResults, 'call_set_input_b3', 'setWorkPaperInputCell'))

const smokeOutput = {
  apiShape: 'AI SDK generateText -> tool -> execute',
  modelProvider: `${model.provider}/${model.modelId}`,
  modelCallCount: model.doGenerateCalls.length,
  toolNames: Object.keys(aiSdkTools),
  toolCalls: allToolCalls.map(({ toolCallId, toolName, input }) => ({
    toolCallId,
    toolName,
    input,
  })),
  toolResults: allToolResults.map(({ toolCallId, toolName, output: toolOutput }) => ({
    toolCallId,
    toolName,
    output: toolOutput,
  })),
  text: result.text,
  readResult: requireWorkPaperReadResult(readResult),
  writeResult,
}

if (smokeOutput.modelCallCount !== 2) {
  throw new Error(`Expected two model calls, received ${smokeOutput.modelCallCount}`)
}
assertAiSdkWorkPaperSmokeProof(smokeOutput, 'AI SDK generateText -> tool -> execute')
console.log(JSON.stringify(smokeOutput, null, 2))
