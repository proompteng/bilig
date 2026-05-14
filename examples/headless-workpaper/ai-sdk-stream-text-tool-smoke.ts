import { simulateReadableStream, stepCountIs, streamText } from 'ai'
import { MockLanguageModelV3 } from 'ai/test'
import {
  assertAiSdkWorkPaperSmokeProof,
  createAiSdkWorkPaperTools,
  modelUsage,
  requireToolOutput,
  requireWorkPaperReadResult,
  requireWorkPaperWriteResult,
} from './ai-sdk-workpaper-tool-smoke-shared.js'

type ModelStreamPart =
  Awaited<ReturnType<MockLanguageModelV3['doStream']>> extends {
    stream: ReadableStream<infer Part>
  }
    ? Part
    : never

const aiSdkTools = createAiSdkWorkPaperTools()
const streamedChunkTypes: string[] = []
const model = new MockLanguageModelV3({
  provider: 'ai-sdk-test',
  modelId: 'deterministic-workpaper-streaming-tool-caller',
  doStream: async () => {
    if (model.doStreamCalls.length === 1) {
      const chunks: ModelStreamPart[] = [
        {
          type: 'stream-start',
          warnings: [],
        },
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
        {
          type: 'finish',
          finishReason: {
            unified: 'tool-calls',
            raw: 'tool-calls',
          },
          usage: modelUsage,
        },
      ]

      return {
        stream: simulateReadableStream<ModelStreamPart>({ chunks }),
      }
    }

    const chunks: ModelStreamPart[] = [
      {
        type: 'stream-start',
        warnings: [],
      },
      {
        type: 'text-start',
        id: 'text_final',
      },
      {
        type: 'text-delta',
        id: 'text_final',
        delta: 'Edited Inputs!B3. ',
      },
      {
        type: 'text-delta',
        id: 'text_final',
        delta: 'Expected ARR moved from 60,000 to 96,000, and the restored WorkPaper matched the post-write value.',
      },
      {
        type: 'text-end',
        id: 'text_final',
      },
      {
        type: 'finish',
        finishReason: {
          unified: 'stop',
          raw: 'stop',
        },
        usage: modelUsage,
      },
    ]

    return {
      stream: simulateReadableStream<ModelStreamPart>({ chunks }),
    }
  },
})

const result = streamText({
  model,
  tools: aiSdkTools,
  stopWhen: stepCountIs(2),
  onChunk: ({ chunk }) => {
    streamedChunkTypes.push(chunk.type)
  },
  prompt: [
    'Read the WorkPaper summary range Summary!A1:B5.',
    'Then set Inputs!B3 to 0.4.',
    'Stream the final answer after the tools run.',
  ].join('\n'),
})
const [text, steps] = await Promise.all([result.text, result.steps])
const allToolCalls = steps.flatMap((step) => step.toolCalls)
const allToolResults = steps.flatMap((step) => step.toolResults)

const readResult = requireToolOutput(allToolResults, 'call_read_summary', 'readWorkPaperSummary')
const writeResult = requireWorkPaperWriteResult(requireToolOutput(allToolResults, 'call_set_input_b3', 'setWorkPaperInputCell'))

const smokeOutput = {
  apiShape: 'AI SDK streamText -> tool -> execute',
  modelProvider: `${model.provider}/${model.modelId}`,
  modelStreamCallCount: model.doStreamCalls.length,
  streamChunkTypes: streamedChunkTypes,
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
  text,
  readResult: requireWorkPaperReadResult(readResult),
  writeResult,
}

if (smokeOutput.modelStreamCallCount !== 2) {
  throw new Error(`Expected two streaming model calls, received ${smokeOutput.modelStreamCallCount}`)
}

for (const requiredChunkType of ['tool-call', 'tool-result', 'text-delta']) {
  if (!smokeOutput.streamChunkTypes.includes(requiredChunkType)) {
    throw new Error(`Expected streamed chunk type ${requiredChunkType}; received ${JSON.stringify(smokeOutput.streamChunkTypes)}`)
  }
}

assertAiSdkWorkPaperSmokeProof(smokeOutput, 'AI SDK streamText -> tool -> execute')
console.log(JSON.stringify(smokeOutput, null, 2))
