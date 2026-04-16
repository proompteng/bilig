import readline from 'node:readline'

const state = {
  capabilities: null,
}

const TOOL_NAME_PATTERN = /^[a-zA-Z0-9_-]+$/

function write(payload) {
  process.stdout.write(`${JSON.stringify(payload)}\n`)
}

const reader = readline.createInterface({
  input: process.stdin,
  crlfDelay: Infinity,
})

reader.on('line', (line) => {
  const trimmed = line.trim()
  if (trimmed.length === 0) {
    return
  }
  const message = JSON.parse(trimmed)
  if (message.method === 'initialize') {
    state.capabilities = message.params?.capabilities ?? null
    write({
      id: message.id,
      result: {
        userAgent: 'fake-codex-app-server',
      },
    })
    return
  }
  if (message.method === 'thread/start') {
    if (
      process.env.BILIG_TEST_EXPECT_OTEL_STRIPPED === '1' &&
      (process.env.OTEL_EXPORTER_OTLP_ENDPOINT || process.env.OTEL_EXPORTER_OTLP_LOGS_ENDPOINT)
    ) {
      write({
        id: message.id,
        error: {
          code: -32602,
          message: 'OTEL env leaked into app-server process',
        },
      })
      return
    }
    const dynamicTools = Array.isArray(message.params?.dynamicTools) ? message.params.dynamicTools : []
    if (dynamicTools.length > 0 && state.capabilities?.experimentalApi !== true) {
      write({
        id: message.id,
        error: {
          code: -32602,
          message: 'thread/start.dynamicTools requires experimentalApi capability',
        },
      })
      return
    }
    const invalidTool = dynamicTools.find((tool) => !TOOL_NAME_PATTERN.test(tool?.name ?? ''))
    if (invalidTool) {
      write({
        id: message.id,
        error: {
          code: -32602,
          message: `Invalid dynamic tool name: ${invalidTool.name}`,
        },
      })
      return
    }
    write({
      id: message.id,
      result: {
        thread: {
          id: 'thr-fixture',
          preview: state.capabilities?.experimentalApi === true ? 'experimentalApi:true' : 'experimentalApi:false',
          turns: [],
        },
      },
    })
    return
  }
  if (message.method === 'thread/resume') {
    write({
      id: message.id,
      result: {
        thread: {
          id: message.params?.threadId ?? 'thr-fixture',
          preview: 'resumed',
          turns: [],
        },
      },
    })
    return
  }
  if (message.method === 'turn/start') {
    if (process.env.BILIG_TEST_EMIT_REASONING_DELTA === '1') {
      write({
        method: 'item/reasoning/delta',
        params: {
          threadId: message.params?.threadId ?? 'thr-fixture',
          turnId: 'turn-fixture',
          itemId: 'reasoning-fixture',
          delta: 'Examining staged changes',
        },
      })
    }
    write({
      id: message.id,
      result: {
        turn: {
          id: 'turn-fixture',
          status: 'inProgress',
          items: [],
          error: null,
        },
      },
    })
    return
  }
  if (typeof message.id === 'number') {
    write({
      id: message.id,
      result: null,
    })
  }
})
