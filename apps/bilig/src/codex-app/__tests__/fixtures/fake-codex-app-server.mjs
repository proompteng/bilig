import readline from 'node:readline'

const state = {
  capabilities: null,
}

const args = new Set(process.argv.slice(2))
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
        codexHome: '/tmp/fake-codex-home',
        platformFamily: 'unix',
        platformOs: 'macos',
      },
    })
    return
  }
  if (message.method === 'thread/start') {
    if (args.has('--expect-otel-stripped') && (process.env.OTEL_EXPORTER_OTLP_ENDPOINT || process.env.OTEL_EXPORTER_OTLP_LOGS_ENDPOINT)) {
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
    const preview = args.has('--echo-thread-start')
      ? JSON.stringify({
          experimentalApi: state.capabilities?.experimentalApi === true,
          approvalPolicy: message.params?.approvalPolicy ?? null,
          sandbox: message.params?.sandbox ?? null,
          config: message.params?.config ?? null,
        })
      : state.capabilities?.experimentalApi === true
        ? 'experimentalApi:true'
        : 'experimentalApi:false'
    write({
      id: message.id,
      result: {
        thread: {
          id: 'thr-fixture',
          preview,
          turns: [],
        },
      },
    })
    if (args.has('--exit-after-thread-start')) {
      setTimeout(() => {
        process.exit(17)
      }, 0)
    }
    return
  }
  if (message.method === 'thread/resume') {
    const preview = args.has('--echo-thread-resume')
      ? JSON.stringify({
          threadId: message.params?.threadId ?? null,
          approvalPolicy: message.params?.approvalPolicy ?? null,
          sandbox: message.params?.sandbox ?? null,
          cwd: message.params?.cwd ?? null,
          runtimeWorkspaceRoots: message.params?.runtimeWorkspaceRoots ?? null,
          config: message.params?.config ?? null,
        })
      : 'resumed'
    write({
      id: message.id,
      result: {
        thread: {
          id: message.params?.threadId ?? 'thr-fixture',
          preview,
          turns: [],
        },
      },
    })
    return
  }
  if (message.method === 'turn/start') {
    if (args.has('--exit-during-turn-start')) {
      setTimeout(() => {
        process.exit(42)
      }, 0)
      return
    }
    if (args.has('--emit-reasoning-delta')) {
      write({
        method: 'item/reasoning/textDelta',
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
