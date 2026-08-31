import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  WorkbookAgentCodexRuntime,
  type WorkbookAgentCodexRuntimeOptions,
  createWorkbookAgentThreadResumeInput,
  createWorkbookAgentThreadStartInput,
  resolveWorkbookAgentCodexSecurityConfig,
} from './workbook-agent-codex-runtime.js'
import { createZeroSyncStub } from './workbook-agent-service.test-helpers.js'

function createRuntimeOptions(env: Readonly<Record<string, string | undefined>>): WorkbookAgentCodexRuntimeOptions {
  return {
    zeroSyncService: createZeroSyncStub(),
    env,
    now: () => 0,
    maxCodexClients: 1,
    maxConcurrentTurnsPerCodexClient: 1,
    maxQueuedTurnsPerCodexClient: 1,
    getSessionByThreadId: () => {
      throw new Error('not used by the security boundary test')
    },
    tryGetSessionByThreadId: () => null,
    listSessions: () => [],
    resolveTurnActorUserId: () => {
      throw new Error('not used by the security boundary test')
    },
    resolveTurnContext: () => {
      throw new Error('not used by the security boundary test')
    },
    stageReviewBundle: () => undefined,
    shouldApplyToolBundleImmediately: () => false,
    applyToolBundleAutomatically: async () => null,
    persistSessionState: async () => undefined,
    emitSnapshot: () => undefined,
    emit: () => undefined,
    finalizeCompletedTurn: async () => undefined,
    startWorkflow: async () => {
      throw new Error('not used by the security boundary test')
    },
  }
}

describe('workbook agent codex runtime helpers', () => {
  it('creates thread start input with workbook-safe Codex defaults', () => {
    const input = createWorkbookAgentThreadStartInput()

    expect(input.model).toBeTypeOf('string')
    expect(input.approvalPolicy).toBe('on-request')
    expect(input.sandbox).toBe('read-only')
    expect(input.config).toEqual({
      approval_policy: 'on-request',
      sandbox_mode: 'read-only',
      sandbox_workspace_write: {
        network_access: false,
      },
      web_search: 'disabled',
      mcp_servers: {},
      features: expect.objectContaining({
        apps: false,
        computer_use: false,
        plugins: false,
        shell_tool: false,
        unified_exec: false,
      }),
    })
    expect(input.cwd).toBe(os.tmpdir())
    expect(input.runtimeWorkspaceRoots).toEqual([])
    expect(input.environments).toEqual([])
    expect(input.baseInstructions).toContain('Use workbook tools for workbook reads, edits, and verification.')
    expect(input.developerInstructions).toContain('Inspect before you edit unfamiliar cells or ranges.')
    expect(input.dynamicTools.length).toBeGreaterThan(0)
  })

  it('creates thread resume input with preserved workbook instructions', () => {
    const input = createWorkbookAgentThreadResumeInput('thr-123')

    expect(input.threadId).toBe('thr-123')
    expect(input.approvalPolicy).toBe('on-request')
    expect(input.sandbox).toBe('read-only')
    expect(input.cwd).toBe(os.tmpdir())
    expect(input.runtimeWorkspaceRoots).toEqual([])
    expect(input.config).toEqual(
      expect.objectContaining({
        sandbox_workspace_write: {
          network_access: false,
        },
        web_search: 'disabled',
        mcp_servers: {},
        features: expect.objectContaining({
          apps: false,
          shell_tool: false,
          unified_exec: false,
        }),
      }),
    )
    expect(input.baseInstructions).toContain('Help with the active workbook only.')
    expect(input.developerInstructions).toContain('Apply workbook changes directly when the session policy allows it.')
  })

  it('resolves a production-safe app-server boundary and filters child environment secrets', () => {
    const codexHome = mkdtempSync(path.join(os.tmpdir(), 'bilig-codex-production-test-'))
    const config = resolveWorkbookAgentCodexSecurityConfig({
      NODE_ENV: 'production',
      BILIG_CODEX_BIN: '/opt/codex/bin/codex',
      BILIG_CODEX_HOME: codexHome,
      PATH: '/usr/bin',
      OPENAI_API_KEY: 'test-key',
      HOME: '/Users/alex',
      CODEX_HOME: '/Users/alex/.codex',
      XDG_CONFIG_HOME: '/Users/alex/.config',
      DATABASE_URL: 'postgres://db.example.test/app',
      BILIG_SESSION_SECRET: 'session-secret',
      AWS_SECRET_ACCESS_KEY: 'aws-secret',
    })

    expect(config.command).toBe('/opt/codex/bin/codex')
    expect(config.args[0]).toBe('app-server')
    expect(config.args).toContain('--strict-config')
    expect(config.args).toContain('mcp_servers={}')
    expect(config.args).toContain('features.shell_tool=false')
    expect(config.args).toContain('features.unified_exec=false')
    expect(config.args).toContain('features.apps=false')
    expect(config.cwd).toBe(os.tmpdir())
    expect(config.cwd).not.toBe(process.cwd())
    expect(config.env).toMatchObject({
      PATH: '/usr/bin',
      OPENAI_API_KEY: 'test-key',
      OTEL_SDK_DISABLED: 'true',
      HOME: codexHome,
      CODEX_HOME: codexHome,
      XDG_CONFIG_HOME: path.join(codexHome, 'xdg-config'),
      XDG_DATA_HOME: path.join(codexHome, 'xdg-data'),
      XDG_CACHE_HOME: path.join(codexHome, 'xdg-cache'),
    })
    expect(config.env).not.toHaveProperty('DATABASE_URL')
    expect(config.env).not.toHaveProperty('BILIG_SESSION_SECRET')
    expect(config.env).not.toHaveProperty('AWS_SECRET_ACCESS_KEY')
    expect(config.env['CODEX_HOME']).not.toBe('/Users/alex/.codex')
    expect(config.unsafeLocalOptIn).toBe(false)
    rmSync(codexHome, { recursive: true, force: true })
  })

  it('rejects the unsafe local opt-in outside development', () => {
    expect(() =>
      resolveWorkbookAgentCodexSecurityConfig({
        NODE_ENV: 'production',
        BILIG_CODEX_ALLOW_UNSAFE_LOCAL: 'true',
      }),
    ).toThrow('BILIG_CODEX_ALLOW_UNSAFE_LOCAL is only permitted when NODE_ENV=development')
  })

  it('rejects an inherited user Codex home in the safe policy', () => {
    expect(() =>
      resolveWorkbookAgentCodexSecurityConfig({
        NODE_ENV: 'production',
        HOME: '/Users/alex',
        CODEX_HOME: '/Users/alex/.codex',
        BILIG_CODEX_HOME: '/Users/alex/.codex',
      }),
    ).toThrow('BILIG_CODEX_HOME must use a dedicated service directory')
  })

  it('rejects policy files in the dedicated safe Codex home', () => {
    const codexHome = mkdtempSync(path.join(os.tmpdir(), 'bilig-codex-security-test-'))
    try {
      writeFileSync(path.join(codexHome, 'config.toml'), '[mcp_servers.unsafe]\nurl = "https://example.test"\n')

      expect(() =>
        resolveWorkbookAgentCodexSecurityConfig({
          NODE_ENV: 'production',
          BILIG_CODEX_HOME: codexHome,
        }),
      ).toThrow('BILIG_CODEX_HOME must not contain config.toml')
    } finally {
      rmSync(codexHome, { recursive: true, force: true })
    }
  })

  it('rejects unsafe production configuration while the runtime is starting', () => {
    expect(
      () =>
        new WorkbookAgentCodexRuntime(
          createRuntimeOptions({
            NODE_ENV: 'production',
            BILIG_CODEX_ALLOW_UNSAFE_LOCAL: 'true',
          }),
        ),
    ).toThrow('BILIG_CODEX_ALLOW_UNSAFE_LOCAL is only permitted when NODE_ENV=development')
  })

  it('uses the runtime-injected policy for thread start input', () => {
    const runtime = new WorkbookAgentCodexRuntime(
      createRuntimeOptions({
        NODE_ENV: 'development',
        BILIG_CODEX_ALLOW_UNSAFE_LOCAL: 'true',
      }),
    )

    expect(runtime.createThreadStartInput()).toMatchObject({
      approvalPolicy: 'on-request',
      sandbox: 'workspace-write',
      config: {
        approval_policy: 'on-request',
        sandbox_mode: 'workspace-write',
        sandbox_workspace_write: {
          network_access: true,
        },
        web_search: 'live',
      },
    })
  })

  it('keeps local networked behavior explicit and bounded', () => {
    const config = resolveWorkbookAgentCodexSecurityConfig({
      NODE_ENV: 'development',
      BILIG_CODEX_ALLOW_UNSAFE_LOCAL: '1',
    })

    expect(config.approvalPolicy).toBe('on-request')
    expect(config.sandbox).toBe('workspace-write')
    expect(config.networkAccess).toBe(true)
    expect(config.webSearch).toBe('live')
    expect(config.cwd).toBe(os.tmpdir())
    expect(config.env).not.toHaveProperty('BILIG_CODEX_ALLOW_UNSAFE_LOCAL')
    expect(config.threadConfig).toEqual(
      expect.objectContaining({
        mcp_servers: {},
        features: expect.objectContaining({
          apps: false,
          shell_tool: false,
          unified_exec: false,
        }),
      }),
    )
  })
})
