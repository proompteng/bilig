import { describe, expect, it } from 'vitest'
import { createWorkbookAgentThreadResumeInput, createWorkbookAgentThreadStartInput } from './workbook-agent-codex-runtime.js'

describe('workbook agent codex runtime helpers', () => {
  it('creates thread start input with workbook-safe Codex defaults', () => {
    const input = createWorkbookAgentThreadStartInput()

    expect(input.model).toBeTypeOf('string')
    expect(input.approvalPolicy).toBe('never')
    expect(input.sandbox).toBe('danger-full-access')
    expect(input.config).toEqual({
      approval_policy: 'never',
      sandbox_mode: 'danger-full-access',
      network_access: true,
      web_search: 'live',
      tools: {
        view_image: true,
      },
    })
    expect(input.baseInstructions).toContain('Use workbook tools for workbook reads, edits, and verification.')
    expect(input.developerInstructions).toContain('Inspect before you edit unfamiliar cells or ranges.')
    expect(input.dynamicTools.length).toBeGreaterThan(0)
  })

  it('creates thread resume input with preserved workbook instructions', () => {
    const input = createWorkbookAgentThreadResumeInput('thr-123')

    expect(input.threadId).toBe('thr-123')
    expect(input.baseInstructions).toContain('Help with the active workbook only.')
    expect(input.developerInstructions).toContain('Apply workbook changes directly when the session policy allows it.')
  })
})
