import { describe, expect, it } from 'vitest'
import {
  agentPanelComposerSendButtonClass,
  agentPanelDisclosureChevronClass,
  agentPanelDisclosureContentClass,
  agentPanelDisclosureSummaryClass,
  agentPanelDisclosureTriggerClass,
} from '../workbook-agent-panel-primitives.js'

describe('workbook agent panel primitives', () => {
  it('keeps the composer send button circular', () => {
    const className = agentPanelComposerSendButtonClass()

    expect(className).toContain('h-8')
    expect(className).toContain('w-8')
    expect(className).toContain('rounded-full')
  })

  it('keeps disclosure rows tightly aligned around the chevron', () => {
    expect(agentPanelDisclosureTriggerClass()).toContain('grid')
    expect(agentPanelDisclosureTriggerClass()).not.toContain('flex')
    expect(agentPanelDisclosureTriggerClass()).toContain('items-start')
    expect(agentPanelDisclosureContentClass({ open: false })).toContain('flex-wrap')
    expect(agentPanelDisclosureContentClass({ open: true })).toContain('grid-cols-1')
    expect(agentPanelDisclosureSummaryClass({ open: true })).toContain('whitespace-normal')
    expect(agentPanelDisclosureSummaryClass({ open: false })).toContain('truncate')
    expect(agentPanelDisclosureChevronClass({ open: false })).toContain('size-3.5')
    expect(agentPanelDisclosureChevronClass({ open: false })).toContain('mt-0.5')
  })
})
