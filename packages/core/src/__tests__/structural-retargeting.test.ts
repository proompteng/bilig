import { describe, expect, it } from 'vitest'
import { createTemplateBank } from '../formula/template-bank.js'
import { rewriteTemplateForStructuralTransform, retargetStructurallyRewrittenTemplateInstance } from '../formula/structural-retargeting.js'

describe('structural-retargeting', () => {
  it('rewrites a template once for a structural insert and retargets later instances from that rewritten family', () => {
    const bank = createTemplateBank()
    const first = bank.resolve('A1+B1', 0, 2)
    const template = bank.get(first.templateId)
    expect(template).toBeDefined()

    const rewrittenTemplate = rewriteTemplateForStructuralTransform({
      template: template!,
      ownerSheetName: 'Sheet1',
      targetSheetName: 'Sheet1',
      transform: {
        kind: 'insert',
        axis: 'column',
        start: 1,
        count: 1,
      },
    })
    expect(rewrittenTemplate).toBeDefined()
    expect(rewrittenTemplate?.source).toBe('A1+C1')
    expect(rewrittenTemplate?.baseCol).toBe(3)

    const retargeted = retargetStructurallyRewrittenTemplateInstance({
      rewrittenTemplate: rewrittenTemplate!,
      ownerRow: 1,
      ownerCol: 3,
    })
    expect(retargeted.source).toBe('A2+C2')
    expect(retargeted.compiled.deps).toEqual(['A2', 'C2'])
  })

  it('falls back when the template base owner is deleted by the structural transform', () => {
    const bank = createTemplateBank()
    const first = bank.resolve('B1+1', 0, 1)
    const template = bank.get(first.templateId)
    expect(template).toBeDefined()

    const rewrittenTemplate = rewriteTemplateForStructuralTransform({
      template: template!,
      ownerSheetName: 'Sheet1',
      targetSheetName: 'Sheet1',
      transform: {
        kind: 'delete',
        axis: 'column',
        start: 1,
        count: 1,
      },
    })

    expect(rewrittenTemplate).toBeUndefined()
  })
})
