import Fastify from 'fastify'
import { describe, expect, it } from 'vitest'
import { registerWorkPaperN8nRoutes } from './workpaper-n8n-routes.js'

describe('workpaper n8n forecast route', () => {
  it('returns formula readback and restore proof for one input edit', async () => {
    const app = Fastify({ logger: false })
    registerWorkPaperN8nRoutes(app)

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/api/workpaper/n8n/forecast',
        payload: {
          address: 'B3',
          value: 0.4,
        },
      })

      expect(response.statusCode).toBe(200)
      expect(response.json()).toMatchObject({
        verified: true,
        editedCell: 'Inputs!B3',
        before: {
          expectedCustomers: 5,
          expectedArr: 60000,
          expansionArr: 66000,
          targetGap: -34000,
        },
        after: {
          expectedCustomers: 8,
          expectedArr: 96000,
          expansionArr: 105600,
          targetGap: 5600,
        },
        formulaContracts: {
          expectedCustomers: '=Inputs!B2*Inputs!B3',
          expectedArr: '=B2*Inputs!B4',
          expansionArr: '=B3*Inputs!B5',
          targetGap: '=B4-100000',
        },
        checks: {
          previousValue: 0.25,
          newValue: 0.4,
          formulasPersisted: true,
          restoredMatchesAfter: true,
          computedOutputChanged: true,
        },
      })
    } finally {
      await app.close()
    }
  })

  it('rejects edits outside the demo input cells', async () => {
    const app = Fastify({ logger: false })
    registerWorkPaperN8nRoutes(app)

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/api/workpaper/n8n/forecast',
        payload: {
          address: 'C9',
          value: 0.4,
        },
      })

      expect(response.statusCode).toBe(400)
      expect(response.json()).toMatchObject({
        verified: false,
        error: 'Editable input address must be one of B2, B3, B4, B5',
      })
    } finally {
      await app.close()
    }
  })
})
