import {
  createAgentSessionSnapshot,
  createReviewQueueItem,
  createSyncServer,
  createZeroSyncStub,
  createWorkbookAgentServiceError,
  createWorkbookAgentServiceStub,
  describe,
  expect,
  it,
  vi,
} from './sync-server-test-helpers.js'

describe('sync-server workbook agent public routes', () => {
  it('reviews workbook agent bundles through the public thread route', async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
      }),
    )
    const reviewReviewItem = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
        reviewQueueItems: [
          createReviewQueueItem({
            id: 'bundle-1',
            documentId: 'doc-1',
            threadId: 'thr-2',
            turnId: 'turn-1',
            goalText: 'Normalize shared workbook',
            summary: 'Normalize shared workbook',
            scope: 'workbook',
            riskClass: 'high',
            baseRevision: 4,
            createdAtUnixMs: 10,
            context: null,
            commands: [],
            affectedRanges: [],
            estimatedAffectedCells: 0,
            sharedReview: {
              ownerUserId: 'alex@example.com',
              status: 'approved',
              decidedByUserId: 'alex@example.com',
              decidedAtUnixMs: 12,
              recommendations: [],
            },
          }),
        ],
      }),
    )

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        reviewReviewItem,
      }),
    })

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/v2/documents/doc-1/chat/threads/thr-2/review-items/bundle-1/review',
        payload: {
          decision: 'approved',
        },
      })

      expect(response.statusCode).toBe(200)
      expect(reviewReviewItem).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          threadId: 'thr-2',
          reviewItemId: 'bundle-1',
          body: {
            decision: 'approved',
          },
        }),
      )
      expect(response.json()).toEqual(
        expect.objectContaining({
          reviewQueueItems: [
            expect.objectContaining({
              reviewMode: 'ownerReview',
              status: 'approved',
              decidedByUserId: 'alex@example.com',
            }),
          ],
        }),
      )
    } finally {
      await app.close()
    }
  })

  it('updates workbook agent context through the public thread route', async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
      }),
    )
    const updateContext = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
      }),
    )

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        updateContext,
      }),
    })

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/v2/documents/doc-1/chat/threads/thr-2/context',
        payload: {
          context: {
            selection: {
              sheetName: 'Sheet1',
              address: 'B2',
            },
            viewport: {
              rowStart: 1,
              rowEnd: 11,
              colStart: 1,
              colEnd: 6,
            },
          },
        },
      })

      expect(response.statusCode).toBe(200)
      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          body: {
            threadId: 'thr-2',
          },
        }),
      )
      expect(updateContext).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          threadId: 'thr-2',
          body: {
            context: {
              selection: {
                sheetName: 'Sheet1',
                address: 'B2',
              },
              viewport: {
                rowStart: 1,
                rowEnd: 11,
                colStart: 1,
                colEnd: 6,
              },
            },
          },
        }),
      )
    } finally {
      await app.close()
    }
  })

  it('interrupts workbook agent turns through the public thread route', async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
      }),
    )
    const interruptTurn = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
        status: 'idle',
        activeTurnId: null,
      }),
    )

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        interruptTurn,
      }),
    })

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/v2/documents/doc-1/chat/threads/thr-2/interrupt',
      })

      expect(response.statusCode).toBe(200)
      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          body: {
            threadId: 'thr-2',
          },
        }),
      )
      expect(interruptTurn).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          threadId: 'thr-2',
        }),
      )
    } finally {
      await app.close()
    }
  })

  it('applies staged workbook bundles through the public thread route', async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
      }),
    )
    const applyReviewItem = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
        reviewQueueItems: [],
      }),
    )

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        applyReviewItem,
      }),
    })

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/v2/documents/doc-1/chat/threads/thr-2/review-items/bundle-1/apply',
        payload: {
          commandIndexes: [1],
        },
      })

      expect(response.statusCode).toBe(200)
      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          body: {
            threadId: 'thr-2',
          },
        }),
      )
      expect(applyReviewItem).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          threadId: 'thr-2',
          reviewItemId: 'bundle-1',
          appliedBy: 'user',
          commandIndexes: [1],
        }),
      )
    } finally {
      await app.close()
    }
  })

  it('returns a structured conflict envelope when agent apply rejects a stale preview', async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-stale',
      }),
    )
    const applyReviewItem = vi.fn(async () => {
      throw createWorkbookAgentServiceError({
        code: 'WORKBOOK_AGENT_PREVIEW_STALE',
        message: 'Workbook changed after preview. Replay the plan to stage a fresh review item.',
        statusCode: 409,
        retryable: true,
      })
    })

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        applyReviewItem,
      }),
    })

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/v2/documents/doc-1/chat/threads/thr-stale/review-items/bundle-1/apply',
        payload: {},
      })

      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          body: {
            threadId: 'thr-stale',
          },
        }),
      )
      expect(response.statusCode).toBe(409)
      expect(response.json()).toEqual(
        expect.objectContaining({
          error: 'WORKBOOK_AGENT_PREVIEW_STALE',
          message: 'Workbook changed after preview. Replay the plan to stage a fresh review item.',
          retryable: true,
        }),
      )
    } finally {
      await app.close()
    }
  })

  it('dismisses staged workbook bundles through the public thread route', async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
      }),
    )
    const dismissReviewItem = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
      }),
    )

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        dismissReviewItem,
      }),
    })

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/v2/documents/doc-1/chat/threads/thr-2/review-items/bundle-1/dismiss',
      })

      expect(response.statusCode).toBe(200)
      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          body: {
            threadId: 'thr-2',
          },
        }),
      )
      expect(dismissReviewItem).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          threadId: 'thr-2',
          reviewItemId: 'bundle-1',
        }),
      )
    } finally {
      await app.close()
    }
  })

  it('replays prior execution records through the public thread route', async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
      }),
    )
    const replayExecutionRecord = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: 'thr-2',
        reviewQueueItems: [
          createReviewQueueItem({
            id: 'bundle-replay-1',
            documentId: 'doc-1',
            threadId: 'thr-2',
            turnId: 'replay:run-1:10',
            goalText: 'Reapply formatting',
            summary: 'Format Sheet1!A1',
            scope: 'selection',
            riskClass: 'low',
            baseRevision: 4,
            createdAtUnixMs: 10,
            context: {
              selection: {
                sheetName: 'Sheet1',
                address: 'A1',
              },
              viewport: {
                rowStart: 0,
                rowEnd: 10,
                colStart: 0,
                colEnd: 5,
              },
            },
            commands: [
              {
                kind: 'formatRange',
                range: {
                  sheetName: 'Sheet1',
                  startAddress: 'A1',
                  endAddress: 'A1',
                },
                patch: {
                  font: {
                    bold: true,
                  },
                },
              },
            ],
            affectedRanges: [
              {
                sheetName: 'Sheet1',
                startAddress: 'A1',
                endAddress: 'A1',
                role: 'target',
              },
            ],
            estimatedAffectedCells: 1,
            sharedReview: null,
          }),
        ],
      }),
    )

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        replayExecutionRecord,
      }),
    })

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/v2/documents/doc-1/chat/threads/thr-2/runs/run-1/replay',
      })

      expect(response.statusCode).toBe(200)
      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          body: {
            threadId: 'thr-2',
          },
        }),
      )
      expect(replayExecutionRecord).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: 'doc-1',
          threadId: 'thr-2',
          recordId: 'run-1',
        }),
      )
      expect(response.json()).toEqual(
        expect.objectContaining({
          reviewQueueItems: [expect.objectContaining({ id: 'bundle-replay-1' })],
        }),
      )
    } finally {
      await app.close()
    }
  })

  it('returns a structured not-found envelope when the chat thread event stream is stale', async () => {
    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        async createSession() {
          throw createWorkbookAgentServiceError({
            code: 'WORKBOOK_AGENT_SESSION_NOT_FOUND',
            message: 'Workbook agent session not found',
            statusCode: 404,
            retryable: true,
          })
        },
      }),
    })

    try {
      const response = await app.inject({
        method: 'GET',
        url: '/v2/documents/doc-1/chat/threads/thr-1/events',
      })

      expect(response.statusCode).toBe(404)
      expect(response.json()).toEqual(
        expect.objectContaining({
          error: 'WORKBOOK_AGENT_SESSION_NOT_FOUND',
          message: 'Workbook agent session not found',
          retryable: true,
        }),
      )
    } finally {
      await app.close()
    }
  })

  it('serves minimal, non-cacheable health status when services are enabled', async () => {
    const getObservabilitySnapshot = vi.fn(() => {
      throw new Error('health checks must not collect detailed observability')
    })
    const { app } = createSyncServer({
      logger: false,
      zeroSyncService: createZeroSyncStub(),
      workbookAgentService: createWorkbookAgentServiceStub({ getObservabilitySnapshot }),
    })

    try {
      const response = await app.inject({
        method: 'GET',
        url: '/healthz',
      })

      expect(response.statusCode).toBe(200)
      expect(response.headers['cache-control']).toBe('no-store')
      expect(response.json()).toEqual({
        ok: true,
        service: 'bilig-app',
        zeroSync: true,
        web: false,
        workbookAgent: true,
      })
      expect(getObservabilitySnapshot).not.toHaveBeenCalled()
    } finally {
      await app.close()
    }
  })

  it('returns 503 from readyz until enabled persistence has finished initializing', async () => {
    const { app } = createSyncServer({
      logger: false,
      zeroSyncService: createZeroSyncStub({ isReady: () => false }),
    })

    try {
      const response = await app.inject({
        method: 'GET',
        url: '/readyz',
      })

      expect(response.statusCode).toBe(503)
      expect(response.headers['cache-control']).toBe('no-store')
      expect(response.json()).toEqual({
        ok: false,
        ready: false,
        service: 'bilig-app',
        zeroSync: true,
        web: false,
        workbookAgent: false,
      })
    } finally {
      await app.close()
    }
  })

  it('reports readyz when persistence is disabled or initialized', async () => {
    const { app } = createSyncServer({ logger: false })

    try {
      const response = await app.inject({
        method: 'GET',
        url: '/readyz',
      })

      expect(response.statusCode).toBe(200)
      expect(response.headers['cache-control']).toBe('no-store')
      expect(response.json()).toEqual({
        ok: true,
        ready: true,
        service: 'bilig-app',
        zeroSync: false,
        web: false,
        workbookAgent: false,
      })
    } finally {
      await app.close()
    }
  })

  it('exposes the workbook agent observability snapshot route', async () => {
    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        getObservabilitySnapshot() {
          return {
            enabled: true,
            generatedAtUnixMs: 99,
            featureFlags: {
              sharedThreadsEnabled: true,
              workflowRunnerEnabled: true,
              autoApplyLowRiskEnabled: true,
              formulaWorkflowFamilyEnabled: true,
              formattingWorkflowFamilyEnabled: true,
              importWorkflowFamilyEnabled: true,
              rollupWorkflowFamilyEnabled: true,
              structuralWorkflowFamilyEnabled: true,
              allowlistedUserCount: 0,
              allowlistedDocumentCount: 0,
            },
            sessions: {
              sessionCount: 0,
              subscriberThreadCount: 0,
              subscriberCount: 0,
              activeTurnCount: 0,
              runningWorkflowCount: 0,
              reviewQueueSessionCount: 0,
              sharedPendingReviewCount: 0,
            },
            pool: {
              slotCount: 0,
              boundThreadCount: 0,
              activeTurnCount: 0,
              queuedTurnCount: 0,
              maxClients: 4,
              maxConcurrentTurnsPerClient: 1,
              maxQueuedTurnsPerClient: 8,
            },
            counters: {
              turnBackpressureCount: 0,
              workflowStartedCount: 0,
              workflowCompletedCount: 0,
              workflowFailedCount: 0,
              workflowCancelledCount: 0,
              sharedReviewApprovedCount: 0,
              sharedReviewRejectedCount: 0,
              sharedRecommendationApprovedCount: 0,
              sharedRecommendationRejectedCount: 0,
            },
          }
        },
      }),
    })

    try {
      const response = await app.inject({
        method: 'GET',
        url: '/v2/agent/observability',
        headers: {
          cookie: 'bilig_session=test',
        },
      })

      expect(response.statusCode).toBe(200)
      expect(response.headers['cache-control']).toBe('no-store')
      expect(response.json()).toEqual(
        expect.objectContaining({
          enabled: true,
          generatedAtUnixMs: 99,
          pool: expect.objectContaining({
            maxClients: 4,
          }),
        }),
      )
    } finally {
      await app.close()
    }
  })
})
