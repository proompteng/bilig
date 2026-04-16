import { Effect } from 'effect'
import type { EngineOpBatch } from '@bilig/workbook-domain'
import { shouldApplyBatch } from '../../replica-state.js'
import type { EngineRuntimeState, EngineSyncClient } from '../runtime-state.js'
import { EngineSyncError } from '../errors.js'

export interface EngineReplicaSyncService {
  readonly connectClient: (client: EngineSyncClient) => Effect.Effect<void, EngineSyncError>
  readonly disconnectClient: () => Effect.Effect<void, EngineSyncError>
  readonly applyRemoteBatch: (batch: EngineOpBatch) => Effect.Effect<boolean, EngineSyncError>
}

export function createEngineReplicaSyncService(args: {
  readonly state: Pick<
    EngineRuntimeState,
    'replicaState' | 'getSyncState' | 'setSyncState' | 'getSyncClientConnection' | 'setSyncClientConnection'
  >
  readonly applyRemoteBatchNow: (batch: EngineOpBatch) => void
  readonly applyRemoteSnapshot: (snapshot: import('@bilig/protocol').WorkbookSnapshot) => void
}): EngineReplicaSyncService {
  return {
    connectClient(client) {
      return Effect.tryPromise({
        try: async () => {
          const existing = args.state.getSyncClientConnection()
          args.state.setSyncClientConnection(null)
          if (existing) {
            await existing.disconnect()
          }
          args.state.setSyncState('syncing')
          const connection = await client.connect({
            applyRemoteBatch: (batch) => Effect.runSync(this.applyRemoteBatch(batch)),
            applyRemoteSnapshot: (snapshot) => {
              args.applyRemoteSnapshot(snapshot)
            },
            setState: (state) => {
              args.state.setSyncState(state)
            },
          })
          args.state.setSyncClientConnection(connection)
          if (args.state.getSyncState() === 'syncing') {
            args.state.setSyncState('live')
          }
        },
        catch: (cause) =>
          new EngineSyncError({
            message: 'Failed to connect sync client',
            cause,
          }),
      })
    },
    disconnectClient() {
      return Effect.tryPromise({
        try: async () => {
          const connection = args.state.getSyncClientConnection()
          args.state.setSyncClientConnection(null)
          if (connection) {
            await connection.disconnect()
          }
          args.state.setSyncState('local-only')
        },
        catch: (cause) =>
          new EngineSyncError({
            message: 'Failed to disconnect sync client',
            cause,
          }),
      })
    },
    applyRemoteBatch(batch) {
      return Effect.try({
        try: () => {
          if (!shouldApplyBatch(args.state.replicaState, batch)) {
            return false
          }
          args.applyRemoteBatchNow(batch)
          return true
        },
        catch: (cause) =>
          new EngineSyncError({
            message: 'Failed to apply remote batch',
            cause,
          }),
      })
    },
  }
}
