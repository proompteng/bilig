import type { HelloFrame, ProtocolFrame } from '@bilig/binary-protocol'
import type { WorkbookBrowserSessionHost } from './browser-session-host.js'
import { routeWorkbookSyncFrame } from './sync-frame-router.js'

type SyncFrameOutput = ProtocolFrame | ProtocolFrame[]

export interface WorkbookSyncSessionHostOptions<Output extends SyncFrameOutput = SyncFrameOutput> {
  browserSessionHost: WorkbookBrowserSessionHost
  hello(frame: HelloFrame): Output | Promise<Output>
  appendBatch(frame: Extract<ProtocolFrame, { kind: 'appendBatch' }>): Output | Promise<Output>
  snapshotChunk(frame: Extract<ProtocolFrame, { kind: 'snapshotChunk' }>): Output | Promise<Output>
  heartbeat(frame: Extract<ProtocolFrame, { kind: 'heartbeat' }>): Output | Promise<Output>
  passthrough(frame: Extract<ProtocolFrame, { kind: 'cursorWatermark' | 'ack' | 'error' }>): Output | Promise<Output>
  unsupported(frame: ProtocolFrame): Output | Promise<Output>
}

export class WorkbookSyncSessionHost<Output extends SyncFrameOutput = SyncFrameOutput> {
  readonly snapshotAssemblies: WorkbookBrowserSessionHost['snapshotAssemblies']

  constructor(private readonly options: WorkbookSyncSessionHostOptions<Output>) {
    this.snapshotAssemblies = options.browserSessionHost.snapshotAssemblies
  }

  attachBrowser(documentId: string, subscriberId: string, send: (frame: ProtocolFrame) => void): () => void {
    return this.options.browserSessionHost.attachBrowser(documentId, subscriberId, send)
  }

  openBrowserSession(frame: HelloFrame): Promise<ProtocolFrame[]> {
    return this.options.browserSessionHost.openBrowserSession(frame)
  }

  broadcast(documentId: string, frame: ProtocolFrame): void {
    this.options.browserSessionHost.broadcast(documentId, frame)
  }

  listSubscriberIds(documentId: string): string[] {
    return this.options.browserSessionHost.listSubscriberIds(documentId)
  }

  async handleSyncFrame(frame: ProtocolFrame): Promise<Output> {
    return routeWorkbookSyncFrame(frame, {
      hello: (helloFrame) => this.options.hello(helloFrame),
      appendBatch: (appendFrame) => this.options.appendBatch(appendFrame),
      snapshotChunk: (snapshotFrame) => this.options.snapshotChunk(snapshotFrame),
      heartbeat: (heartbeatFrame) => this.options.heartbeat(heartbeatFrame),
      passthrough: (passthroughFrame) => this.options.passthrough(passthroughFrame),
      unsupported: (unsupportedFrame) => this.options.unsupported(unsupportedFrame),
    })
  }
}
