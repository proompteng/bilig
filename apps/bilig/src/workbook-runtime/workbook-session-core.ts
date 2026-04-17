import type { HelloFrame, ProtocolFrame } from '@bilig/binary-protocol'
import type { AgentFrame, AgentRequest, AgentResponse, LoadWorkbookFileRequest } from '@bilig/agent-api'
import type { AgentFrameContext, WorksheetAgentRequest } from './agent-routing.js'
import { handleWorkbookAgentFrame } from './workbook-session-shared.js'
import type { WorkbookSyncSessionHost } from './workbook-sync-session-host.js'

type SyncFrameOutput = ProtocolFrame | ProtocolFrame[]

export interface WorkbookSessionCoreOptions<SyncOutput extends SyncFrameOutput> {
  syncSessionHost: WorkbookSyncSessionHost<SyncOutput>
  invalidFrameMessage: string
  errorCode: string
  loadWorkbookFile: (
    request: LoadWorkbookFileRequest,
    context: AgentFrameContext,
  ) => AgentResponse | AgentFrame | Promise<AgentResponse | AgentFrame>
  openWorkbookSession: (
    request: Extract<AgentRequest, { kind: 'openWorkbookSession' }>,
  ) => string | AgentResponse | AgentFrame | Promise<string | AgentResponse | AgentFrame>
  closeWorkbookSession: (
    request: Extract<AgentRequest, { kind: 'closeWorkbookSession' }>,
  ) => void | AgentResponse | AgentFrame | Promise<void | AgentResponse | AgentFrame>
  getMetrics: (request: Extract<AgentRequest, { kind: 'getMetrics' }>) => AgentResponse | AgentFrame | Promise<AgentResponse | AgentFrame>
  handleWorksheetRequest?: (
    frame: Extract<AgentFrame, { kind: 'request' }>,
    request: WorksheetAgentRequest,
  ) => AgentResponse | AgentFrame | Promise<AgentResponse | AgentFrame>
}

export class WorkbookSessionCore<SyncOutput extends SyncFrameOutput> {
  constructor(private readonly options: WorkbookSessionCoreOptions<SyncOutput>) {}

  attachBrowser(documentId: string, subscriberId: string, send: (frame: ProtocolFrame) => void): () => void {
    return this.options.syncSessionHost.attachBrowser(documentId, subscriberId, send)
  }

  openBrowserSession(frame: HelloFrame): Promise<ProtocolFrame[]> {
    return this.options.syncSessionHost.openBrowserSession(frame)
  }

  handleSyncFrame(frame: ProtocolFrame): Promise<SyncOutput> {
    return this.options.syncSessionHost.handleSyncFrame(frame)
  }

  handleAgentFrame(frame: AgentFrame, context: AgentFrameContext = {}): Promise<AgentFrame> {
    return handleWorkbookAgentFrame(frame, context, {
      invalidFrameMessage: this.options.invalidFrameMessage,
      errorCode: this.options.errorCode,
      loadWorkbookFile: this.options.loadWorkbookFile,
      openWorkbookSession: this.options.openWorkbookSession,
      closeWorkbookSession: this.options.closeWorkbookSession,
      getMetrics: this.options.getMetrics,
      ...(this.options.handleWorksheetRequest
        ? {
            handleWorksheetRequest: this.options.handleWorksheetRequest,
          }
        : {}),
    })
  }

  broadcast(documentId: string, frame: ProtocolFrame): void {
    this.options.syncSessionHost.broadcast(documentId, frame)
  }

  listSubscriberIds(documentId: string): string[] {
    return this.options.syncSessionHost.listSubscriberIds(documentId)
  }

  get snapshotAssemblies() {
    return this.options.syncSessionHost.snapshotAssemblies
  }
}
