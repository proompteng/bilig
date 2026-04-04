import type { HelloFrame, ProtocolFrame } from "@bilig/binary-protocol";
import { openWorkbookBrowserSession } from "./browser-session-shared.js";
import type { SnapshotReplayState } from "./browser-sync-replay.js";
import {
  attachBrowserSubscriber,
  broadcastToBrowsers,
  listBrowserSubscriberIds,
  type BrowserSubscriberRegistry,
  type SnapshotAssemblyRegistry,
} from "./session-shared.js";

type AppendBatchFrame = Extract<ProtocolFrame, { kind: "appendBatch" }>;

export interface WorkbookBrowserSessionHostOptions {
  register?(frame: HelloFrame): void | Promise<void>;
  latestCursor(documentId: string): number | Promise<number>;
  latestSnapshot(documentId: string): SnapshotReplayState | Promise<SnapshotReplayState>;
  listMissedFrames(
    documentId: string,
    cursorFloor: number,
  ): AppendBatchFrame[] | Promise<AppendBatchFrame[]>;
}

export class WorkbookBrowserSessionHost {
  readonly snapshotAssemblies: SnapshotAssemblyRegistry = new Map();
  private readonly browserSubscribers: BrowserSubscriberRegistry = new Map();

  constructor(private readonly options: WorkbookBrowserSessionHostOptions) {}

  attachBrowser(
    documentId: string,
    subscriberId: string,
    send: (frame: ProtocolFrame) => void,
  ): () => void {
    return attachBrowserSubscriber(this.browserSubscribers, documentId, subscriberId, send);
  }

  async openBrowserSession(frame: HelloFrame): Promise<ProtocolFrame[]> {
    return openWorkbookBrowserSession(frame, {
      ...(this.options.register
        ? {
            register: (helloFrame: HelloFrame) => this.options.register?.(helloFrame),
          }
        : {}),
      latestCursor: this.options.latestCursor(frame.documentId),
      latestSnapshot: this.options.latestSnapshot(frame.documentId),
      listMissedFrames: (cursorFloor) =>
        this.options.listMissedFrames(frame.documentId, cursorFloor),
    });
  }

  broadcast(documentId: string, frame: ProtocolFrame): void {
    broadcastToBrowsers(this.browserSubscribers, documentId, frame);
  }

  listSubscriberIds(documentId: string): string[] {
    return listBrowserSubscriberIds(this.browserSubscribers, documentId);
  }
}
