import type { CodexInitializeResponse, CodexServerNotification, CodexTurn } from "@bilig/agent-api";
import type {
  CodexAppServerClientOptions,
  CodexAppServerTransport,
} from "./codex-app-server-client.js";

export class CodexAppServerPoolBackpressureError extends Error {
  readonly retryable = true;

  constructor(message: string) {
    super(message);
    this.name = "CodexAppServerPoolBackpressureError";
  }
}

export function isCodexAppServerPoolBackpressureError(
  value: unknown,
): value is CodexAppServerPoolBackpressureError {
  return value instanceof CodexAppServerPoolBackpressureError;
}

interface CodexAppServerPoolSlot {
  readonly id: number;
  readonly transport: CodexAppServerTransport;
  readonly threadIds: Set<string>;
  readonly unsubscribe: () => void;
  readonly initializeResponse: CodexInitializeResponse;
  activeTurnCount: number;
  readonly waiters: Array<() => void>;
  closed: boolean;
}

export interface CodexAppServerClientPoolOptions {
  readonly codexClientFactory: (options: CodexAppServerClientOptions) => CodexAppServerTransport;
  readonly clientOptions: CodexAppServerClientOptions;
  readonly maxClients?: number;
  readonly maxConcurrentTurnsPerClient?: number;
  readonly maxQueuedTurnsPerClient?: number;
}

export interface CodexAppServerClientPoolStats {
  readonly slotCount: number;
  readonly boundThreadCount: number;
  readonly activeTurnCount: number;
  readonly queuedTurnCount: number;
  readonly maxClients: number;
  readonly maxConcurrentTurnsPerClient: number;
  readonly maxQueuedTurnsPerClient: number;
}

export class CodexAppServerClientPool implements CodexAppServerTransport {
  private readonly codexClientFactory: CodexAppServerClientPoolOptions["codexClientFactory"];
  private readonly clientOptions: CodexAppServerClientPoolOptions["clientOptions"];
  private readonly maxClients: number;
  private readonly maxConcurrentTurnsPerClient: number;
  private readonly maxQueuedTurnsPerClient: number;
  private readonly slots = new Map<number, CodexAppServerPoolSlot>();
  private readonly threadToSlotId = new Map<string, number>();
  private readonly listeners = new Set<(notification: CodexServerNotification) => void>();
  private nextSlotId = 1;
  private slotCreationTask: Promise<void> = Promise.resolve();

  constructor(options: CodexAppServerClientPoolOptions) {
    this.codexClientFactory = options.codexClientFactory;
    this.clientOptions = options.clientOptions;
    this.maxClients = options.maxClients ?? 4;
    this.maxConcurrentTurnsPerClient = options.maxConcurrentTurnsPerClient ?? 1;
    this.maxQueuedTurnsPerClient = options.maxQueuedTurnsPerClient ?? 8;
  }

  async ensureReady(): Promise<CodexInitializeResponse> {
    const slot = await this.getSlotForNewThread();
    return slot.initializeResponse;
  }

  subscribe(listener: (notification: CodexServerNotification) => void): () => void {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  }

  async threadStart(
    input: Parameters<CodexAppServerTransport["threadStart"]>[0],
  ): Promise<Awaited<ReturnType<CodexAppServerTransport["threadStart"]>>> {
    const slot = await this.getSlotForNewThread();
    const thread = await slot.transport.threadStart(input);
    this.bindThread(slot, thread.id);
    return thread;
  }

  async threadResume(
    input: Parameters<CodexAppServerTransport["threadResume"]>[0],
  ): Promise<Awaited<ReturnType<CodexAppServerTransport["threadResume"]>>> {
    const slot =
      (await this.getSlotForThread(input.threadId)) ?? (await this.getSlotForNewThread());
    const thread = await slot.transport.threadResume(input);
    this.bindThread(slot, thread.id);
    return thread;
  }

  async turnStart(input: { threadId: string; prompt: string }): Promise<CodexTurn> {
    const slot = this.requireSlotForThread(input.threadId);
    const releasePermit = await this.acquireTurnPermit(slot, input.threadId);
    try {
      return await slot.transport.turnStart(input);
    } finally {
      releasePermit();
    }
  }

  async turnInterrupt(threadId: string): Promise<void> {
    const slot = this.requireSlotForThread(threadId);
    await slot.transport.turnInterrupt(threadId);
  }

  releaseThread(threadId: string): void {
    const slotId = this.threadToSlotId.get(threadId);
    if (slotId === undefined) {
      return;
    }
    this.threadToSlotId.delete(threadId);
    const slot = this.slots.get(slotId);
    if (!slot) {
      return;
    }
    slot.threadIds.delete(threadId);
    if (slot.threadIds.size === 0 && slot.activeTurnCount === 0 && slot.waiters.length === 0) {
      void this.closeSlot(slot);
    }
  }

  async close(): Promise<void> {
    const slots = [...this.slots.values()];
    this.slots.clear();
    this.threadToSlotId.clear();
    this.listeners.clear();
    await Promise.all(slots.map(async (slot) => await this.closeSlot(slot)));
  }

  getStats(): CodexAppServerClientPoolStats {
    const slots = [...this.slots.values()];
    return {
      slotCount: slots.length,
      boundThreadCount: [...this.threadToSlotId.keys()].length,
      activeTurnCount: slots.reduce((sum, slot) => sum + slot.activeTurnCount, 0),
      queuedTurnCount: slots.reduce((sum, slot) => sum + slot.waiters.length, 0),
      maxClients: this.maxClients,
      maxConcurrentTurnsPerClient: this.maxConcurrentTurnsPerClient,
      maxQueuedTurnsPerClient: this.maxQueuedTurnsPerClient,
    };
  }

  private bindThread(slot: CodexAppServerPoolSlot, threadId: string): void {
    const previousSlotId = this.threadToSlotId.get(threadId);
    if (previousSlotId !== undefined && previousSlotId !== slot.id) {
      const previousSlot = this.slots.get(previousSlotId);
      previousSlot?.threadIds.delete(threadId);
      if (
        previousSlot &&
        previousSlot.threadIds.size === 0 &&
        previousSlot.activeTurnCount === 0 &&
        previousSlot.waiters.length === 0
      ) {
        void this.closeSlot(previousSlot);
      }
    }
    this.threadToSlotId.set(threadId, slot.id);
    slot.threadIds.add(threadId);
  }

  private async getSlotForThread(threadId: string): Promise<CodexAppServerPoolSlot | null> {
    const slotId = this.threadToSlotId.get(threadId);
    return slotId === undefined ? null : (this.slots.get(slotId) ?? null);
  }

  private requireSlotForThread(threadId: string): CodexAppServerPoolSlot {
    const slotId = this.threadToSlotId.get(threadId);
    const slot = slotId === undefined ? null : (this.slots.get(slotId) ?? null);
    if (!slot) {
      throw new Error(`Codex thread ${threadId} is not assigned to a live pool slot.`);
    }
    return slot;
  }

  private async getSlotForNewThread(): Promise<CodexAppServerPoolSlot> {
    if (this.slots.size === 0) {
      return await this.createSlot();
    }
    if (this.slots.size < this.maxClients) {
      return await this.createSlot();
    }
    const leastLoadedSlot = [...this.slots.values()].toSorted(
      (left, right) => this.getSlotLoad(left) - this.getSlotLoad(right),
    )[0];
    if (!leastLoadedSlot) {
      throw new Error("Codex pool could not find an available slot.");
    }
    return leastLoadedSlot;
  }

  private getSlotLoad(slot: CodexAppServerPoolSlot): number {
    return slot.threadIds.size + slot.activeTurnCount * 4 + slot.waiters.length * 2;
  }

  private async createSlot(): Promise<CodexAppServerPoolSlot> {
    let createdSlot: CodexAppServerPoolSlot | null = null;
    this.slotCreationTask = this.slotCreationTask.then(async (): Promise<void> => {
      if (this.slots.size >= this.maxClients) {
        return undefined;
      }
      const transport = this.codexClientFactory(this.clientOptions);
      const initializeResponse = await transport.ensureReady();
      const slotId = this.nextSlotId;
      this.nextSlotId += 1;
      const unsubscribe = transport.subscribe((notification) => {
        this.listeners.forEach((listener) => {
          listener(notification);
        });
      });
      createdSlot = {
        id: slotId,
        transport,
        threadIds: new Set<string>(),
        unsubscribe,
        initializeResponse,
        activeTurnCount: 0,
        waiters: [],
        closed: false,
      };
      this.slots.set(slotId, createdSlot);
      return undefined;
    });
    await this.slotCreationTask;
    if (createdSlot) {
      return createdSlot;
    }
    const fallback = [...this.slots.values()].toSorted(
      (left, right) => this.getSlotLoad(left) - this.getSlotLoad(right),
    )[0];
    if (!fallback) {
      throw new Error("Codex pool could not create a slot.");
    }
    return fallback;
  }

  private async acquireTurnPermit(
    slot: CodexAppServerPoolSlot,
    threadId: string,
  ): Promise<() => void> {
    if (slot.activeTurnCount < this.maxConcurrentTurnsPerClient) {
      slot.activeTurnCount += 1;
      return () => {
        this.releaseTurnPermit(slot);
      };
    }
    if (slot.waiters.length >= this.maxQueuedTurnsPerClient) {
      throw new CodexAppServerPoolBackpressureError(
        `Workbook assistant is saturated for ${threadId}. Retry in a moment.`,
      );
    }
    await new Promise<void>((resolve) => {
      slot.waiters.push(() => {
        slot.activeTurnCount += 1;
        resolve();
      });
    });
    return () => {
      this.releaseTurnPermit(slot);
    };
  }

  private releaseTurnPermit(slot: CodexAppServerPoolSlot): void {
    slot.activeTurnCount = Math.max(0, slot.activeTurnCount - 1);
    const nextWaiter = slot.waiters.shift();
    if (nextWaiter) {
      nextWaiter();
      return;
    }
    if (slot.threadIds.size === 0) {
      void this.closeSlot(slot);
    }
  }

  private async closeSlot(slot: CodexAppServerPoolSlot): Promise<void> {
    if (slot.closed) {
      return;
    }
    slot.closed = true;
    this.slots.delete(slot.id);
    for (const threadId of slot.threadIds) {
      this.threadToSlotId.delete(threadId);
    }
    slot.threadIds.clear();
    slot.unsubscribe();
    await slot.transport.close();
  }
}
