import type { EngineEvent } from "@bilig/protocol";

import type { EngineOpBatch } from "@bilig/crdt";

export type WorkerTransportChannel = "events" | "batches";

interface RequestMessage {
  kind: "request";
  id: number;
  method: string;
  args: unknown[];
}

interface ResponseMessage {
  kind: "response";
  id: number;
  ok: boolean;
  value?: unknown;
  error?: string;
}

type SubscribeMessage =
  | { kind: "subscribe"; id: number; channel: "events" }
  | { kind: "subscribe"; id: number; channel: "batches" };

interface UnsubscribeMessage {
  kind: "unsubscribe";
  subscriptionId: number;
}

type EventMessage =
  | { kind: "event"; subscriptionId: number; channel: "events"; payload: EngineEvent }
  | { kind: "event"; subscriptionId: number; channel: "batches"; payload: EngineOpBatch };

type TransportMessage = RequestMessage | ResponseMessage | SubscribeMessage | UnsubscribeMessage | EventMessage;

export interface MessagePortLike {
  postMessage(message: TransportMessage): void;
  addEventListener?(type: "message", listener: (event: MessageEvent<TransportMessage>) => void): void;
  removeEventListener?(type: "message", listener: (event: MessageEvent<TransportMessage>) => void): void;
  on?(type: "message", listener: (message: TransportMessage) => void): void;
  off?(type: "message", listener: (message: TransportMessage) => void): void;
  start?: () => void;
}

export interface WorkerTransportEngine {
  ready?: () => Promise<void>;
  subscribe?: (listener: (event: EngineEvent) => void) => () => void;
  subscribeBatches?: (listener: (batch: EngineOpBatch) => void) => () => void;
  [method: string]: unknown;
}

export interface WorkerEngineClient {
  invoke(method: string, ...args: unknown[]): Promise<unknown>;
  ready(): Promise<void>;
  subscribe(listener: (event: EngineEvent) => void): () => void;
  subscribeBatches(listener: (batch: EngineOpBatch) => void): () => void;
  dispose(): void;
}

function isCallableMethod(value: unknown): value is (...args: unknown[]) => unknown {
  return typeof value === "function";
}

export function createWorkerEngineHost(engine: WorkerTransportEngine, port: MessagePortLike): { dispose(): void } {
  const subscriptions = new Map<number, () => void>();

  const listener = (message: TransportMessage) => {

    if (message.kind === "request") {
      void handleRequest(engine, port, message);
      return;
    }

    if (message.kind === "subscribe") {
      const unsubscribe = message.channel === "events"
        ? subscribeEventChannel(engine, (payload: EngineEvent) => {
          port.postMessage({
            kind: "event",
            subscriptionId: message.id,
            channel: message.channel,
            payload
          });
        })
        : subscribeBatchChannel(engine, (payload: EngineOpBatch) => {
          port.postMessage({
            kind: "event",
            subscriptionId: message.id,
            channel: message.channel,
            payload
          });
        });
      subscriptions.set(message.id, unsubscribe);
      port.postMessage({
        kind: "response",
        id: message.id,
        ok: true
      });
      return;
    }

    if (message.kind === "unsubscribe") {
      subscriptions.get(message.subscriptionId)?.();
      subscriptions.delete(message.subscriptionId);
    }
  };

  const detach = attachMessageListener(port, listener);
  port.start?.();

  return {
    dispose() {
      subscriptions.forEach((unsubscribe) => unsubscribe());
      subscriptions.clear();
      detach();
    }
  };
}

async function handleRequest(
  engine: WorkerTransportEngine,
  port: MessagePortLike,
  message: RequestMessage
): Promise<void> {
  try {
    const method = engine[message.method];
    if (!isCallableMethod(method)) {
      throw new Error(`Unknown worker engine method: ${message.method}`);
    }
    const value = await method(...message.args);
    port.postMessage({
      kind: "response",
      id: message.id,
      ok: true,
      value
    });
  } catch (error) {
    port.postMessage({
      kind: "response",
      id: message.id,
      ok: false,
      error: error instanceof Error ? error.message : String(error)
    });
  }
}

function subscribeEventChannel(engine: WorkerTransportEngine, listener: (payload: EngineEvent) => void): () => void {
  if (!engine.subscribe) {
    throw new Error("Engine does not expose subscribe()");
  }
  return engine.subscribe(listener);
}

function subscribeBatchChannel(engine: WorkerTransportEngine, listener: (payload: EngineOpBatch) => void): () => void {
  if (!engine.subscribeBatches) {
    throw new Error("Engine does not expose subscribeBatches()");
  }
  return engine.subscribeBatches(listener);
}

export function createWorkerEngineClient(options: { port: MessagePortLike }): WorkerEngineClient {
  const { port } = options;
  let nextId = 1;
  const pending = new Map<number, { resolve: (value: unknown) => void; reject: (error: Error) => void }>();
  const listeners = new Map<
    number,
    | { channel: "events"; callback: (payload: EngineEvent) => void }
    | { channel: "batches"; callback: (payload: EngineOpBatch) => void }
  >();

  const onMessage = (message: TransportMessage) => {
    if (message.kind === "response") {
      const promise = pending.get(message.id);
      if (!promise) {
        return;
      }
      pending.delete(message.id);
      if (message.ok) {
        promise.resolve(message.value);
      } else {
        promise.reject(new Error(message.error ?? "Worker transport failed"));
      }
      return;
    }

    if (message.kind === "event") {
      const listener = listeners.get(message.subscriptionId);
      if (!listener || listener.channel !== message.channel) {
        return;
      }
      if (message.channel === "events" && listener.channel === "events") {
        listener.callback(message.payload);
        return;
      }
      if (message.channel === "batches" && listener.channel === "batches") {
        listener.callback(message.payload);
      }
    }
  };

  const detach = attachMessageListener(port, onMessage);
  port.start?.();

  function invoke(method: string, ...args: unknown[]): Promise<unknown> {
    const id = nextId++;
    return new Promise<unknown>((resolve, reject) => {
      pending.set(id, {
        resolve,
        reject
      });
      port.postMessage({
        kind: "request",
        id,
        method,
        args
      });
    });
  }

  function subscribeEvents(callback: (payload: EngineEvent) => void): () => void {
    const id = nextId++;
    listeners.set(id, { channel: "events", callback });
    port.postMessage({ kind: "subscribe", id, channel: "events" });
    return () => {
      listeners.delete(id);
      port.postMessage({ kind: "unsubscribe", subscriptionId: id });
    };
  }

  function subscribeBatches(callback: (payload: EngineOpBatch) => void): () => void {
    const id = nextId++;
    listeners.set(id, { channel: "batches", callback });
    port.postMessage({ kind: "subscribe", id, channel: "batches" });
    return () => {
      listeners.delete(id);
      port.postMessage({ kind: "unsubscribe", subscriptionId: id });
    };
  }

  return {
    invoke,
    ready() {
      return invoke("ready").then(() => undefined);
    },
    subscribe: subscribeEvents,
    subscribeBatches,
    dispose() {
      pending.forEach((entry) => entry.reject(new Error("Worker engine client disposed")));
      pending.clear();
      listeners.clear();
      detach();
    }
  };
}

function attachMessageListener(
  port: MessagePortLike,
  listener: (message: TransportMessage) => void
): () => void {
  if (typeof port.addEventListener === "function" && typeof port.removeEventListener === "function") {
    const wrapped = (event: MessageEvent<TransportMessage>) => {
      listener(event.data);
    };
    port.addEventListener("message", wrapped);
    return () => port.removeEventListener?.("message", wrapped);
  }

  if (typeof port.on === "function" && typeof port.off === "function") {
    port.on("message", listener);
    return () => port.off?.("message", listener);
  }

  throw new Error("Unsupported message port implementation");
}
