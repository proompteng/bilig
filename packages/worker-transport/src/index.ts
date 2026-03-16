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

interface SubscribeMessage {
  kind: "subscribe";
  id: number;
  channel: WorkerTransportChannel;
}

interface UnsubscribeMessage {
  kind: "unsubscribe";
  subscriptionId: number;
}

interface EventMessage {
  kind: "event";
  subscriptionId: number;
  channel: WorkerTransportChannel;
  payload: unknown;
}

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
  invoke<T = unknown>(method: string, ...args: unknown[]): Promise<T>;
  ready(): Promise<void>;
  subscribe(listener: (event: EngineEvent) => void): () => void;
  subscribeBatches(listener: (batch: EngineOpBatch) => void): () => void;
  dispose(): void;
}

export function createWorkerEngineHost(engine: WorkerTransportEngine, port: MessagePortLike): { dispose(): void } {
  const subscriptions = new Map<number, () => void>();

  const listener = (message: TransportMessage) => {

    if (message.kind === "request") {
      void handleRequest(engine, port, message);
      return;
    }

    if (message.kind === "subscribe") {
      const unsubscribe = subscribeChannel(engine, message.channel, (payload) => {
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
    if (typeof method !== "function") {
      throw new Error(`Unknown worker engine method: ${message.method}`);
    }
    const value = await (method as (...args: unknown[]) => unknown)(...message.args);
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

function subscribeChannel(
  engine: WorkerTransportEngine,
  channel: WorkerTransportChannel,
  listener: (payload: unknown) => void
): () => void {
  if (channel === "events") {
    if (!engine.subscribe) {
      throw new Error("Engine does not expose subscribe()");
    }
    return engine.subscribe(listener as (event: EngineEvent) => void);
  }

  if (!engine.subscribeBatches) {
    throw new Error("Engine does not expose subscribeBatches()");
  }
  return engine.subscribeBatches(listener as (batch: EngineOpBatch) => void);
}

export function createWorkerEngineClient(options: { port: MessagePortLike }): WorkerEngineClient {
  const { port } = options;
  let nextId = 1;
  const pending = new Map<number, { resolve: (value: unknown) => void; reject: (error: Error) => void }>();
  const listeners = new Map<number, { channel: WorkerTransportChannel; callback: (payload: unknown) => void }>();

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
      listeners.get(message.subscriptionId)?.callback(message.payload);
    }
  };

  const detach = attachMessageListener(port, onMessage);
  port.start?.();

  function invoke<T>(method: string, ...args: unknown[]): Promise<T> {
    const id = nextId++;
    return new Promise<T>((resolve, reject) => {
      pending.set(id, {
        resolve: (value) => resolve(value as T),
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

  function subscribe(channel: WorkerTransportChannel, callback: (payload: unknown) => void): () => void {
    const id = nextId++;
    listeners.set(id, { channel, callback });
    port.postMessage({ kind: "subscribe", id, channel });
    return () => {
      listeners.delete(id);
      port.postMessage({ kind: "unsubscribe", subscriptionId: id });
    };
  }

  return {
    invoke,
    ready() {
      return invoke<void>("ready");
    },
    subscribe(listener) {
      return subscribe("events", listener as (payload: unknown) => void);
    },
    subscribeBatches(listener) {
      return subscribe("batches", listener as (payload: unknown) => void);
    },
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
