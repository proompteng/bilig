import type { EngineEvent } from "@bilig/protocol";

import type { EngineOpBatch } from "@bilig/crdt";
import type { ViewportPatchSubscription } from "./viewport-patch.js";

export type WorkerTransportChannel = "events" | "batches" | "viewportPatches";

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
  | { kind: "subscribe"; id: number; channel: "events"; args?: [] }
  | { kind: "subscribe"; id: number; channel: "batches"; args?: [] }
  | {
      kind: "subscribe";
      id: number;
      channel: "viewportPatches";
      args: [ViewportPatchSubscription];
    };

interface UnsubscribeMessage {
  kind: "unsubscribe";
  subscriptionId: number;
}

type EventMessage =
  | { kind: "event"; subscriptionId: number; channel: "events"; payload: EngineEvent }
  | { kind: "event"; subscriptionId: number; channel: "batches"; payload: EngineOpBatch }
  | { kind: "event"; subscriptionId: number; channel: "viewportPatches"; payload: Uint8Array };

type TransportMessage =
  | RequestMessage
  | ResponseMessage
  | SubscribeMessage
  | UnsubscribeMessage
  | EventMessage;

export interface MessagePortLike {
  postMessage(message: TransportMessage): void;
  addEventListener?(
    type: "message",
    listener: (event: MessageEvent<TransportMessage>) => void,
  ): void;
  removeEventListener?(
    type: "message",
    listener: (event: MessageEvent<TransportMessage>) => void,
  ): void;
  on?(type: "message", listener: (message: TransportMessage) => void): void;
  off?(type: "message", listener: (message: TransportMessage) => void): void;
  start?: () => void;
}

export interface WorkerTransportEngine {
  ready?: () => Promise<void>;
  subscribe?: (listener: (event: EngineEvent) => void) => () => void;
  subscribeBatches?: (listener: (batch: EngineOpBatch) => void) => () => void;
  subscribeViewportPatches?: (
    subscription: ViewportPatchSubscription,
    listener: (patch: Uint8Array) => void,
  ) => () => void;
  [method: string]: unknown;
}

export interface WorkerEngineClient {
  invoke(method: string, ...args: unknown[]): Promise<unknown>;
  ready(): Promise<void>;
  subscribe(listener: (event: EngineEvent) => void): () => void;
  subscribeBatches(listener: (batch: EngineOpBatch) => void): () => void;
  subscribeViewportPatches(
    subscription: ViewportPatchSubscription,
    listener: (patch: Uint8Array) => void,
  ): () => void;
  dispose(): void;
}

function isCallableMethod(value: unknown): value is (...args: unknown[]) => unknown {
  return typeof value === "function";
}

export function createWorkerEngineHost(
  engine: WorkerTransportEngine,
  port: MessagePortLike,
): { dispose(): void } {
  const subscriptions = new Map<number, () => void>();

  const listener = (message: TransportMessage) => {
    if (message.kind === "request") {
      void handleRequest(engine, port, message);
      return;
    }

    if (message.kind === "subscribe") {
      const unsubscribe = createChannelSubscription(engine, port, message);
      subscriptions.set(message.id, unsubscribe);
      port.postMessage({
        kind: "response",
        id: message.id,
        ok: true,
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
    },
  };
}

async function handleRequest(
  engine: WorkerTransportEngine,
  port: MessagePortLike,
  message: RequestMessage,
): Promise<void> {
  try {
    const method = engine[message.method];
    if (!isCallableMethod(method)) {
      throw new Error(`Unknown worker engine method: ${message.method}`);
    }
    const value = await Reflect.apply(method, engine, message.args);
    port.postMessage({
      kind: "response",
      id: message.id,
      ok: true,
      value,
    });
  } catch (error) {
    port.postMessage({
      kind: "response",
      id: message.id,
      ok: false,
      error: error instanceof Error ? error.message : String(error),
    });
  }
}

function subscribeEventChannel(
  engine: WorkerTransportEngine,
  listener: (payload: EngineEvent) => void,
): () => void {
  if (!engine.subscribe) {
    throw new Error("Engine does not expose subscribe()");
  }
  return engine.subscribe(listener);
}

function subscribeBatchChannel(
  engine: WorkerTransportEngine,
  listener: (payload: EngineOpBatch) => void,
): () => void {
  if (!engine.subscribeBatches) {
    throw new Error("Engine does not expose subscribeBatches()");
  }
  return engine.subscribeBatches(listener);
}

function subscribeViewportPatchChannel(
  engine: WorkerTransportEngine,
  subscription: ViewportPatchSubscription,
  listener: (payload: Uint8Array) => void,
): () => void {
  if (!engine.subscribeViewportPatches) {
    throw new Error("Engine does not expose subscribeViewportPatches()");
  }
  return engine.subscribeViewportPatches(subscription, listener);
}

function createChannelSubscription(
  engine: WorkerTransportEngine,
  port: MessagePortLike,
  message: SubscribeMessage,
): () => void {
  switch (message.channel) {
    case "events":
      return subscribeEventChannel(engine, (payload: EngineEvent) => {
        port.postMessage({
          kind: "event",
          subscriptionId: message.id,
          channel: "events",
          payload,
        });
      });
    case "batches":
      return subscribeBatchChannel(engine, (payload: EngineOpBatch) => {
        port.postMessage({
          kind: "event",
          subscriptionId: message.id,
          channel: "batches",
          payload,
        });
      });
    case "viewportPatches":
      return subscribeViewportPatchChannel(engine, message.args[0], (payload: Uint8Array) => {
        port.postMessage({
          kind: "event",
          subscriptionId: message.id,
          channel: "viewportPatches",
          payload,
        });
      });
  }
}

export function createWorkerEngineClient(options: { port: MessagePortLike }): WorkerEngineClient {
  const { port } = options;
  let nextId = 1;
  const pending = new Map<
    number,
    { resolve: (value: unknown) => void; reject: (error: Error) => void }
  >();
  const listeners = new Map<
    number,
    | { channel: "events"; callback: (payload: EngineEvent) => void }
    | { channel: "batches"; callback: (payload: EngineOpBatch) => void }
    | { channel: "viewportPatches"; callback: (payload: Uint8Array) => void }
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
        return;
      }
      if (message.channel === "viewportPatches" && listener.channel === "viewportPatches") {
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
        reject,
      });
      port.postMessage({
        kind: "request",
        id,
        method,
        args,
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
    port.postMessage({ kind: "subscribe", id, channel: "batches", args: [] });
    return () => {
      listeners.delete(id);
      port.postMessage({ kind: "unsubscribe", subscriptionId: id });
    };
  }

  function subscribeViewportPatches(
    subscription: ViewportPatchSubscription,
    callback: (payload: Uint8Array) => void,
  ): () => void {
    const id = nextId++;
    listeners.set(id, { channel: "viewportPatches", callback });
    port.postMessage({ kind: "subscribe", id, channel: "viewportPatches", args: [subscription] });
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
    subscribeViewportPatches,
    dispose() {
      pending.forEach((entry) => entry.reject(new Error("Worker engine client disposed")));
      pending.clear();
      listeners.clear();
      detach();
    },
  };
}

function attachMessageListener(
  port: MessagePortLike,
  listener: (message: TransportMessage) => void,
): () => void {
  if (
    typeof port.addEventListener === "function" &&
    typeof port.removeEventListener === "function"
  ) {
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

export * from "./viewport-patch.js";
export * from "./websocket-sync-client.js";
