import type { Writable } from "node:stream";

import { decodeStdioMessages, encodeStdioMessage, type AgentEvent, type AgentFrame } from "@bilig/agent-api";

export interface AgentFrameHandler {
  handleAgentFrame(frame: AgentFrame): Promise<AgentFrame>;
}

export interface AgentEventStreamHandler extends AgentFrameHandler {
  subscribeAgentEvents(listener: (event: AgentEvent) => void): () => void;
}

export interface StdioAgentLoopOptions {
  handler: AgentFrameHandler;
  input?: NodeJS.ReadStream;
  output?: Writable;
  onError?: (error: Error) => void;
}

function appendBytes(left: Uint8Array, right: Uint8Array): Uint8Array {
  const joined = new Uint8Array(left.byteLength + right.byteLength);
  joined.set(left, 0);
  joined.set(right, left.byteLength);
  return joined;
}

function toBytes(chunk: Buffer | string): Uint8Array {
  return typeof chunk === "string" ? new Uint8Array(Buffer.from(chunk)) : new Uint8Array(chunk);
}

function isAgentEventStreamHandler(handler: AgentFrameHandler): handler is AgentEventStreamHandler {
  return "subscribeAgentEvents" in handler && typeof handler.subscribeAgentEvents === "function";
}

export function attachStdioAgentLoop(options: StdioAgentLoopOptions): { dispose(): void } {
  const input = options.input ?? process.stdin;
  const output = options.output ?? process.stdout;
  const onError = options.onError ?? ((error) => {
    console.error(error);
  });
  let buffer: Uint8Array = new Uint8Array(0);
  let queue = Promise.resolve();
  let disposed = false;

  const enqueueFrame = (frame: AgentFrame) => {
    queue = queue
      .then(() => {
        if (disposed) {
          return undefined;
        }
        output.write(Buffer.from(encodeStdioMessage(frame)));
        return undefined;
      })
      .catch((error: unknown) => {
        onError(error instanceof Error ? error : new Error(String(error)));
      });
  };

  const detachEvents = isAgentEventStreamHandler(options.handler)
    ? options.handler.subscribeAgentEvents((event) => {
        enqueueFrame({ kind: "event", event });
      })
    : () => {};

  const handleData = (chunk: Buffer | string) => {
    try {
      buffer = appendBytes(buffer, toBytes(chunk));
      const { frames, remainder } = decodeStdioMessages(buffer);
      buffer = remainder;
      if (frames.length === 0) {
        return;
      }
      queue = queue
        .then(async () => {
          for (const frame of frames) {
            const response = await options.handler.handleAgentFrame(frame);
            if (disposed) {
              return undefined;
            }
            output.write(Buffer.from(encodeStdioMessage(response)));
          }
          return undefined;
        })
        .catch((error: unknown) => {
          onError(error instanceof Error ? error : new Error(String(error)));
        });
    } catch (error) {
      onError(error instanceof Error ? error : new Error(String(error)));
    }
  };

  input.on("data", handleData);
  input.resume();

  return {
    dispose() {
      disposed = true;
      input.off("data", handleData);
      detachEvents();
    }
  };
}
