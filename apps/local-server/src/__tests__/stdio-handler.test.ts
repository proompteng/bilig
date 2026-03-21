import { PassThrough } from "node:stream";
import { describe, expect, it } from "vitest";

import { decodeStdioMessages, encodeStdioMessage, type AgentFrame } from "@bilig/agent-api";
import { ValueTag } from "@bilig/protocol";

import { LocalWorkbookSessionManager } from "../local-workbook-session-manager.js";
import { attachStdioAgentLoop } from "../stdio-handler.js";

function createFrameReader(stream: PassThrough) {
  let buffer = new Uint8Array(0);
  const queued: AgentFrame[] = [];

  return async function readSingleFrame() {
    if (queued.length > 0) {
      return queued.shift()!;
    }

    return await new Promise<AgentFrame>((resolve, reject) => {
      const onData = (chunk: Buffer) => {
        const joined = new Uint8Array(buffer.byteLength + chunk.byteLength);
        joined.set(buffer, 0);
        joined.set(chunk, buffer.byteLength);
        const decoded = decodeStdioMessages(joined);
        buffer = decoded.remainder;
        if (decoded.frames.length === 0) {
          return;
        }
        queued.push(...decoded.frames);
        stream.off("data", onData);
        const nextFrame = queued.shift();
        if (!nextFrame) {
          reject(new Error("Expected at least one decoded stdio frame"));
          return;
        }
        resolve(nextFrame);
      };
      stream.on("data", onData);
      stream.once("error", reject);
    });
  };
}

describe("stdio handler", () => {
  it("executes workbook mutations through length-prefixed stdio frames", async () => {
    const input = new PassThrough();
    const output = new PassThrough();
    const readSingleFrame = createFrameReader(output);
    const handler = attachStdioAgentLoop({
      handler: new LocalWorkbookSessionManager(),
      input,
      output
    });

    try {
      input.write(Buffer.from(encodeStdioMessage({
        kind: "request",
        request: {
          kind: "openWorkbookSession",
          id: "open-stdio",
          documentId: "stdio-doc",
          replicaId: "stdio-agent"
        }
      })));
      const openFrame = await readSingleFrame();
      expect(openFrame).toMatchObject({
        kind: "response",
        response: {
          kind: "ok",
          sessionId: "stdio-doc:stdio-agent"
        }
      });

      input.write(Buffer.from(encodeStdioMessage({
        kind: "request",
        request: {
          kind: "writeRange",
          id: "write-stdio",
          sessionId: "stdio-doc:stdio-agent",
          range: {
            sheetName: "Sheet1",
            startAddress: "A1",
            endAddress: "A1"
          },
          values: [[42]]
        }
      })));
      await readSingleFrame();

      input.write(Buffer.from(encodeStdioMessage({
        kind: "request",
        request: {
          kind: "readRange",
          id: "read-stdio",
          sessionId: "stdio-doc:stdio-agent",
          range: {
            sheetName: "Sheet1",
            startAddress: "A1",
            endAddress: "A1"
          }
        }
      })));
      const readFrame = await readSingleFrame();
      expect(readFrame).toMatchObject({
        kind: "response",
        response: {
          kind: "rangeValues"
        }
      });
      if (readFrame.kind !== "response" || readFrame.response.kind !== "rangeValues") {
        throw new Error("Expected rangeValues response");
      }
      expect(readFrame.response.values[0]?.[0]).toEqual({
        tag: ValueTag.Number,
        value: 42
      });
    } finally {
      handler.dispose();
      input.destroy();
      output.destroy();
    }
  });

  it("streams rangeChanged events over stdio without polling", async () => {
    const input = new PassThrough();
    const output = new PassThrough();
    const readSingleFrame = createFrameReader(output);
    const manager = new LocalWorkbookSessionManager();
    const handler = attachStdioAgentLoop({
      handler: manager,
      input,
      output
    });

    try {
      input.write(Buffer.from(encodeStdioMessage({
        kind: "request",
        request: {
          kind: "openWorkbookSession",
          id: "open-stream",
          documentId: "stream-doc",
          replicaId: "stream-agent"
        }
      })));
      await readSingleFrame();

      input.write(Buffer.from(encodeStdioMessage({
        kind: "request",
        request: {
          kind: "subscribeRange",
          id: "sub-stream",
          sessionId: "stream-doc:stream-agent",
          subscriptionId: "sub-1",
          range: {
            sheetName: "Sheet1",
            startAddress: "A1",
            endAddress: "B2"
          }
        }
      })));
      const subscribeFrame = await readSingleFrame();
      expect(subscribeFrame).toMatchObject({
        kind: "response",
        response: {
          kind: "ok",
          value: {
            subscriptionId: "sub-1"
          }
        }
      });

      await manager.handleSyncFrame({
        kind: "appendBatch",
        documentId: "stream-doc",
        cursor: 0,
        batch: {
          id: "browser-stream:1",
          replicaId: "browser-stream",
          clock: { counter: 1 },
          ops: [{ kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 9 }]
        }
      });

      const eventFrame = await readSingleFrame();
      expect(eventFrame).toEqual({
        kind: "event",
        event: {
          kind: "rangeChanged",
          subscriptionId: "sub-1",
          range: {
            sheetName: "Sheet1",
            startAddress: "A1",
            endAddress: "B2"
          },
          changedAddresses: ["A1", "B1", "A2", "B2"]
        }
      });

      input.write(Buffer.from(encodeStdioMessage({
        kind: "request",
        request: {
          kind: "unsubscribe",
          id: "unsub-stream",
          sessionId: "stream-doc:stream-agent",
          subscriptionId: "sub-1"
        }
      })));
      const unsubscribeFrame = await readSingleFrame();
      expect(unsubscribeFrame).toMatchObject({
        kind: "response",
        response: {
          kind: "ok",
          id: "unsub-stream"
        }
      });

      await manager.handleSyncFrame({
        kind: "appendBatch",
        documentId: "stream-doc",
        cursor: 0,
        batch: {
          id: "browser-stream:2",
          replicaId: "browser-stream",
          clock: { counter: 2 },
          ops: [{ kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 10 }]
        }
      });

      const pendingFrames = decodeStdioMessages(output.read() ?? Buffer.alloc(0)).frames;
      expect(pendingFrames).toHaveLength(0);
    } finally {
      handler.dispose();
      input.destroy();
      output.destroy();
    }
  });
});
