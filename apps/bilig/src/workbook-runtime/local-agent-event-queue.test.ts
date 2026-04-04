import { describe, expect, it, vi } from "vitest";
import type { AgentEvent } from "@bilig/agent-api";
import {
  flushQueuedLocalAgentEvents,
  queueLocalAgentEvent,
  removeQueuedSubscriptionEvents,
} from "./local-agent-event-queue.js";

function createEvent(subscriptionId: string): AgentEvent {
  return {
    kind: "rangeChanged",
    subscriptionId,
    range: {
      sheetName: "Sheet1",
      startAddress: "A1",
      endAddress: "A2",
    },
    changedAddresses: ["A1"],
  };
}

describe("local-agent-event-queue", () => {
  it("schedules and flushes queued events to listeners", () => {
    const listener = vi.fn();
    const session = {
      documentId: "doc-1",
      eventBacklog: [] as AgentEvent[],
      eventFlushScheduled: false,
    };

    const scheduledCallbacks: Array<() => void> = [];
    queueLocalAgentEvent(
      {
        getSession: () => session,
        listeners: new Set([listener]),
        schedule: (callback) => {
          scheduledCallbacks.push(callback);
        },
      },
      "doc-1",
      createEvent("sub-1"),
    );

    expect(session.eventBacklog).toHaveLength(1);
    expect(listener).not.toHaveBeenCalled();
    expect(scheduledCallbacks).toHaveLength(1);
    scheduledCallbacks[0]!();
    expect(listener).toHaveBeenCalledWith(createEvent("sub-1"));
    expect(session.eventBacklog).toEqual([]);
  });

  it("removes queued events for a deleted subscription", () => {
    const session = {
      documentId: "doc-1",
      eventBacklog: [createEvent("sub-1"), createEvent("sub-2")],
      eventFlushScheduled: false,
    };

    removeQueuedSubscriptionEvents(session, "sub-1");

    expect(session.eventBacklog).toEqual([createEvent("sub-2")]);
  });

  it("can flush pending events synchronously when listeners attach later", () => {
    const listener = vi.fn();
    const session = {
      documentId: "doc-1",
      eventBacklog: [createEvent("sub-1")],
      eventFlushScheduled: false,
    };

    flushQueuedLocalAgentEvents(
      {
        getSession: () => session,
        listeners: new Set([listener]),
      },
      "doc-1",
    );

    expect(listener).toHaveBeenCalledWith(createEvent("sub-1"));
    expect(session.eventBacklog).toEqual([]);
  });
});
