import { describe, expect, it } from "vitest";
import { createMemoryWorkbookLocalStoreFactory } from "@bilig/storage-browser";
import { ValueTag } from "@bilig/protocol";
import type { AuthoritativeWorkbookEventRecord } from "@bilig/zero-sync";
import { WorkbookWorkerRuntime } from "../worker-runtime.js";

function buildSetCellValueEvent(input: {
  revision: number;
  address: string;
  value: number;
  clientMutationId?: string | null;
}): AuthoritativeWorkbookEventRecord {
  return {
    revision: input.revision,
    clientMutationId: input.clientMutationId ?? null,
    payload: {
      kind: "setCellValue",
      sheetName: "Sheet1",
      address: input.address,
      value: input.value,
    },
  };
}

function expectNumberValue(runtime: WorkbookWorkerRuntime, address: string, value: number): void {
  expect(runtime.getCell("Sheet1", address).value).toEqual({
    tag: ValueTag.Number,
    value,
  });
}

describe("WorkbookWorkerRuntime reconnect rebase", () => {
  it("replays pending local mutations over authoritative drift and absorbs them on ack", async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryWorkbookLocalStoreFactory(),
    });
    await runtime.bootstrap({
      documentId: "reconnect-doc",
      replicaId: "browser:test",
      persistState: true,
    });

    const pendingA1 = await runtime.enqueuePendingMutation({
      method: "setCellValue",
      args: ["Sheet1", "A1", 11],
    });
    const pendingA2 = await runtime.enqueuePendingMutation({
      method: "setCellValue",
      args: ["Sheet1", "A2", 22],
    });
    const pendingA3 = await runtime.enqueuePendingMutation({
      method: "setCellValue",
      args: ["Sheet1", "A3", 33],
    });

    await runtime.applyAuthoritativeEvents(
      [
        buildSetCellValueEvent({ revision: 1, address: "B1", value: 101 }),
        buildSetCellValueEvent({ revision: 2, address: "B2", value: 202 }),
        buildSetCellValueEvent({ revision: 3, address: "B3", value: 303 }),
      ],
      3,
    );

    expect(runtime.listPendingMutations().map((mutation) => mutation.id)).toEqual([
      pendingA1.id,
      pendingA2.id,
      pendingA3.id,
    ]);
    expectNumberValue(runtime, "A1", 11);
    expectNumberValue(runtime, "A2", 22);
    expectNumberValue(runtime, "A3", 33);
    expectNumberValue(runtime, "B1", 101);
    expectNumberValue(runtime, "B2", 202);
    expectNumberValue(runtime, "B3", 303);

    await runtime.markPendingMutationSubmitted(pendingA1.id);
    await runtime.markPendingMutationSubmitted(pendingA2.id);
    await runtime.markPendingMutationSubmitted(pendingA3.id);
    await runtime.applyAuthoritativeEvents(
      [
        buildSetCellValueEvent({
          revision: 4,
          address: "A1",
          value: 11,
          clientMutationId: pendingA1.id,
        }),
        buildSetCellValueEvent({
          revision: 5,
          address: "A2",
          value: 22,
          clientMutationId: pendingA2.id,
        }),
        buildSetCellValueEvent({
          revision: 6,
          address: "A3",
          value: 33,
          clientMutationId: pendingA3.id,
        }),
      ],
      6,
    );

    expect(runtime.listPendingMutations()).toEqual([]);
    expectNumberValue(runtime, "A1", 11);
    expectNumberValue(runtime, "A2", 22);
    expectNumberValue(runtime, "A3", 33);
    expectNumberValue(runtime, "B1", 101);
    expectNumberValue(runtime, "B2", 202);
    expectNumberValue(runtime, "B3", 303);
  });
});
