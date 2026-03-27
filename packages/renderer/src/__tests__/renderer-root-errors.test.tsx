import React from "react";
import { SpreadsheetEngine } from "@bilig/core";
import { afterEach, describe, expect, it, vi } from "vitest";
import { Cell, Sheet, Workbook } from "../components.js";

type CompatMocks = {
  createFiberRoot: ReturnType<typeof vi.fn>;
  updateFiberRoot: ReturnType<typeof vi.fn>;
};

async function loadRendererRootWithCompatMock(
  updateImpl: (
    element: React.ReactNode,
    callback: () => void,
    container: { lastError: unknown },
  ) => void,
): Promise<
  CompatMocks & {
    createWorkbookRendererRoot: typeof import("../renderer-root.js").createWorkbookRendererRoot;
  }
> {
  vi.resetModules();
  let capturedContainer: { lastError: unknown } | undefined;
  const createFiberRoot = vi.fn((container: { lastError: unknown }) => {
    capturedContainer = container;
    return { kind: "fiber-root" };
  });
  const updateFiberRoot = vi.fn(
    (_root: unknown, element: React.ReactNode, callback: () => void) => {
      updateImpl(element, callback, capturedContainer ?? { lastError: null });
    },
  );
  vi.doMock("../compat.js", () => ({
    createFiberRoot,
    updateFiberRoot,
  }));
  const { createWorkbookRendererRoot } = await import("../renderer-root.js");
  return { createWorkbookRendererRoot, createFiberRoot, updateFiberRoot };
}

afterEach(() => {
  vi.resetModules();
  vi.doUnmock("../compat.js");
  vi.clearAllMocks();
});

describe("renderer root error handling", () => {
  it("ignores duplicate compat callbacks and deletes committed sheets on unmount", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "renderer-root-double-callback" });
    await engine.ready();
    const renderCommit = vi.spyOn(engine, "renderCommit");
    const { createWorkbookRendererRoot } = await loadRendererRootWithCompatMock(
      (_element, callback) => {
        callback();
        callback();
      },
    );
    const root = createWorkbookRendererRoot(engine);

    await root.render(
      <Workbook name="book">
        <Sheet name="Sheet1">
          <Cell addr="A1" value={1} />
        </Sheet>
      </Workbook>,
    );

    const container = vi.mocked((await import("../compat.js")).createFiberRoot).mock
      .calls[0]?.[0] as
      | {
          root: unknown;
        }
      | undefined;
    if (!container) {
      throw new Error("Expected mocked container");
    }
    container.root = {
      kind: "Workbook",
      parent: null,
      container,
      props: { name: "book" },
      children: [
        {
          kind: "Sheet",
          parent: null,
          container,
          props: { name: "Sheet1" },
          children: [],
        },
      ],
    };

    await root.unmount();

    expect(renderCommit).toHaveBeenCalledWith([{ kind: "deleteSheet", name: "Sheet1" }]);
  });

  it("rejects unmount when the compat callback surfaces an async error", async () => {
    const { createWorkbookRendererRoot } = await loadRendererRootWithCompatMock(
      (element, callback, container) => {
        callback();
        if (element === null) {
          queueMicrotask(() => {
            container.lastError = new Error("async unmount failed");
          });
        }
      },
    );
    const engine = new SpreadsheetEngine({ workbookName: "renderer-root-unmount-async-error" });
    await engine.ready();
    const root = createWorkbookRendererRoot(engine);

    await root.render(
      <Workbook name="book">
        <Sheet name="Sheet1">
          <Cell addr="A1" value={1} />
        </Sheet>
      </Workbook>,
    );

    await expect(root.unmount()).rejects.toThrow("async unmount failed");
  });

  it("rejects root text nodes before commit", async () => {
    const { createWorkbookRendererRoot } = await loadRendererRootWithCompatMock(() => {});
    const engine = new SpreadsheetEngine({ workbookName: "renderer-root-text-node-error" });
    await engine.ready();
    const root = createWorkbookRendererRoot(engine);

    await expect(root.render("bad")).rejects.toThrow("Workbook DSL does not support text nodes.");
  });
});
