// @vitest-environment jsdom
import { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";
import { App } from "../App";

vi.mock("@rocicorp/zero/react", () => ({
  useZero: () => ({
    mutate: () => ({
      client: Promise.resolve({ type: "complete" }),
    }),
    materialize: () => ({
      data: [],
      addListener: () => () => {},
      destroy() {},
    }),
  }),
  useConnectionState: () => ({ name: "connected" }),
}));

vi.mock("../WorkerWorkbookApp", () => ({
  WorkerWorkbookApp: ({
    config,
    connectionState,
  }: {
    config: {
      defaultDocumentId: string;
      zeroCacheUrl: string;
      persistState: boolean;
      currentUserId: string;
    };
    connectionState: { name: string };
  }) => (
    <div
      data-connection-state={connectionState.name}
      data-default-document-id={config.defaultDocumentId}
      data-current-user-id={config.currentUserId}
      data-persist-state={String(config.persistState)}
      data-testid="workbook-shell"
      data-zero-cache-url={config.zeroCacheUrl}
    >
      workbook shell
    </div>
  ),
}));

afterEach(() => {
  document.body.innerHTML = "";
});

describe("web shell", () => {
  it("renders the minimal product shell without legacy demo chrome", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <App
          config={{
            zeroCacheUrl: "http://127.0.0.1:4848",
            defaultDocumentId: "bilig-demo",
            persistState: true,
            currentUserId: "guest:test",
          }}
        />,
      );
    });

    const workbookShell = host.querySelector("[data-testid='workbook-shell']");
    expect(workbookShell).not.toBeNull();
    expect(workbookShell?.getAttribute("data-default-document-id")).toBe("bilig-demo");
    expect(workbookShell?.getAttribute("data-current-user-id")).toBe("guest:test");
    expect(workbookShell?.getAttribute("data-zero-cache-url")).toBe("http://127.0.0.1:4848");
    expect(workbookShell?.getAttribute("data-persist-state")).toBe("true");
    expect(workbookShell?.getAttribute("data-connection-state")).toBe("connected");

    expect(host.querySelector("[data-testid='preset-strip']")).toBeNull();
    expect(host.querySelector("[data-testid='metrics-panel']")).toBeNull();
    expect(host.querySelector("[data-testid='replica-panel']")).toBeNull();
    expect(host.querySelector("[data-testid='ax-rail']")).toBeNull();
    expect(host.querySelector("[data-testid='ax-presence-chip']")).toBeNull();
    expect(host.querySelector("[data-testid='worker-loading']")).toBeNull();
    expect(host.querySelector("[data-testid='worker-error']")).toBeNull();
    expect(host.querySelector("h1")).toBeNull();
    expect(host.textContent).not.toContain("Excel-scale shell on top of the local-first engine");
    expect(host.textContent).not.toContain("Starting workbook runtime...");

    await act(async () => {
      root.unmount();
    });
  });
});
