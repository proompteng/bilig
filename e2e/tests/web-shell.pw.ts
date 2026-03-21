import { spawn, type ChildProcess } from "node:child_process";
import { createRequire } from "node:module";
import { createServer } from "node:net";
import { expect, test } from "@playwright/test";

const PRODUCT_ROW_MARKER_WIDTH = 46;
const PRODUCT_COLUMN_WIDTH = 104;
const PRODUCT_HEADER_HEIGHT = 24;
const PRODUCT_ROW_HEIGHT = 22;
const PRIMARY_MODIFIER = process.platform === "darwin" ? "Meta" : "Control";
const AGENT_STDIN_MAGIC = 0x41474e54;
const AGENT_PROTOCOL_VERSION = 1;
const textEncoder = new TextEncoder();
const textDecoder = new TextDecoder();
const require = createRequire(import.meta.url);
const tsxCliPath = require.resolve("tsx/cli");

interface LocalDocumentStateSummary {
  documentId: string;
  cursor: number;
  browserSessions: string[];
  agentSessions: string[];
  lastBatchId: string | null;
}

type CellRangeRef = {
  sheetName: string;
  startAddress: string;
  endAddress: string;
};

type AgentRequest =
  | { kind: "openWorkbookSession"; id: string; documentId: string; replicaId: string }
  | { kind: "closeWorkbookSession"; id: string; sessionId: string }
  | { kind: "writeRange"; id: string; sessionId: string; range: CellRangeRef; values: unknown[][] };

type AgentResponse =
  | { kind: "ok"; id: string; sessionId?: string; value?: unknown }
  | { kind: "error"; id: string; code: string; message: string; retryable: boolean };

type AgentFrame =
  | { kind: "request"; request: AgentRequest }
  | { kind: "response"; response: AgentResponse };

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isAgentFrame(value: unknown): value is AgentFrame {
  return isRecord(value)
    && (value.kind === "request" || value.kind === "response")
    && (("request" in value && isRecord(value.request)) || ("response" in value && isRecord(value.response)));
}

function delay(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function reserveLocalPort(): Promise<number> {
  const server = createServer();
  await new Promise<void>((resolve, reject) => {
    server.listen(0, "127.0.0.1", () => resolve());
    server.once("error", reject);
  });
  const address = server.address();
  if (!address || typeof address === "string") {
    server.close();
    throw new Error("Failed to reserve local-server port");
  }
  const port = address.port;
  await new Promise<void>((resolve, reject) => {
    server.close((error) => {
      if (error) {
        reject(error);
        return;
      }
      resolve();
    });
  });
  return port;
}

async function waitForLocalServerHealthy(localServerUrl: string, timeoutMs = 15_000) {
  const deadline = Date.now() + timeoutMs;
  const poll = async (lastError: string | null): Promise<void> => {
    if (Date.now() >= deadline) {
      throw new Error(`Timed out waiting for local-server on ${localServerUrl}${lastError ? `: ${lastError}` : ""}`);
    }
    try {
      const response = await fetch(`${localServerUrl}/healthz`);
      if (response.ok) {
        return;
      }
      lastError = `healthz returned ${response.status}`;
    } catch (error) {
      lastError = error instanceof Error ? error.message : String(error);
    }
    await delay(250);
    await poll(lastError);
  };

  await poll(null);
}

async function stopChildProcess(process: ChildProcess) {
  if (process.exitCode !== null) {
    return;
  }

  const exitPromise = new Promise<void>((resolve) => {
    process.once("exit", () => {
      resolve();
    });
  });

  process.kill("SIGTERM");
  const exited = await Promise.race([
    exitPromise.then(() => true),
    delay(5_000).then(() => false)
  ]);
  if (exited) {
    return;
  }

  process.kill("SIGKILL");
  await exitPromise;
}

async function startLocalServer(port: number) {
  const child = spawn(process.execPath, [tsxCliPath, "apps/local-server/src/index.ts"], {
    cwd: process.cwd(),
    env: {
      ...process.env,
      HOST: "127.0.0.1",
      PORT: String(port)
    },
    stdio: ["ignore", "pipe", "pipe"]
  });
  const localServerUrl = `http://127.0.0.1:${port}`;

  let logs = "";
  const appendLogChunk = (chunk: Buffer | string) => {
    logs += chunk.toString();
    if (logs.length > 12_000) {
      logs = logs.slice(-12_000);
    }
  };
  child.stdout?.on("data", appendLogChunk);
  child.stderr?.on("data", appendLogChunk);

  const exitPromise = new Promise<{ code: number | null; signal: NodeJS.Signals | null }>((resolve) => {
    child.once("exit", (code, signal) => {
      resolve({ code, signal });
    });
  });

  try {
    await Promise.race([
      waitForLocalServerHealthy(localServerUrl),
      exitPromise.then(({ code, signal }) => {
        throw new Error(
          `local-server exited before becoming healthy (code=${code ?? "null"}, signal=${signal ?? "null"})\n${logs}`
        );
      })
    ]);
  } catch (error) {
    await stopChildProcess(child);
    throw error;
  }

  return {
    localServerUrl,
    stop: async () => {
      await stopChildProcess(child);
    },
    getLogs: () => logs
  };
}

function encodeAgentFrame(frame: AgentFrame): Uint8Array {
  const payload = textEncoder.encode(JSON.stringify(frame));
  const output = new Uint8Array(10 + payload.byteLength);
  const view = new DataView(output.buffer);
  view.setUint32(0, AGENT_STDIN_MAGIC, true);
  view.setUint16(4, AGENT_PROTOCOL_VERSION, true);
  view.setUint32(6, payload.byteLength, true);
  output.set(payload, 10);
  return output;
}

function decodeAgentFrame(bytes: Uint8Array | ArrayBuffer): AgentFrame {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (data.byteLength < 10) {
    throw new Error("Agent frame too short");
  }
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  if (view.getUint32(0, true) !== AGENT_STDIN_MAGIC) {
    throw new Error("Agent frame magic mismatch");
  }
  if (view.getUint16(4, true) !== AGENT_PROTOCOL_VERSION) {
    throw new Error("Unsupported agent protocol version");
  }
  const payloadLength = view.getUint32(6, true);
  if (payloadLength !== data.byteLength - 10) {
    throw new Error("Agent frame length mismatch");
  }
  const parsed: unknown = JSON.parse(textDecoder.decode(data.subarray(10)));
  if (!isAgentFrame(parsed)) {
    throw new Error("Invalid agent frame payload");
  }
  return parsed;
}

async function sendAgentRequest(localServerUrl: string, request: AgentRequest): Promise<AgentResponse> {
  const response = await fetch(`${localServerUrl}/v1/agent/frames`, {
    method: "POST",
    headers: {
      "content-type": "application/octet-stream"
    },
    body: Buffer.from(encodeAgentFrame({
      kind: "request",
      request
    }))
  });
  if (!response.ok) {
    throw new Error(`Agent request failed with status ${response.status}`);
  }
  const frame = decodeAgentFrame(new Uint8Array(await response.arrayBuffer()));
  if (frame.kind !== "response") {
    throw new Error(`Expected agent response frame, received ${frame.kind}`);
  }
  if (frame.response.kind === "error") {
    throw new Error(`${frame.response.code}: ${frame.response.message}`);
  }
  return frame.response;
}

async function withAgentSession(
  localServerUrl: string,
  documentId: string,
  replicaId: string,
  callback: (sessionId: string) => Promise<void>
) {
  const openResponse = await sendAgentRequest(localServerUrl, {
    kind: "openWorkbookSession",
    id: `open:${Date.now()}`,
    documentId,
    replicaId
  });
  if (openResponse.kind !== "ok" || !openResponse.sessionId) {
    throw new Error("Failed to open local-server workbook session");
  }

  try {
    await callback(openResponse.sessionId);
  } finally {
    await sendAgentRequest(localServerUrl, {
      kind: "closeWorkbookSession",
      id: `close:${Date.now()}`,
      sessionId: openResponse.sessionId
    }).catch(() => undefined);
  }
}

function isLocalDocumentStateSummary(value: unknown): value is LocalDocumentStateSummary {
  return isRecord(value)
    && typeof value.documentId === "string"
    && typeof value.cursor === "number"
    && Array.isArray(value.browserSessions)
    && Array.isArray(value.agentSessions)
    && (typeof value.lastBatchId === "string" || value.lastBatchId === null);
}

async function fetchDocumentState(localServerUrl: string, documentId: string): Promise<LocalDocumentStateSummary> {
  const response = await fetch(`${localServerUrl}/v1/documents/${documentId}/state`);
  if (!response.ok) {
    throw new Error(`Failed to fetch local-server document state: ${response.status}`);
  }
  const payload: unknown = await response.json();
  if (!isLocalDocumentStateSummary(payload)) {
    throw new Error("Invalid local-server document state payload");
  }
  return payload;
}

function parseColumnWidthOverrides(raw: string | null): Record<string, number> {
  if (!raw) {
    return {};
  }
  const parsed: unknown = JSON.parse(raw);
  if (typeof parsed !== "object" || parsed === null) {
    return {};
  }
  const entries = Object.entries(parsed).filter((entry): entry is [string, number] => typeof entry[1] === "number");
  return Object.fromEntries(entries);
}

async function getProductColumnWidth(page: Parameters<typeof test>[0]["page"], columnIndex: number) {
  const grid = page.getByTestId("sheet-grid");
  const [defaultWidthRaw, overridesRaw] = await Promise.all([
    grid.getAttribute("data-default-column-width"),
    grid.getAttribute("data-column-width-overrides")
  ]);
  const defaultWidth = Number(defaultWidthRaw ?? String(PRODUCT_COLUMN_WIDTH));
  const overrides = parseColumnWidthOverrides(overridesRaw);
  return overrides[String(columnIndex)] ?? defaultWidth;
}

async function getProductColumnLeft(page: Parameters<typeof test>[0]["page"], columnIndex: number) {
  const widths = await Promise.all(
    Array.from({ length: columnIndex }, (_, index) => getProductColumnWidth(page, index))
  );
  return PRODUCT_ROW_MARKER_WIDTH + widths.reduce((total, width) => total + width, 0);
}

async function dragProductColumnResize(
  page: Parameters<typeof test>[0]["page"],
  columnIndex: number,
  deltaX: number
) {
  const gridLocator = page.getByTestId("sheet-grid");
  await expect(gridLocator).toBeVisible();
  const grid = await gridLocator.boundingBox();
  if (!grid) {
    throw new Error("sheet grid is not visible");
  }

  const columnLeft = await getProductColumnLeft(page, columnIndex);
  const columnWidth = await getProductColumnWidth(page, columnIndex);
  const edgeX = grid.x + columnLeft + columnWidth - 1;
  const edgeY = grid.y + Math.floor(PRODUCT_HEADER_HEIGHT / 2);

  await page.mouse.move(edgeX, edgeY);
  await page.mouse.down();
  await page.mouse.move(edgeX + deltaX, edgeY, { steps: 10 });
  await page.mouse.up();
}

async function doubleClickProductColumnResizeHandle(
  page: Parameters<typeof test>[0]["page"],
  columnIndex: number
) {
  const gridLocator = page.getByTestId("sheet-grid");
  await expect(gridLocator).toBeVisible();
  const grid = await gridLocator.boundingBox();
  if (!grid) {
    throw new Error("sheet grid is not visible");
  }

  const columnLeft = await getProductColumnLeft(page, columnIndex);
  const columnWidth = await getProductColumnWidth(page, columnIndex);
  const edgeX = grid.x + columnLeft + columnWidth - 1;
  const headerY = grid.y + Math.floor(PRODUCT_HEADER_HEIGHT / 2);
  await page.mouse.click(edgeX, headerY, { clickCount: 2 });
}

async function dragProductHeaderSelection(
  page: Parameters<typeof test>[0]["page"],
  axis: "column" | "row",
  startIndex: number,
  endIndex: number
) {
  const gridLocator = page.getByTestId("sheet-grid");
  await expect(gridLocator).toBeVisible();
  const grid = await gridLocator.boundingBox();
  if (!grid) {
    throw new Error("sheet grid is not visible");
  }

  const startColumnLeft = axis === "column" ? await getProductColumnLeft(page, startIndex) : 0;
  const endColumnLeft = axis === "column" ? await getProductColumnLeft(page, endIndex) : 0;
  const startColumnWidth = axis === "column" ? await getProductColumnWidth(page, startIndex) : 0;
  const endColumnWidth = axis === "column" ? await getProductColumnWidth(page, endIndex) : 0;
  const startX = axis === "column"
    ? grid.x + startColumnLeft + Math.floor(startColumnWidth / 2)
    : grid.x + Math.floor(PRODUCT_ROW_MARKER_WIDTH / 2);
  const startY = axis === "column"
    ? grid.y + Math.floor(PRODUCT_HEADER_HEIGHT / 2)
    : grid.y + PRODUCT_HEADER_HEIGHT + (startIndex * PRODUCT_ROW_HEIGHT) + Math.floor(PRODUCT_ROW_HEIGHT / 2);
  const endX = axis === "column"
    ? grid.x + endColumnLeft + Math.floor(endColumnWidth / 2)
    : grid.x + Math.floor(PRODUCT_ROW_MARKER_WIDTH / 2);
  const endY = axis === "column"
    ? grid.y + Math.floor(PRODUCT_HEADER_HEIGHT / 2)
    : grid.y + PRODUCT_HEADER_HEIGHT + (endIndex * PRODUCT_ROW_HEIGHT) + Math.floor(PRODUCT_ROW_HEIGHT / 2);

  await page.mouse.move(startX, startY);
  await page.mouse.down();
  await page.mouse.move(endX, endY, { steps: 8 });
  await page.mouse.up();
}

async function clickGridRightEdge(page: Parameters<typeof test>[0]["page"], rowIndex = 2) {
  const gridLocator = page.getByTestId("sheet-grid");
  await expect(gridLocator).toBeVisible();
  const grid = await gridLocator.boundingBox();
  if (!grid) {
    throw new Error("sheet grid is not visible");
  }

  const x = grid.x + grid.width - 3;
  const y = grid.y + PRODUCT_HEADER_HEIGHT + (rowIndex * PRODUCT_ROW_HEIGHT) + Math.floor(PRODUCT_ROW_HEIGHT / 2);
  await page.mouse.click(x, y);
}

async function dragProductFillHandle(
  page: Parameters<typeof test>[0]["page"],
  sourceCol: number,
  sourceRow: number,
  targetCol: number,
  targetRow: number
) {
  const gridLocator = page.getByTestId("sheet-grid");
  await expect(gridLocator).toBeVisible();
  const grid = await gridLocator.boundingBox();
  if (!grid) {
    throw new Error("sheet grid is not visible");
  }

  const sourceLeft = grid.x + await getProductColumnLeft(page, sourceCol);
  const sourceTop = grid.y + PRODUCT_HEADER_HEIGHT + (sourceRow * PRODUCT_ROW_HEIGHT);
  const targetLeft = grid.x + await getProductColumnLeft(page, targetCol);
  const targetTop = grid.y + PRODUCT_HEADER_HEIGHT + (targetRow * PRODUCT_ROW_HEIGHT);
  const sourceWidth = await getProductColumnWidth(page, sourceCol);
  const targetWidth = await getProductColumnWidth(page, targetCol);

  await page.mouse.move(sourceLeft + sourceWidth - 3, sourceTop + PRODUCT_ROW_HEIGHT - 3);
  await page.mouse.down();
  await page.mouse.move(targetLeft + targetWidth - 3, targetTop + PRODUCT_ROW_HEIGHT - 3, { steps: 10 });
  await page.mouse.up();
}

async function clickProductBodyOffset(
  page: Parameters<typeof test>[0]["page"],
  offsetX: number,
  rowIndex = 0
) {
  const gridLocator = page.getByTestId("sheet-grid");
  await expect(gridLocator).toBeVisible();
  const grid = await gridLocator.boundingBox();
  if (!grid) {
    throw new Error("sheet grid is not visible");
  }

  await page.mouse.click(
    grid.x + PRODUCT_ROW_MARKER_WIDTH + offsetX,
    grid.y + PRODUCT_HEADER_HEIGHT + (rowIndex * PRODUCT_ROW_HEIGHT) + Math.floor(PRODUCT_ROW_HEIGHT / 2)
  );
}

async function clickProductCell(
  page: Parameters<typeof test>[0]["page"],
  columnIndex: number,
  rowIndex: number
) {
  const gridLocator = page.getByTestId("sheet-grid");
  await expect(gridLocator).toBeVisible();
  const grid = await gridLocator.boundingBox();
  if (!grid) {
    throw new Error("sheet grid is not visible");
  }

  const columnLeft = await getProductColumnLeft(page, columnIndex);
  const columnWidth = await getProductColumnWidth(page, columnIndex);
  await page.mouse.click(
    grid.x + columnLeft + Math.floor(columnWidth / 2),
    grid.y + PRODUCT_HEADER_HEIGHT + (rowIndex * PRODUCT_ROW_HEIGHT) + Math.floor(PRODUCT_ROW_HEIGHT / 2)
  );
}

async function clickProductCellUpperHalf(
  page: Parameters<typeof test>[0]["page"],
  columnIndex: number,
  rowIndex: number
) {
  const gridLocator = page.getByTestId("sheet-grid");
  await expect(gridLocator).toBeVisible();
  const grid = await gridLocator.boundingBox();
  if (!grid) {
    throw new Error("sheet grid is not visible");
  }

  const columnLeft = await getProductColumnLeft(page, columnIndex);
  const columnWidth = await getProductColumnWidth(page, columnIndex);
  await page.mouse.click(
    grid.x + columnLeft + Math.floor(columnWidth / 2),
    grid.y + PRODUCT_HEADER_HEIGHT + (rowIndex * PRODUCT_ROW_HEIGHT) + 4
  );
}

async function clickProductSelectedCellTopBorder(
  page: Parameters<typeof test>[0]["page"],
  columnIndex: number,
  rowIndex: number
) {
  const gridLocator = page.getByTestId("sheet-grid");
  await expect(gridLocator).toBeVisible();
  const grid = await gridLocator.boundingBox();
  if (!grid) {
    throw new Error("sheet grid is not visible");
  }

  const columnLeft = await getProductColumnLeft(page, columnIndex);
  const columnWidth = await getProductColumnWidth(page, columnIndex);
  await page.mouse.click(
    grid.x + columnLeft + Math.floor(columnWidth / 2),
    grid.y + PRODUCT_HEADER_HEIGHT + (rowIndex * PRODUCT_ROW_HEIGHT) - 1
  );
}

async function dragProductBodySelection(
  page: Parameters<typeof test>[0]["page"],
  startColumn: number,
  startRow: number,
  endColumn: number,
  endRow: number
) {
  const gridLocator = page.getByTestId("sheet-grid");
  await expect(gridLocator).toBeVisible();
  const grid = await gridLocator.boundingBox();
  if (!grid) {
    throw new Error("sheet grid is not visible");
  }

  const startLeft = await getProductColumnLeft(page, startColumn);
  const startWidth = await getProductColumnWidth(page, startColumn);
  const endLeft = await getProductColumnLeft(page, endColumn);
  const endWidth = await getProductColumnWidth(page, endColumn);

  const startX = grid.x + startLeft + Math.floor(startWidth / 2);
  const startY = grid.y + PRODUCT_HEADER_HEIGHT + (startRow * PRODUCT_ROW_HEIGHT) + Math.floor(PRODUCT_ROW_HEIGHT / 2);
  const endX = grid.x + endLeft + Math.floor(endWidth / 2);
  const endY = grid.y + PRODUCT_HEADER_HEIGHT + (endRow * PRODUCT_ROW_HEIGHT) + Math.floor(PRODUCT_ROW_HEIGHT / 2);

  await page.mouse.move(startX, startY);
  await page.mouse.down();
  await page.mouse.move(endX, endY, { steps: 12 });
  await page.mouse.up();
}

test("web app renders the minimal product shell without playground chrome", async ({ page }) => {
  await page.goto("/");

  await expect(page.getByTestId("formula-bar")).toBeVisible();
  await expect(page.getByTestId("name-box")).toBeVisible();
  await expect(page.getByTestId("sheet-grid")).toBeVisible();
  await expect(page.getByRole("tab", { name: "Sheet1" })).toBeVisible();
  await expect(page.getByRole("heading", { name: "bilig-demo" })).toHaveCount(0);

  await expect(page.getByTestId("preset-strip")).toHaveCount(0);
  await expect(page.getByTestId("metrics-panel")).toHaveCount(0);
  await expect(page.getByTestId("replica-panel")).toHaveCount(0);

  await expect(page.getByTestId("status-mode")).toHaveText(/^(Local|Live)$/);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A1");
  await expect(page.getByTestId("status-sync")).toHaveText("Ready");
  await expect(page.locator(".formula-result-shell")).toHaveCount(0);
});

test("web app reflects a local-server agent write in the rendered spreadsheet", async ({ page }) => {
  test.slow();
  const port = await reserveLocalPort();
  const documentId = `playwright-${Date.now()}`;
  const localServer = await startLocalServer(port);

  try {
    await page.goto(`/?document=${encodeURIComponent(documentId)}&server=${encodeURIComponent(localServer.localServerUrl)}`);

    const nameBox = page.getByTestId("name-box");
    const formulaInput = page.getByTestId("formula-input");

    await expect(nameBox).toHaveValue("A1");
    await expect(formulaInput).toHaveValue("");

    await expect.poll(async () => {
      const documentState = await fetchDocumentState(localServer.localServerUrl, documentId);
      return documentState.browserSessions.length > 0;
    }, {
      message: "browser should attach to the local-server document session"
    }).toBe(true);

    await withAgentSession(localServer.localServerUrl, documentId, `playwright-agent:${Date.now()}`, async (sessionId) => {
      await sendAgentRequest(localServer.localServerUrl, {
        kind: "writeRange",
        id: `write:${Date.now()}`,
        sessionId,
        range: {
          sheetName: "Sheet1",
          startAddress: "A1",
          endAddress: "A1"
        },
        values: [[42]]
      });
    });

    await expect(formulaInput).toHaveValue("42");
  } catch (error) {
    throw new Error(
      `${error instanceof Error ? error.message : String(error)}\nLocal-server logs:\n${localServer.getLogs()}`,
      { cause: error }
    );
  } finally {
    await localServer.stop();
  }
});

test("web app keeps sheet tabs and status bar visible in a short viewport", async ({ page }) => {
  await page.setViewportSize({ width: 2048, height: 220 });
  await page.goto("/");

  const sheetTab = page.getByRole("tab", { name: "Sheet1" });
  const statusSync = page.getByTestId("status-sync");

  await expect(sheetTab).toBeVisible();
  await expect(statusSync).toBeVisible();

  const tabBox = await sheetTab.boundingBox();
  const statusBox = await statusSync.boundingBox();
  if (!tabBox || !statusBox) {
    throw new Error("footer controls are not visible");
  }

  expect(tabBox.y + tabBox.height).toBeLessThanOrEqual(220);
  expect(statusBox.y + statusBox.height).toBeLessThanOrEqual(220);
});

test("web app supports column and row header selection", async ({ page }) => {
  await page.goto("/");

  const grid = page.getByTestId("sheet-grid");

  await grid.click({ position: { x: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH + Math.floor(PRODUCT_COLUMN_WIDTH / 2), y: Math.floor(PRODUCT_HEADER_HEIGHT / 2) } });
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!B:B");

  await grid.click({ position: { x: Math.floor(PRODUCT_ROW_MARKER_WIDTH / 2), y: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2) } });
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!2:2");
});

test("web app supports row and column header drag selection", async ({ page }) => {
  await page.goto("/");

  await dragProductHeaderSelection(page, "column", 1, 3);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!B:D");

  await dragProductHeaderSelection(page, "row", 1, 3);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!2:4");
});

test("web app supports rectangular drag selection", async ({ page }) => {
  await page.goto("/");

  await dragProductBodySelection(page, 1, 1, 3, 3);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!B2:D4");
});

test("web app keeps the active focus inside the Glide grid when clicking a cell", async ({ page }) => {
  await page.goto("/");

  await clickProductCell(page, 2, 2);
  await expect(page.getByTestId("name-box")).toHaveValue("C3");

  const activeElementState = await page.evaluate(() => {
    const active = document.activeElement;
    return {
      testId: active?.getAttribute("data-testid") ?? null,
      insideSheetGrid: Boolean(active?.closest('[data-testid="sheet-grid"]'))
    };
  });

  expect(activeElementState.insideSheetGrid).toBe(true);
  expect(activeElementState.testId).not.toBe("sheet-grid");
});

test("web app maps clicks in the upper half of a cell to that same visible cell", async ({ page }) => {
  await page.goto("/");

  await clickProductCellUpperHalf(page, 4, 11);
  await expect(page.getByTestId("name-box")).toHaveValue("E12");
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!E12");

  await clickProductCellUpperHalf(page, 2, 4);
  await expect(page.getByTestId("name-box")).toHaveValue("C5");
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!C5");
});

test("web app supports column resize without breaking hit testing", async ({ page }) => {
  await page.goto("/");

  await clickProductBodyOffset(page, 82, 0);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A1");

  await dragProductColumnResize(page, 0, -36);

  await clickProductBodyOffset(page, 82, 0);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!B1");
});

test("web app supports column edge double-click autofit", async ({ page }) => {
  await page.goto("/");

  const nameBox = page.getByTestId("name-box");
  const formulaInput = page.getByTestId("formula-input");

  await nameBox.fill("A1");
  await nameBox.press("Enter");
  await formulaInput.fill("supercalifragilisticexpialidocious");
  await formulaInput.press("Enter");

  await clickProductBodyOffset(page, 126, 0);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!B1");

  await doubleClickProductColumnResizeHandle(page, 0);

  await clickProductBodyOffset(page, 126, 0);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A1");
});

test("web app accepts string values and string comparison formulas", async ({ page }) => {
  await page.goto("/");

  const nameBox = page.getByTestId("name-box");
  const formulaInput = page.getByTestId("formula-input");
  const resolvedValue = page.getByTestId("formula-resolved-value");

  await clickProductCell(page, 0, 0);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A1");
  await formulaInput.fill("hello");
  await formulaInput.press("Enter");
  await expect(nameBox).toHaveValue("A1");
  await expect(formulaInput).toHaveValue("hello");
  await clickProductCell(page, 0, 0);
  await expect(resolvedValue).toHaveText("hello");

  await nameBox.fill("A2");
  await nameBox.press("Enter");
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A2");
  await formulaInput.fill("=A1=\"HELLO\"");
  await formulaInput.press("Enter");
  await clickProductCell(page, 0, 1);
  await expect(resolvedValue).toHaveText("TRUE");
});

test("web app supports type-to-replace and Enter or Tab commit movement", async ({ page }) => {
  await page.goto("/");

  const grid = page.getByTestId("sheet-grid");
  const nameBox = page.getByTestId("name-box");
  const formulaInput = page.getByTestId("formula-input");
  const cellEditor = page.getByTestId("cell-editor-input");

  await clickProductCell(page, 0, 0);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A1");
  await grid.press("h");
  await expect(cellEditor).toBeVisible();
  await expect(cellEditor).toHaveValue("h");
  await page.keyboard.press("Enter");

  await expect(nameBox).toHaveValue("A2");
  await nameBox.fill("A1");
  await nameBox.press("Enter");
  await expect(formulaInput).toHaveValue("h");

  await clickProductCell(page, 0, 1);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A2");
  await grid.press("w");
  await expect(cellEditor).toBeVisible();
  await expect(cellEditor).toHaveValue("w");
  await page.keyboard.press("Tab");

  await expect(nameBox).toHaveValue("B2");
  await nameBox.fill("A2");
  await nameBox.press("Enter");
  await expect(formulaInput).toHaveValue("w");

  await grid.press("Enter");
  await expect(nameBox).toHaveValue("A3");
  await grid.press("Shift+Enter");
  await expect(nameBox).toHaveValue("A2");
});

test("web app preserves multi-digit numeric type-to-replace input", async ({ page }) => {
  await page.goto("/");

  const grid = page.getByTestId("sheet-grid");
  const nameBox = page.getByTestId("name-box");
  const formulaInput = page.getByTestId("formula-input");
  const cellEditor = page.getByTestId("cell-editor-input");

  await clickProductCell(page, 0, 0);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A1");

  await page.keyboard.type("123");
  await expect(cellEditor).toBeVisible();
  await expect(cellEditor).toHaveValue("123");
  await page.keyboard.press("Enter");

  await expect(nameBox).toHaveValue("A2");
  await nameBox.fill("A1");
  await nameBox.press("Enter");
  await expect(formulaInput).toHaveValue("123");

  await clickProductCell(page, 1, 0);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!B1");
  await grid.press("4");
  await expect(cellEditor).toBeVisible();
  await expect(cellEditor).toHaveValue("4");
});

test("web app right-aligns numeric in-cell editing like numeric view state", async ({ page }) => {
  await page.goto("/");

  const grid = page.getByTestId("sheet-grid");
  const cellEditor = page.getByTestId("cell-editor-input");

  await clickProductCell(page, 0, 0);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A1");
  await page.keyboard.type("123");
  await expect(cellEditor).toBeVisible();
  await expect(cellEditor).toHaveValue("123");
  await expect(cellEditor).toHaveCSS("text-align", "right");

  await page.keyboard.press("Escape");
  await clickProductCell(page, 1, 0);
  await grid.press("h");
  await expect(cellEditor).toBeVisible();
  await expect(cellEditor).toHaveValue("h");
  await expect(cellEditor).toHaveCSS("text-align", "left");
});

test("web app accepts numpad digits for in-cell numeric entry", async ({ page }) => {
  await page.goto("/");

  const nameBox = page.getByTestId("name-box");
  const formulaInput = page.getByTestId("formula-input");
  const cellEditor = page.getByTestId("cell-editor-input");

  await clickProductCell(page, 0, 0);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A1");

  await page.keyboard.press("Numpad1");
  await page.keyboard.press("Numpad2");
  await page.keyboard.press("Numpad3");
  await expect(cellEditor).toBeVisible();
  await expect(cellEditor).toHaveValue("123");
  await page.keyboard.press("Enter");

  await expect(nameBox).toHaveValue("A2");
  await nameBox.fill("A1");
  await nameBox.press("Enter");
  await expect(formulaInput).toHaveValue("123");
});

test("web app supports F2 edit in the product shell", async ({ page }) => {
  await page.goto("/");

  const grid = page.getByTestId("sheet-grid");
  const nameBox = page.getByTestId("name-box");
  const formulaInput = page.getByTestId("formula-input");
  const cellEditor = page.getByTestId("cell-editor-input");

  await nameBox.fill("C3");
  await nameBox.press("Enter");
  await formulaInput.fill("seed");
  await formulaInput.press("Enter");

  await clickProductCell(page, 2, 2);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!C3");
  await grid.press("F2");
  await expect(cellEditor).toBeVisible();
  await expect(cellEditor).toHaveValue("seed");
  await cellEditor.press("!");
  await expect(cellEditor).toHaveValue("seed!");
  await clickProductCell(page, 3, 2);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!D3");

  await clickProductCell(page, 2, 2);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!C3");
  await expect(formulaInput).toHaveValue("seed!");
});

test("web app double-click edits the exact clicked cell", async ({ page }) => {
  await page.goto("/");

  const nameBox = page.getByTestId("name-box");
  const formulaInput = page.getByTestId("formula-input");
  const cellEditor = page.getByTestId("cell-editor-input");
  const gridLocator = page.getByTestId("sheet-grid");

  await nameBox.fill("C4");
  await nameBox.press("Enter");
  await formulaInput.fill("above");
  await formulaInput.press("Enter");

  await nameBox.fill("C5");
  await nameBox.press("Enter");
  await formulaInput.fill("target");
  await formulaInput.press("Enter");

  await expect(gridLocator).toBeVisible();
  const grid = await gridLocator.boundingBox();
  if (!grid) {
    throw new Error("sheet grid is not visible");
  }

  const columnLeft = await getProductColumnLeft(page, 2);
  const columnWidth = await getProductColumnWidth(page, 2);
  const targetX = grid.x + columnLeft + Math.floor(columnWidth / 2);
  const targetY = grid.y + PRODUCT_HEADER_HEIGHT + (4 * PRODUCT_ROW_HEIGHT) + Math.floor(PRODUCT_ROW_HEIGHT / 2);
  await page.mouse.dblclick(targetX, targetY);

  await expect(nameBox).toHaveValue("C5");
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!C5");
  await expect(cellEditor).toBeVisible();
  await expect(cellEditor).toHaveValue("target");
  await expect(cellEditor).toHaveAttribute("aria-label", "Sheet1!C5 editor");
});

test("web app keeps the selected cell when clicking its top border", async ({ page }) => {
  await page.goto("/");

  const nameBox = page.getByTestId("name-box");

  await nameBox.fill("C5");
  await nameBox.press("Enter");
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!C5");

  await clickProductSelectedCellTopBorder(page, 2, 4);
  await expect(nameBox).toHaveValue("C5");
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!C5");
});

test("web app supports fill-handle propagation", async ({ page }) => {
  await page.goto("/");

  const nameBox = page.getByTestId("name-box");
  const formulaInput = page.getByTestId("formula-input");
  const resolvedValue = page.getByTestId("formula-resolved-value");

  await nameBox.fill("F6");
  await nameBox.press("Enter");
  await formulaInput.fill("7");
  await formulaInput.press("Enter");

  await dragProductFillHandle(page, 5, 5, 5, 7);

  await nameBox.fill("F8");
  await nameBox.press("Enter");
  await expect(formulaInput).toHaveValue("7");
  await expect(resolvedValue).toHaveText("7");
});

test("web app supports rectangular clipboard copy and external paste", async ({ page, context }) => {
  await context.grantPermissions(["clipboard-read", "clipboard-write"]);
  await page.goto("/");

  const grid = page.getByTestId("sheet-grid");
  const nameBox = page.getByTestId("name-box");
  const formulaInput = page.getByTestId("formula-input");
  const resolvedValue = page.getByTestId("formula-resolved-value");

  await nameBox.fill("B2");
  await nameBox.press("Enter");
  await formulaInput.fill("11");
  await formulaInput.press("Enter");

  await nameBox.fill("C2");
  await nameBox.press("Enter");
  await formulaInput.fill("12");
  await formulaInput.press("Enter");

  await nameBox.fill("B3");
  await nameBox.press("Enter");
  await formulaInput.fill("13");
  await formulaInput.press("Enter");

  await nameBox.fill("C3");
  await nameBox.press("Enter");
  await formulaInput.fill("14");
  await formulaInput.press("Enter");

  await dragProductBodySelection(page, 1, 1, 2, 2);
  await grid.press(`${PRIMARY_MODIFIER}+C`);

  await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).toBe("11\t12\n13\t14");

  await page.evaluate(() => navigator.clipboard.writeText("21\t22\n23\t24"));
  await clickProductCell(page, 4, 4);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!E5");
  await grid.press(`${PRIMARY_MODIFIER}+V`);

  await nameBox.fill("E5");
  await nameBox.press("Enter");
  await expect(formulaInput).toHaveValue("21");
  await expect(resolvedValue).toHaveText("21");

  await nameBox.fill("F5");
  await nameBox.press("Enter");
  await expect(formulaInput).toHaveValue("22");
  await expect(resolvedValue).toHaveText("22");

  await nameBox.fill("E6");
  await nameBox.press("Enter");
  await expect(formulaInput).toHaveValue("23");
  await expect(resolvedValue).toHaveText("23");

  await nameBox.fill("F6");
  await nameBox.press("Enter");
  await expect(formulaInput).toHaveValue("24");
  await expect(resolvedValue).toHaveText("24");
});

test("web app relocates formulas when using rectangular clipboard paste", async ({ page, context }) => {
  await context.grantPermissions(["clipboard-read", "clipboard-write"]);
  await page.goto("/");

  const grid = page.getByTestId("sheet-grid");
  const nameBox = page.getByTestId("name-box");
  const formulaInput = page.getByTestId("formula-input");
  const resolvedValue = page.getByTestId("formula-resolved-value");

  await nameBox.fill("B2");
  await nameBox.press("Enter");
  await formulaInput.fill("3");
  await formulaInput.press("Enter");

  await nameBox.fill("B3");
  await nameBox.press("Enter");
  await formulaInput.fill("4");
  await formulaInput.press("Enter");

  await nameBox.fill("C2");
  await nameBox.press("Enter");
  await formulaInput.fill("=B2*2");
  await formulaInput.press("Enter");

  await nameBox.fill("C3");
  await nameBox.press("Enter");
  await formulaInput.fill("=B3*2");
  await formulaInput.press("Enter");

  await dragProductBodySelection(page, 1, 1, 2, 2);
  await grid.press(`${PRIMARY_MODIFIER}+C`);

  await clickProductCell(page, 3, 1);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!D2");
  await grid.press(`${PRIMARY_MODIFIER}+V`);

  await nameBox.fill("E2");
  await nameBox.press("Enter");
  await expect(formulaInput).toHaveValue("=D2*2");
  await expect(resolvedValue).toHaveText("6");

  await nameBox.fill("E3");
  await nameBox.press("Enter");
  await expect(formulaInput).toHaveValue("=D3*2");
  await expect(resolvedValue).toHaveText("8");
});

test("web app supports product-shell column resize", async ({ page }) => {
  await page.goto("/");

  const baselineWidth = await getProductColumnWidth(page, 0);
  await dragProductColumnResize(page, 0, 48);
  await expect.poll(() => getProductColumnWidth(page, 0)).toBeGreaterThan(baselineWidth + 30);
});

test("web app relocates relative formulas when using the fill handle", async ({ page }) => {
  await page.goto("/");

  const nameBox = page.getByTestId("name-box");
  const formulaInput = page.getByTestId("formula-input");
  const resolvedValue = page.getByTestId("formula-resolved-value");

  await nameBox.fill("F6");
  await nameBox.press("Enter");
  await formulaInput.fill("3");
  await formulaInput.press("Enter");

  await nameBox.fill("F7");
  await nameBox.press("Enter");
  await formulaInput.fill("4");
  await formulaInput.press("Enter");

  await nameBox.fill("G6");
  await nameBox.press("Enter");
  await formulaInput.fill("=F6*2");
  await formulaInput.press("Enter");

  await dragProductFillHandle(page, 6, 5, 6, 6);

  await nameBox.fill("G7");
  await nameBox.press("Enter");
  await expect(nameBox).toHaveValue("G7");
  await expect(formulaInput).toHaveValue("=F7*2");
  await expect(resolvedValue).toHaveText("8");
});

test("web app shows #VALUE! for invalid formulas", async ({ page }) => {
  await page.goto("/");

  const formulaInput = page.getByTestId("formula-input");
  const resolvedValue = page.getByTestId("formula-resolved-value");

  await formulaInput.focus();
  await formulaInput.selectText();
  await page.keyboard.type("=1+");
  await formulaInput.press("Enter");

  await expect(formulaInput).toHaveValue("#VALUE!");
  await expect(resolvedValue).toHaveText("#VALUE!");
});

test("web app commits in-cell string edits when clicking away", async ({ page }) => {
  await page.goto("/");

  const grid = page.getByTestId("sheet-grid");
  const nameBox = page.getByTestId("name-box");
  const formulaInput = page.getByTestId("formula-input");
  const resolvedValue = page.getByTestId("formula-resolved-value");
  const cellEditor = page.getByTestId("cell-editor-input");

  await clickProductCell(page, 1, 0);
  await expect(nameBox).toHaveValue("B1");
  await grid.press("h");
  await expect(cellEditor).toBeVisible();
  await expect(cellEditor).toHaveValue("h");
  await clickProductCell(page, 2, 0);

  await expect(nameBox).toHaveValue("C1");
  await clickProductCell(page, 1, 0);
  await expect(nameBox).toHaveValue("B1");
  await expect(formulaInput).toHaveValue("h");
  await expect(resolvedValue).toHaveText("h");
});

test("web app ignores right gutter clicks", async ({ page }) => {
  await page.goto("/");

  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A1");
  await clickGridRightEdge(page, 3);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A1");
});
