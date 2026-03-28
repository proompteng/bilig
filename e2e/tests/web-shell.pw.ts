import { spawn, type ChildProcess } from "node:child_process";
import { createServer } from "node:net";
import { expect, test, type Locator } from "@playwright/test";

const PRODUCT_ROW_MARKER_WIDTH = 46;
const PRODUCT_COLUMN_WIDTH = 104;
const PRODUCT_HEADER_HEIGHT = 24;
const PRODUCT_ROW_HEIGHT = 22;
const PRIMARY_MODIFIER = process.platform === "darwin" ? "Meta" : "Control";
const AGENT_STDIN_MAGIC = 0x41474e54;
const AGENT_PROTOCOL_VERSION = 1;
const BORDER_SIDES = ["top", "right", "bottom", "left"] as const;
const textEncoder = new TextEncoder();
const textDecoder = new TextDecoder();
const bunExecutable = process.env["BUN_BIN"] ?? "bun";

function parseTestCellAddress(address: string): { row: number; col: number } {
  const match = /^([A-Z]+)(\d+)$/i.exec(address.trim());
  if (!match) {
    throw new Error(`Invalid test cell address: ${address}`);
  }
  const letters = match[1].toUpperCase();
  const row = Number(match[2]) - 1;
  let col = 0;
  for (const letter of letters) {
    col = col * 26 + (letter.charCodeAt(0) - 64);
  }
  return { row, col: col - 1 };
}

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
  | { kind: "writeRange"; id: string; sessionId: string; range: CellRangeRef; values: unknown[][] }
  | { kind: "setRangeStyle"; id: string; sessionId: string; range: CellRangeRef; patch: unknown }
  | {
      kind: "setRangeNumberFormat";
      id: string;
      sessionId: string;
      range: CellRangeRef;
      format: unknown;
    }
  | { kind: "exportSnapshot"; id: string; sessionId: string };

type AgentResponse =
  | { kind: "ok"; id: string; sessionId?: string; value?: unknown }
  | { kind: "snapshot"; id: string; snapshot: WorkbookSnapshotLike }
  | { kind: "error"; id: string; code: string; message: string; retryable: boolean };

type AgentFrame =
  | { kind: "request"; request: AgentRequest }
  | { kind: "response"; response: AgentResponse };

interface WorkbookSnapshotLike {
  workbook: {
    name: string;
    metadata?: {
      styles?: Array<{
        id: string;
        fill?: { backgroundColor?: string };
        font?: {
          family?: string;
          size?: number;
          bold?: boolean;
          italic?: boolean;
          underline?: boolean;
          color?: string;
        };
        alignment?: {
          horizontal?: string;
          vertical?: string;
          wrap?: boolean;
          indent?: number;
        };
        borders?: {
          top?: { style?: string; weight?: string; color?: string };
          right?: { style?: string; weight?: string; color?: string };
          bottom?: { style?: string; weight?: string; color?: string };
          left?: { style?: string; weight?: string; color?: string };
        };
      }>;
      formats?: Array<{
        id: string;
        code: string;
        kind: string;
      }>;
    };
  };
  sheets: Array<{
    name: string;
    metadata?: {
      styleRanges?: Array<{ range: CellRangeRef; styleId: string }>;
      formatRanges?: Array<{ range: CellRangeRef; formatId: string }>;
    };
    cells: Array<{ address: string; value?: unknown; formula?: string; format?: string }>;
  }>;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isAgentFrame(value: unknown): value is AgentFrame {
  return (
    isRecord(value) &&
    (value.kind === "request" || value.kind === "response") &&
    (("request" in value && isRecord(value.request)) ||
      ("response" in value && isRecord(value.response)))
  );
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
      throw new Error(
        `Timed out waiting for local-server on ${localServerUrl}${lastError ? `: ${lastError}` : ""}`,
      );
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
  const exited = await Promise.race([exitPromise.then(() => true), delay(5_000).then(() => false)]);
  if (exited) {
    return;
  }

  process.kill("SIGKILL");
  await exitPromise;
}

async function startLocalServer(port: number) {
  const child = spawn(bunExecutable, ["apps/local-server/src/index.ts"], {
    cwd: process.cwd(),
    env: {
      ...process.env,
      HOST: "127.0.0.1",
      PORT: String(port),
    },
    stdio: ["ignore", "pipe", "pipe"],
  });
  const localServerUrl = `http://127.0.0.1:${port}`;

  let logs = "";
  const appendLogChunk = (chunk: Uint8Array | string) => {
    logs += chunk.toString();
    if (logs.length > 12_000) {
      logs = logs.slice(-12_000);
    }
  };
  child.stdout?.on("data", appendLogChunk);
  child.stderr?.on("data", appendLogChunk);

  const exitPromise = new Promise<{ code: number | null; signal: string | null }>((resolve) => {
    child.once("exit", (code, signal) => {
      resolve({ code, signal });
    });
  });

  try {
    await Promise.race([
      waitForLocalServerHealthy(localServerUrl),
      exitPromise.then(({ code, signal }) => {
        throw new Error(
          `local-server exited before becoming healthy (code=${code ?? "null"}, signal=${signal ?? "null"})\n${logs}`,
        );
      }),
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
    getLogs: () => logs,
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

async function sendAgentRequest(
  localServerUrl: string,
  request: AgentRequest,
): Promise<AgentResponse> {
  const response = await fetch(`${localServerUrl}/v1/agent/frames`, {
    method: "POST",
    headers: {
      "content-type": "application/octet-stream",
    },
    body: Buffer.from(
      encodeAgentFrame({
        kind: "request",
        request,
      }),
    ),
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
  callback: (sessionId: string) => Promise<void>,
) {
  const openResponse = await sendAgentRequest(localServerUrl, {
    kind: "openWorkbookSession",
    id: `open:${Date.now()}`,
    documentId,
    replicaId,
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
      sessionId: openResponse.sessionId,
    }).catch(() => undefined);
  }
}

async function exportWorkbookSnapshot(
  localServerUrl: string,
  documentId: string,
): Promise<WorkbookSnapshotLike> {
  let snapshot: WorkbookSnapshotLike | null = null;
  await withAgentSession(
    localServerUrl,
    documentId,
    `playwright-export:${Date.now()}`,
    async (sessionId) => {
      const response = await sendAgentRequest(localServerUrl, {
        kind: "exportSnapshot",
        id: `snapshot:${Date.now()}`,
        sessionId,
      });
      if (response.kind !== "snapshot") {
        throw new Error(`Expected snapshot response, received ${response.kind}`);
      }
      snapshot = response.snapshot;
    },
  );
  if (!snapshot) {
    throw new Error("Agent exportSnapshot returned no snapshot");
  }
  return snapshot;
}

function getSheetSnapshot(snapshot: WorkbookSnapshotLike, sheetName: string) {
  const sheet = snapshot.sheets.find((entry) => entry.name === sheetName);
  if (!sheet) {
    throw new Error(`Missing sheet snapshot for ${sheetName}`);
  }
  return sheet;
}

function getSingleCellStyleRecord(
  snapshot: WorkbookSnapshotLike,
  sheetName: string,
  address: string,
) {
  const sheet = getSheetSnapshot(snapshot, sheetName);
  const styleRange = sheet.metadata?.styleRanges?.find(
    (entry) =>
      entry.range.sheetName === sheetName &&
      entry.range.startAddress === address &&
      entry.range.endAddress === address,
  );
  if (!styleRange) {
    throw new Error(`Missing style range for ${sheetName}!${address}`);
  }
  const styleRecord = snapshot.workbook.metadata?.styles?.find(
    (entry) => entry.id === styleRange.styleId,
  );
  if (!styleRecord) {
    throw new Error(`Missing style record ${styleRange.styleId}`);
  }
  return styleRecord;
}

function getSingleCellFormatRecord(
  snapshot: WorkbookSnapshotLike,
  sheetName: string,
  address: string,
) {
  const sheet = getSheetSnapshot(snapshot, sheetName);
  const formatRange = sheet.metadata?.formatRanges?.find(
    (entry) =>
      entry.range.sheetName === sheetName &&
      entry.range.startAddress === address &&
      entry.range.endAddress === address,
  );
  if (!formatRange) {
    throw new Error(`Missing format range for ${sheetName}!${address}`);
  }
  const formatRecord = snapshot.workbook.metadata?.formats?.find(
    (entry) => entry.id === formatRange.formatId,
  );
  if (!formatRecord) {
    throw new Error(`Missing format record ${formatRange.formatId}`);
  }
  return formatRecord;
}

async function setColorInput(locator: Locator, value: string) {
  await locator.evaluate((element, nextValue) => {
    if (!(element instanceof HTMLInputElement)) {
      throw new Error("color input is not an HTMLInputElement");
    }
    const input = element;
    const descriptor = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value");
    descriptor?.set?.call(input, String(nextValue));
    input.dispatchEvent(new Event("input", { bubbles: true }));
    input.dispatchEvent(new Event("change", { bubbles: true }));
  }, value);
}

async function expectToolbarColor(locator: Locator, value: string) {
  await expect(locator).toHaveAttribute("data-current-color", value.toLowerCase());
}

async function setToolbarCustomColor(
  page: Parameters<typeof test>[0]["page"],
  controlLabel: "Fill color" | "Text color",
  value: string,
) {
  await page.getByLabel(controlLabel).click();
  await page.getByLabel(`Open custom ${controlLabel.toLowerCase()} picker`).click();
  await setColorInput(
    page.getByLabel(controlLabel === "Fill color" ? "Custom fill color" : "Custom text color", {
      exact: true,
    }),
    value,
  );
}

async function pickToolbarPresetColor(
  page: Parameters<typeof test>[0]["page"],
  controlLabel: "Fill color" | "Text color",
  swatchLabel: string,
) {
  await page.getByLabel(controlLabel).click();
  await page.getByLabel(`${controlLabel} ${swatchLabel}`).click();
}

const TOOLBAR_ACTION_CELL = {
  sheetName: "Sheet1",
  address: "B2",
  startAddress: "B2",
  endAddress: "B2",
  columnIndex: 1,
  rowIndex: 1,
} as const;

function findSingleCellStyleRange(
  snapshot: WorkbookSnapshotLike,
  sheetName: string,
  address: string,
) {
  const sheet = getSheetSnapshot(snapshot, sheetName);
  return sheet.metadata?.styleRanges?.find(
    (entry) =>
      entry.range.sheetName === sheetName &&
      entry.range.startAddress === address &&
      entry.range.endAddress === address,
  );
}

function findSingleCellFormatRange(
  snapshot: WorkbookSnapshotLike,
  sheetName: string,
  address: string,
) {
  const sheet = getSheetSnapshot(snapshot, sheetName);
  return sheet.metadata?.formatRanges?.find(
    (entry) =>
      entry.range.sheetName === sheetName &&
      entry.range.startAddress === address &&
      entry.range.endAddress === address,
  );
}

function getSingleCellStyleRecordOrNull(
  snapshot: WorkbookSnapshotLike,
  sheetName: string,
  address: string,
) {
  const styleRange = findSingleCellStyleRange(snapshot, sheetName, address);
  if (!styleRange) {
    return null;
  }
  return (
    snapshot.workbook.metadata?.styles?.find((entry) => entry.id === styleRange.styleId) ?? null
  );
}

function getStyleRecordAtCell(snapshot: WorkbookSnapshotLike, sheetName: string, address: string) {
  const sheet = getSheetSnapshot(snapshot, sheetName);
  const parsed = parseTestCellAddress(address);
  const styleRange = [...(sheet.metadata?.styleRanges ?? [])].toReversed().find((entry) => {
    if (entry.range.sheetName !== sheetName) {
      return false;
    }
    const start = parseTestCellAddress(entry.range.startAddress);
    const end = parseTestCellAddress(entry.range.endAddress);
    const startRow = Math.min(start.row, end.row);
    const endRow = Math.max(start.row, end.row);
    const startCol = Math.min(start.col, end.col);
    const endCol = Math.max(start.col, end.col);
    return (
      parsed.row >= startRow &&
      parsed.row <= endRow &&
      parsed.col >= startCol &&
      parsed.col <= endCol
    );
  });
  if (!styleRange) {
    return null;
  }
  return (
    snapshot.workbook.metadata?.styles?.find((entry) => entry.id === styleRange.styleId) ?? null
  );
}

function getSingleCellFormatRecordOrNull(
  snapshot: WorkbookSnapshotLike,
  sheetName: string,
  address: string,
) {
  const formatRange = findSingleCellFormatRange(snapshot, sheetName, address);
  if (!formatRange) {
    return null;
  }
  return (
    snapshot.workbook.metadata?.formats?.find((entry) => entry.id === formatRange.formatId) ?? null
  );
}

async function waitForBrowserSession(localServerUrl: string, documentId: string) {
  await expect
    .poll(
      async () => {
        const documentState = await fetchDocumentState(localServerUrl, documentId);
        return documentState.browserSessions.length > 0;
      },
      {
        message: "browser should attach to the local-server document session",
      },
    )
    .toBe(true);
}

async function prepareToolbarActionDocument(
  page: Parameters<typeof test>[0]["page"],
  localServerUrl: string,
  documentId: string,
  options?: {
    value?: unknown;
    setup?: (sessionId: string) => Promise<void>;
  },
) {
  await withAgentSession(
    localServerUrl,
    documentId,
    `playwright-toolbar-action:${Date.now()}`,
    async (sessionId) => {
      if (options && "value" in options) {
        await sendAgentRequest(localServerUrl, {
          kind: "writeRange",
          id: `write:${Date.now()}`,
          sessionId,
          range: {
            sheetName: TOOLBAR_ACTION_CELL.sheetName,
            startAddress: TOOLBAR_ACTION_CELL.startAddress,
            endAddress: TOOLBAR_ACTION_CELL.endAddress,
          },
          values: [[options.value ?? null]],
        });
      }
      if (options?.setup) {
        await options.setup(sessionId);
      }
    },
  );

  await page.goto(
    `/?document=${encodeURIComponent(documentId)}&server=${encodeURIComponent(localServerUrl)}`,
  );
  await waitForBrowserSession(localServerUrl, documentId);
  await clickProductCell(page, TOOLBAR_ACTION_CELL.columnIndex, TOOLBAR_ACTION_CELL.rowIndex);
  await expect(page.getByTestId("status-selection")).toHaveText(
    `${TOOLBAR_ACTION_CELL.sheetName}!${TOOLBAR_ACTION_CELL.address}`,
  );
}

function createToolbarActionDocumentId(actionName: string) {
  return `playwright-toolbar-action-${actionName}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

async function expectToolbarSnapshotProjection(
  localServerUrl: string,
  documentId: string,
  projector: (snapshot: WorkbookSnapshotLike) => unknown,
  expected: unknown,
) {
  await expect
    .poll(async () => projector(await exportWorkbookSnapshot(localServerUrl, documentId)), {
      message: `snapshot for ${documentId} should match expected toolbar action result`,
    })
    .toEqual(expected);
}

function isLocalDocumentStateSummary(value: unknown): value is LocalDocumentStateSummary {
  return (
    isRecord(value) &&
    typeof value.documentId === "string" &&
    typeof value.cursor === "number" &&
    Array.isArray(value.browserSessions) &&
    Array.isArray(value.agentSessions) &&
    (typeof value.lastBatchId === "string" || value.lastBatchId === null)
  );
}

async function fetchDocumentState(
  localServerUrl: string,
  documentId: string,
): Promise<LocalDocumentStateSummary> {
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
  const entries = Object.entries(parsed).filter(
    (entry): entry is [string, number] => typeof entry[1] === "number",
  );
  return Object.fromEntries(entries);
}

async function getProductColumnWidth(
  page: Parameters<typeof test>[0]["page"],
  columnIndex: number,
) {
  const grid = page.getByTestId("sheet-grid");
  const [defaultWidthRaw, overridesRaw] = await Promise.all([
    grid.getAttribute("data-default-column-width"),
    grid.getAttribute("data-column-width-overrides"),
  ]);
  const defaultWidth = Number(defaultWidthRaw ?? String(PRODUCT_COLUMN_WIDTH));
  const overrides = parseColumnWidthOverrides(overridesRaw);
  return overrides[String(columnIndex)] ?? defaultWidth;
}

async function getProductColumnLeft(page: Parameters<typeof test>[0]["page"], columnIndex: number) {
  const widths = await Promise.all(
    Array.from({ length: columnIndex }, (_, index) => getProductColumnWidth(page, index)),
  );
  return PRODUCT_ROW_MARKER_WIDTH + widths.reduce((total, width) => total + width, 0);
}

async function dragProductColumnResize(
  page: Parameters<typeof test>[0]["page"],
  columnIndex: number,
  deltaX: number,
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
  columnIndex: number,
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
  endIndex: number,
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
  const startX =
    axis === "column"
      ? grid.x + startColumnLeft + Math.floor(startColumnWidth / 2)
      : grid.x + Math.floor(PRODUCT_ROW_MARKER_WIDTH / 2);
  const startY =
    axis === "column"
      ? grid.y + Math.floor(PRODUCT_HEADER_HEIGHT / 2)
      : grid.y +
        PRODUCT_HEADER_HEIGHT +
        startIndex * PRODUCT_ROW_HEIGHT +
        Math.floor(PRODUCT_ROW_HEIGHT / 2);
  const endX =
    axis === "column"
      ? grid.x + endColumnLeft + Math.floor(endColumnWidth / 2)
      : grid.x + Math.floor(PRODUCT_ROW_MARKER_WIDTH / 2);
  const endY =
    axis === "column"
      ? grid.y + Math.floor(PRODUCT_HEADER_HEIGHT / 2)
      : grid.y +
        PRODUCT_HEADER_HEIGHT +
        endIndex * PRODUCT_ROW_HEIGHT +
        Math.floor(PRODUCT_ROW_HEIGHT / 2);

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
  const y =
    grid.y +
    PRODUCT_HEADER_HEIGHT +
    rowIndex * PRODUCT_ROW_HEIGHT +
    Math.floor(PRODUCT_ROW_HEIGHT / 2);
  await page.mouse.click(x, y);
}

async function dragProductFillHandle(
  page: Parameters<typeof test>[0]["page"],
  sourceCol: number,
  sourceRow: number,
  targetCol: number,
  targetRow: number,
) {
  const gridLocator = page.getByTestId("sheet-grid");
  await expect(gridLocator).toBeVisible();
  const grid = await gridLocator.boundingBox();
  if (!grid) {
    throw new Error("sheet grid is not visible");
  }

  const sourceLeft = grid.x + (await getProductColumnLeft(page, sourceCol));
  const sourceTop = grid.y + PRODUCT_HEADER_HEIGHT + sourceRow * PRODUCT_ROW_HEIGHT;
  const targetLeft = grid.x + (await getProductColumnLeft(page, targetCol));
  const targetTop = grid.y + PRODUCT_HEADER_HEIGHT + targetRow * PRODUCT_ROW_HEIGHT;
  const sourceWidth = await getProductColumnWidth(page, sourceCol);
  const targetWidth = await getProductColumnWidth(page, targetCol);

  await page.mouse.move(sourceLeft + sourceWidth - 3, sourceTop + PRODUCT_ROW_HEIGHT - 3);
  await page.mouse.down();
  await page.mouse.move(targetLeft + targetWidth - 3, targetTop + PRODUCT_ROW_HEIGHT - 3, {
    steps: 10,
  });
  await page.mouse.up();
}

async function clickProductBodyOffset(
  page: Parameters<typeof test>[0]["page"],
  offsetX: number,
  rowIndex = 0,
) {
  const gridLocator = page.getByTestId("sheet-grid");
  await expect(gridLocator).toBeVisible();
  const grid = await gridLocator.boundingBox();
  if (!grid) {
    throw new Error("sheet grid is not visible");
  }

  await page.mouse.click(
    grid.x + PRODUCT_ROW_MARKER_WIDTH + offsetX,
    grid.y +
      PRODUCT_HEADER_HEIGHT +
      rowIndex * PRODUCT_ROW_HEIGHT +
      Math.floor(PRODUCT_ROW_HEIGHT / 2),
  );
}

async function getBox(locator: Locator) {
  const box = await locator.boundingBox();
  if (!box) {
    throw new Error("locator is not visible");
  }
  return box;
}

async function getGridBox(page: Parameters<typeof test>[0]["page"]) {
  return await getBox(page.getByTestId("sheet-grid"));
}

function getProductCellClientPoint(
  grid: { x: number; y: number; width: number; height: number },
  columnIndex: number,
  rowIndex: number,
  xRatio = 0.5,
  yRatio = 0.5,
) {
  return {
    x:
      grid.x +
      PRODUCT_ROW_MARKER_WIDTH +
      columnIndex * PRODUCT_COLUMN_WIDTH +
      PRODUCT_COLUMN_WIDTH * xRatio,
    y: grid.y + PRODUCT_HEADER_HEIGHT + rowIndex * PRODUCT_ROW_HEIGHT + PRODUCT_ROW_HEIGHT * yRatio,
  };
}

async function sampleGridPixel(
  page: Parameters<typeof test>[0]["page"],
  x: number,
  y: number,
): Promise<{ r: number; g: number; b: number; a: number }> {
  const sampled = await page.evaluate(
    ({ x: clientX, y: clientY }) => {
      const canvases = [
        ...document.querySelectorAll<HTMLCanvasElement>('[data-testid="sheet-grid"] canvas'),
      ];
      for (let index = canvases.length - 1; index >= 0; index -= 1) {
        const canvas = canvases[index];
        const rect = canvas.getBoundingClientRect();
        if (
          clientX < rect.left ||
          clientX > rect.right ||
          clientY < rect.top ||
          clientY > rect.bottom
        ) {
          continue;
        }
        const context = canvas.getContext("2d");
        if (!context) {
          continue;
        }
        const scaleX = canvas.width / rect.width;
        const scaleY = canvas.height / rect.height;
        const sampleX = Math.max(
          0,
          Math.min(canvas.width - 1, Math.round((clientX - rect.left) * scaleX)),
        );
        const sampleY = Math.max(
          0,
          Math.min(canvas.height - 1, Math.round((clientY - rect.top) * scaleY)),
        );
        const [r, g, b, a] = context.getImageData(sampleX, sampleY, 1, 1).data;
        return { r, g, b, a };
      }
      return null;
    },
    { x, y },
  );
  if (!sampled) {
    throw new Error(`Unable to sample grid pixel at ${x}, ${y}`);
  }
  return sampled;
}

async function findClosestGridColorDistance(
  page: Parameters<typeof test>[0]["page"],
  x: number,
  y: number,
  target: { r: number; g: number; b: number },
  radius = 4,
): Promise<number> {
  const result = await page.evaluate(
    ({ clientX, clientY, targetColor, radiusPx }) => {
      const canvases = [
        ...document.querySelectorAll<HTMLCanvasElement>('[data-testid="sheet-grid"] canvas'),
      ];
      for (let index = canvases.length - 1; index >= 0; index -= 1) {
        const canvas = canvases[index];
        const rect = canvas.getBoundingClientRect();
        if (
          clientX < rect.left ||
          clientX > rect.right ||
          clientY < rect.top ||
          clientY > rect.bottom
        ) {
          continue;
        }
        const context = canvas.getContext("2d");
        if (!context) {
          continue;
        }
        const scaleX = canvas.width / rect.width;
        const scaleY = canvas.height / rect.height;
        const centerX = Math.max(
          0,
          Math.min(canvas.width - 1, Math.round((clientX - rect.left) * scaleX)),
        );
        const centerY = Math.max(
          0,
          Math.min(canvas.height - 1, Math.round((clientY - rect.top) * scaleY)),
        );
        let closest = Number.POSITIVE_INFINITY;
        for (let dx = -radiusPx; dx <= radiusPx; dx += 1) {
          for (let dy = -radiusPx; dy <= radiusPx; dy += 1) {
            const sampleX = Math.max(0, Math.min(canvas.width - 1, centerX + dx));
            const sampleY = Math.max(0, Math.min(canvas.height - 1, centerY + dy));
            const [r, g, b] = context.getImageData(sampleX, sampleY, 1, 1).data;
            const distance = Math.sqrt(
              (r - targetColor.r) ** 2 + (g - targetColor.g) ** 2 + (b - targetColor.b) ** 2,
            );
            if (distance < closest) {
              closest = distance;
            }
          }
        }
        return closest;
      }
      return null;
    },
    { clientX: x, clientY: y, targetColor: target, radiusPx: radius },
  );
  if (result === null) {
    throw new Error(`Unable to inspect grid color near ${x}, ${y}`);
  }
  return result;
}

function colorDistance(
  left: { r: number; g: number; b: number },
  right: { r: number; g: number; b: number },
) {
  return Math.sqrt((left.r - right.r) ** 2 + (left.g - right.g) ** 2 + (left.b - right.b) ** 2);
}

async function findClosestBorderOverlayDistance(
  page: Parameters<typeof test>[0]["page"],
  x: number,
  y: number,
): Promise<number> {
  const distance = await page.evaluate(
    ({ clientX, clientY }) => {
      const nodes = [
        ...document.querySelectorAll<HTMLElement>('[data-testid="grid-border-overlay-segment"]'),
      ];
      if (nodes.length === 0) {
        return null;
      }

      let closest = Number.POSITIVE_INFINITY;
      for (const node of nodes) {
        const rect = node.getBoundingClientRect();
        const dx =
          clientX < rect.left
            ? rect.left - clientX
            : clientX > rect.right
              ? clientX - rect.right
              : 0;
        const dy =
          clientY < rect.top
            ? rect.top - clientY
            : clientY > rect.bottom
              ? clientY - rect.bottom
              : 0;
        closest = Math.min(closest, Math.hypot(dx, dy));
      }

      return closest;
    },
    { clientX: x, clientY: y },
  );

  if (distance === null) {
    throw new Error(`No rendered border overlay segments found near ${x}, ${y}`);
  }

  return distance;
}

async function expectRenderedBorderNearPoint(
  page: Parameters<typeof test>[0]["page"],
  x: number,
  y: number,
  tolerance = 5,
) {
  const distance = await findClosestBorderOverlayDistance(page, x, y);
  expect(distance).toBeLessThanOrEqual(tolerance);
}

async function borderVisibilityMatches(
  page: Parameters<typeof test>[0]["page"],
  checks: readonly {
    x: number;
    y: number;
    mode: "present" | "absent";
    tolerance?: number;
  }[],
): Promise<boolean> {
  const results = await Promise.all(
    checks.map(async (check) => {
      const distance = await findClosestBorderOverlayDistance(page, check.x, check.y);
      const tolerance = check.tolerance ?? (check.mode === "present" ? 5 : 6);
      return check.mode === "present" ? distance <= tolerance : distance > tolerance;
    }),
  );
  return results.every(Boolean);
}

async function waitForRenderedBorders(page: Parameters<typeof test>[0]["page"], minimumCount = 1) {
  await expect
    .poll(async () => await page.getByTestId("grid-border-overlay-segment").count(), {
      message: `Expected at least ${minimumCount} rendered border overlay segments`,
    })
    .toBeGreaterThan(minimumCount - 1);
}

async function expectAllBordersVisibleForRange(
  page: Parameters<typeof test>[0]["page"],
  startColumnIndex: number,
  startRowIndex: number,
  endColumnIndex: number,
  endRowIndex: number,
) {
  await waitForRenderedBorders(page);
  const grid = await getGridBox(page);
  const topLeft = getProductCellClientPoint(grid, startColumnIndex, startRowIndex, 0.03, 0.5);
  const top = getProductCellClientPoint(grid, startColumnIndex, startRowIndex, 0.5, 0.08);
  const innerVertical = getProductCellClientPoint(grid, startColumnIndex, startRowIndex, 0.97, 0.5);
  const innerHorizontal = getProductCellClientPoint(
    grid,
    startColumnIndex,
    startRowIndex,
    0.5,
    0.92,
  );
  const right = getProductCellClientPoint(grid, endColumnIndex, startRowIndex, 0.97, 0.5);
  const bottom = getProductCellClientPoint(grid, startColumnIndex, endRowIndex, 0.5, 0.92);

  await Promise.all([
    expectRenderedBorderNearPoint(page, topLeft.x, topLeft.y),
    expectRenderedBorderNearPoint(page, top.x, top.y),
    expectRenderedBorderNearPoint(page, innerVertical.x, innerVertical.y),
    expectRenderedBorderNearPoint(page, innerHorizontal.x, innerHorizontal.y),
    expectRenderedBorderNearPoint(page, right.x, right.y),
    expectRenderedBorderNearPoint(page, bottom.x, bottom.y),
  ]);
}

async function clickProductCell(
  page: Parameters<typeof test>[0]["page"],
  columnIndex: number,
  rowIndex: number,
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
    grid.y +
      PRODUCT_HEADER_HEIGHT +
      rowIndex * PRODUCT_ROW_HEIGHT +
      Math.floor(PRODUCT_ROW_HEIGHT / 2),
  );
}

async function clickProductCellUpperHalf(
  page: Parameters<typeof test>[0]["page"],
  columnIndex: number,
  rowIndex: number,
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
    grid.y + PRODUCT_HEADER_HEIGHT + rowIndex * PRODUCT_ROW_HEIGHT + 4,
  );
}

async function clickProductSelectedCellTopBorder(
  page: Parameters<typeof test>[0]["page"],
  columnIndex: number,
  rowIndex: number,
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
    grid.y + PRODUCT_HEADER_HEIGHT + rowIndex * PRODUCT_ROW_HEIGHT - 1,
  );
}

async function dragProductBodySelection(
  page: Parameters<typeof test>[0]["page"],
  startColumn: number,
  startRow: number,
  endColumn: number,
  endRow: number,
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
  const startY =
    grid.y +
    PRODUCT_HEADER_HEIGHT +
    startRow * PRODUCT_ROW_HEIGHT +
    Math.floor(PRODUCT_ROW_HEIGHT / 2);
  const endX = grid.x + endLeft + Math.floor(endWidth / 2);
  const endY =
    grid.y +
    PRODUCT_HEADER_HEIGHT +
    endRow * PRODUCT_ROW_HEIGHT +
    Math.floor(PRODUCT_ROW_HEIGHT / 2);

  await page.mouse.move(startX, startY);
  await page.mouse.down();
  await page.mouse.move(endX, endY, { steps: 12 });
  await page.mouse.up();
}

test("web app renders the minimal product shell without legacy demo chrome", async ({ page }) => {
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

test("web app keeps toolbar controls aligned and consistently sized", async ({ page }) => {
  await page.goto("/");

  const toolbar = page.getByRole("toolbar", { name: "Formatting toolbar" });
  await expect(toolbar).toBeVisible();

  const controls = [
    page.getByLabel("Number format"),
    page.getByLabel("Decrease decimals"),
    page.getByLabel("Font family"),
    page.getByLabel("Font size"),
    page.getByLabel("Bold"),
    page.getByLabel("Italic"),
    page.getByLabel("Underline"),
    page.getByLabel("Fill color"),
    page.getByLabel("Text color"),
    page.getByLabel("Align left"),
    page.getByLabel("Align center"),
    page.getByLabel("Align right"),
    page.getByLabel("Bottom border"),
    page.getByLabel("All borders"),
    page.getByLabel("No borders"),
    page.getByLabel("Wrap"),
    page.getByLabel("Clear style"),
  ];

  const metrics = await Promise.all(
    controls.map(async (locator) => {
      const label =
        (await locator.getAttribute("aria-label")) ??
        (await locator.evaluate((element) => element.textContent?.trim() ?? "")) ??
        "unknown";
      const box = await getBox(locator);
      return {
        label,
        x: Math.round(box.x),
        y: Math.round(box.y),
        width: Math.round(box.width),
        height: Math.round(box.height),
      };
    }),
  );
  const boxes = metrics.map(({ x, y, width, height }) => ({ x, y, width, height }));
  const heights = boxes.map((box) => Math.round(box.height));
  const tops = boxes.map((box) => Math.round(box.y));
  const bottoms = boxes.map((box) => Math.round(box.y + box.height));

  const heightDelta = Math.max(...heights) - Math.min(...heights);
  const topDelta = Math.max(...tops) - Math.min(...tops);
  const bottomDelta = Math.max(...bottoms) - Math.min(...bottoms);
  if (heightDelta > 1 || topDelta > 1 || bottomDelta > 1) {
    throw new Error(
      `Toolbar geometry mismatch (height=${heightDelta}, top=${topDelta}, bottom=${bottomDelta}): ${JSON.stringify(metrics)}`,
    );
  }

  const toolbarBox = await getBox(toolbar);
  expect(toolbarBox.height).toBeLessThanOrEqual(48);
});

test("web app keeps toolbar, formula bar, grid, and footer tightly stacked", async ({ page }) => {
  await page.goto("/");

  const toolbar = page.getByRole("toolbar", { name: "Formatting toolbar" });
  const formulaBar = page.getByTestId("formula-bar");
  const grid = page.getByTestId("sheet-grid");
  const sheetTab = page.getByRole("tab", { name: "Sheet1" });

  const [toolbarBox, formulaBarBox, gridBox, sheetTabBox] = await Promise.all([
    getBox(toolbar),
    getBox(formulaBar),
    getBox(grid),
    getBox(sheetTab),
  ]);

  expect(Math.abs(formulaBarBox.y - (toolbarBox.y + toolbarBox.height))).toBeLessThanOrEqual(2);
  expect(Math.abs(gridBox.y - (formulaBarBox.y + formulaBarBox.height))).toBeLessThanOrEqual(2);
  expect(gridBox.height).toBeGreaterThan(300);
  expect(sheetTabBox.y).toBeGreaterThan(gridBox.y + gridBox.height - 40);
});

test("web app keeps formula bar controls aligned and consistently sized", async ({ page }) => {
  await page.goto("/");

  const nameBox = page.getByTestId("name-box");
  const formulaFrame = page.getByTestId("formula-input-frame");

  const [nameBoxBox, formulaFrameBox] = await Promise.all([getBox(nameBox), getBox(formulaFrame)]);

  expect(Math.abs(nameBoxBox.height - formulaFrameBox.height)).toBeLessThanOrEqual(1);
  expect(Math.abs(nameBoxBox.y - formulaFrameBox.y)).toBeLessThanOrEqual(1);
  expect(
    Math.abs(nameBoxBox.y + nameBoxBox.height - (formulaFrameBox.y + formulaFrameBox.height)),
  ).toBeLessThanOrEqual(1);
});

test("web app keeps the toolbar compact on narrow viewports", async ({ page }) => {
  await page.setViewportSize({ width: 620, height: 760 });
  await page.goto("/");

  const toolbar = page.getByRole("toolbar", { name: "Formatting toolbar" });
  const firstControl = page.getByLabel("Number format");
  const lastControl = page.getByLabel("Clear style");
  await expect(toolbar).toBeVisible();
  const [toolbarBox, firstControlBox, lastControlBox] = await Promise.all([
    getBox(toolbar),
    getBox(firstControl),
    getBox(lastControl),
  ]);

  expect(toolbarBox.height).toBeLessThanOrEqual(48);
  expect(Math.abs(firstControlBox.y - lastControlBox.y)).toBeLessThanOrEqual(1);
  expect(lastControlBox.y + lastControlBox.height).toBeLessThanOrEqual(
    toolbarBox.y + toolbarBox.height + 1,
  );
});

test("web app shows preset color swatches first and only reveals the custom picker on demand", async ({
  page,
}) => {
  await page.goto("/");

  await page.getByLabel("Fill color").click();
  await expect(page.getByRole("dialog", { name: "Fill color palette" })).toBeVisible();
  await expect(page.getByLabel("Fill color white")).toBeVisible();
  await expect(page.getByLabel("Fill color light cornflower blue 3")).toBeVisible();
  await expect(page.getByLabel("Fill color dark cornflower blue 3")).toBeVisible();
  await expect(page.getByLabel("Fill color theme cornflower blue")).toBeVisible();
  await expect(page.getByLabel("Custom fill color", { exact: true })).toHaveCount(0);

  await page.getByLabel("Open custom fill color picker").click();
  await expect(page.getByLabel("Custom fill color", { exact: true })).toBeVisible();
});

test("web app renders the fill color palette as a visible popover below the toolbar", async ({
  page,
}) => {
  await page.goto("/");

  await page.getByLabel("Fill color").click();

  const toolbar = page.getByRole("toolbar", { name: "Formatting toolbar" });
  const palette = page.getByRole("dialog", { name: "Fill color palette" });
  const swatch = page.getByLabel("Fill color light cornflower blue 3");
  const [toolbarBox, paletteBox, swatchBox] = await Promise.all([
    getBox(toolbar),
    getBox(palette),
    getBox(swatch),
  ]);

  expect(paletteBox.y).toBeGreaterThanOrEqual(toolbarBox.y + toolbarBox.height - 1);
  expect(paletteBox.height).toBeGreaterThan(120);
  expect(paletteBox.width).toBeGreaterThan(200);
  expect(swatchBox.y + swatchBox.height).toBeLessThanOrEqual(paletteBox.y + paletteBox.height);
  await expect(page.getByText("Standard")).toBeVisible();
});

test("web app applies preset swatch colors directly from the palette", async ({ page }) => {
  await page.goto("/");

  await pickToolbarPresetColor(page, "Fill color", "light cornflower blue 3");
  await expectToolbarColor(page.getByLabel("Fill color"), "#c9daf8");

  await pickToolbarPresetColor(page, "Text color", "dark blue 1");
  await expectToolbarColor(page.getByLabel("Text color"), "#3d85c6");
});

test("web app keeps preset fill color visible after clicking another cell", async ({ page }) => {
  await page.goto("/");

  const grid = await getGridBox(page);
  const center = getProductCellClientPoint(grid, 0, 0, 0.55, 0.55);
  const target = { r: 201, g: 218, b: 248 };

  await pickToolbarPresetColor(page, "Fill color", "light cornflower blue 3");
  await expectToolbarColor(page.getByLabel("Fill color"), "#c9daf8");

  await expect
    .poll(async () => await findClosestGridColorDistance(page, center.x, center.y, target, 10))
    .toBeLessThan(30);

  await clickProductCell(page, 9, 23);

  await expect
    .poll(async () => await findClosestGridColorDistance(page, center.x, center.y, target, 10))
    .toBeLessThan(30);
});

test("web app hydrates toolbar controls from the selected cell style and number format", async ({
  page,
}) => {
  test.slow();
  const port = await reserveLocalPort();
  const documentId = `playwright-toolbar-hydration-${Date.now()}`;
  const localServer = await startLocalServer(port);

  try {
    await withAgentSession(
      localServer.localServerUrl,
      documentId,
      `playwright-agent:${Date.now()}`,
      async (sessionId) => {
        await sendAgentRequest(localServer.localServerUrl, {
          kind: "writeRange",
          id: `write:${Date.now()}`,
          sessionId,
          range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B1" },
          values: [[1234.5, 88]],
        });
        await sendAgentRequest(localServer.localServerUrl, {
          kind: "setRangeStyle",
          id: `style:${Date.now()}`,
          sessionId,
          range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "A1" },
          patch: {
            fill: { backgroundColor: "#dbeafe" },
            font: {
              family: "Georgia",
              size: 14,
              bold: true,
              color: "#7c2d12",
            },
            alignment: { horizontal: "right", wrap: true },
          },
        });
        await sendAgentRequest(localServer.localServerUrl, {
          kind: "setRangeNumberFormat",
          id: `format:${Date.now()}`,
          sessionId,
          range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "A1" },
          format: {
            kind: "accounting",
            currency: "USD",
            decimals: 2,
            useGrouping: true,
            negativeStyle: "parentheses",
            zeroStyle: "dash",
          },
        });
      },
    );

    await page.goto(
      `/?document=${encodeURIComponent(documentId)}&server=${encodeURIComponent(localServer.localServerUrl)}`,
    );

    await expect
      .poll(
        async () => {
          const documentState = await fetchDocumentState(localServer.localServerUrl, documentId);
          return documentState.browserSessions.length > 0;
        },
        { message: "browser should attach to the local-server document session" },
      )
      .toBe(true);

    await clickProductCell(page, 0, 0);
    await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A1");
    await expect(page.getByLabel("Number format")).toHaveValue("accounting");
    await expect(page.getByLabel("Font family")).toHaveValue("Georgia");
    await expect(page.getByLabel("Font size")).toHaveValue("14");
    await expectToolbarColor(page.getByLabel("Fill color"), "#dbeafe");
    await expectToolbarColor(page.getByLabel("Text color"), "#7c2d12");
    await expect(page.getByLabel("Wrap")).toHaveAttribute("aria-pressed", "true");
    await expect(page.getByLabel("Bold")).toHaveClass(/bg-\[#e6f4ea\]/);
    await expect(page.getByLabel("Align right")).toHaveClass(/bg-\[#e6f4ea\]/);

    await clickProductCell(page, 1, 0);
    await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!B1");
    await expect(page.getByLabel("Number format")).toHaveValue("general");
    await expect(page.getByLabel("Font family")).toHaveValue("");
    await expect(page.getByLabel("Font size")).toHaveValue("11");
    await expect(page.getByLabel("Wrap")).toHaveAttribute("aria-pressed", "false");
    await expect(page.getByLabel("Bold")).not.toHaveClass(/bg-\[#e6f4ea\]/);
    await expect(page.getByLabel("Align right")).not.toHaveClass(/bg-\[#e6f4ea\]/);
  } catch (error) {
    throw new Error(
      `${error instanceof Error ? error.message : String(error)}\nLocal-server logs:\n${localServer.getLogs()}`,
      { cause: error },
    );
  } finally {
    await localServer.stop();
  }
});

test("web app persists toolbar formatting actions to the synced workbook snapshot", async ({
  page,
}) => {
  test.slow();
  const port = await reserveLocalPort();
  const documentId = `playwright-toolbar-persist-${Date.now()}`;
  const localServer = await startLocalServer(port);

  try {
    await withAgentSession(
      localServer.localServerUrl,
      documentId,
      `playwright-agent:${Date.now()}`,
      async (sessionId) => {
        await sendAgentRequest(localServer.localServerUrl, {
          kind: "writeRange",
          id: `write:${Date.now()}`,
          sessionId,
          range: { sheetName: "Sheet1", startAddress: "B2", endAddress: "B2" },
          values: [[1234.5]],
        });
      },
    );

    await page.goto(
      `/?document=${encodeURIComponent(documentId)}&server=${encodeURIComponent(localServer.localServerUrl)}`,
    );

    await expect
      .poll(
        async () => {
          const documentState = await fetchDocumentState(localServer.localServerUrl, documentId);
          return documentState.browserSessions.length > 0;
        },
        { message: "browser should attach to the local-server document session" },
      )
      .toBe(true);

    await clickProductCell(page, 1, 1);
    await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!B2");

    await page.getByLabel("Number format").selectOption("accounting");
    await page.getByLabel("Font family").selectOption("Georgia");
    await page.getByLabel("Font size").selectOption("14");
    await page.getByLabel("Bold").click();
    await page.getByLabel("Italic").click();
    await page.getByLabel("Underline").click();
    await setToolbarCustomColor(page, "Fill color", "#dbeafe");
    await setToolbarCustomColor(page, "Text color", "#7c2d12");
    await page.getByLabel("Align right").click();
    await page.getByLabel("All borders").click();
    await page.getByLabel("Wrap").click();

    await expect(page.getByLabel("Number format")).toHaveValue("accounting");
    await expect(page.getByLabel("Font family")).toHaveValue("Georgia");
    await expect(page.getByLabel("Font size")).toHaveValue("14");
    await expectToolbarColor(page.getByLabel("Fill color"), "#dbeafe");
    await expectToolbarColor(page.getByLabel("Text color"), "#7c2d12");
    await expect(page.getByLabel("Wrap")).toHaveAttribute("aria-pressed", "true");

    await expect
      .poll(async () => {
        const snapshot = await exportWorkbookSnapshot(localServer.localServerUrl, documentId);
        const style = getSingleCellStyleRecord(snapshot, "Sheet1", "B2");
        const format = getSingleCellFormatRecord(snapshot, "Sheet1", "B2");
        return {
          fontFamily: style.font?.family ?? null,
          fontSize: style.font?.size ?? null,
          bold: style.font?.bold ?? false,
          italic: style.font?.italic ?? false,
          underline: style.font?.underline ?? false,
          fontColor: style.font?.color ?? null,
          fillColor: style.fill?.backgroundColor ?? null,
          horizontal: style.alignment?.horizontal ?? null,
          wrap: style.alignment?.wrap ?? false,
          borderCount: BORDER_SIDES.filter((side) => Boolean(style.borders?.[side])).length,
          formatKind: format.kind,
        };
      })
      .toEqual({
        fontFamily: "Georgia",
        fontSize: 14,
        bold: true,
        italic: true,
        underline: true,
        fontColor: "#7c2d12",
        fillColor: "#dbeafe",
        horizontal: "right",
        wrap: true,
        borderCount: 4,
        formatKind: "accounting",
      });
  } catch (error) {
    throw new Error(
      `${error instanceof Error ? error.message : String(error)}\nLocal-server logs:\n${localServer.getLogs()}`,
      { cause: error },
    );
  } finally {
    await localServer.stop();
  }
});

test("web app clears style and number format back to workbook defaults", async ({ page }) => {
  test.slow();
  const port = await reserveLocalPort();
  const documentId = `playwright-toolbar-clear-${Date.now()}`;
  const localServer = await startLocalServer(port);

  try {
    await withAgentSession(
      localServer.localServerUrl,
      documentId,
      `playwright-agent:${Date.now()}`,
      async (sessionId) => {
        await sendAgentRequest(localServer.localServerUrl, {
          kind: "writeRange",
          id: `write:${Date.now()}`,
          sessionId,
          range: { sheetName: "Sheet1", startAddress: "C3", endAddress: "C3" },
          values: [[8840]],
        });
      },
    );

    await page.goto(
      `/?document=${encodeURIComponent(documentId)}&server=${encodeURIComponent(localServer.localServerUrl)}`,
    );

    await expect
      .poll(
        async () => {
          const documentState = await fetchDocumentState(localServer.localServerUrl, documentId);
          return documentState.browserSessions.length > 0;
        },
        { message: "browser should attach to the local-server document session" },
      )
      .toBe(true);

    await clickProductCell(page, 2, 2);
    await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!C3");

    await page.getByLabel("Number format").selectOption("accounting");
    await page.getByLabel("Bold").click();
    await setToolbarCustomColor(page, "Fill color", "#fef3c7");
    await page.getByLabel("Clear style").click();
    await page.getByLabel("Number format").selectOption("general");

    await expect(page.getByLabel("Number format")).toHaveValue("general");
    await expectToolbarColor(page.getByLabel("Fill color"), "#ffffff");
    await expect(page.getByLabel("Bold")).not.toHaveClass(/bg-\[#e6f4ea\]/);

    await expect
      .poll(async () => {
        const snapshot = await exportWorkbookSnapshot(localServer.localServerUrl, documentId);
        const sheet = getSheetSnapshot(snapshot, "Sheet1");
        const hasStyleRange = Boolean(
          sheet.metadata?.styleRanges?.find(
            (entry) =>
              entry.range.sheetName === "Sheet1" &&
              entry.range.startAddress === "C3" &&
              entry.range.endAddress === "C3",
          ),
        );
        const hasFormatRange = Boolean(
          sheet.metadata?.formatRanges?.find(
            (entry) =>
              entry.range.sheetName === "Sheet1" &&
              entry.range.startAddress === "C3" &&
              entry.range.endAddress === "C3",
          ),
        );
        return { hasStyleRange, hasFormatRange };
      })
      .toEqual({ hasStyleRange: false, hasFormatRange: false });
  } catch (error) {
    throw new Error(
      `${error instanceof Error ? error.message : String(error)}\nLocal-server logs:\n${localServer.getLogs()}`,
      { cause: error },
    );
  } finally {
    await localServer.stop();
  }
});

test.describe("web app validates each toolbar formatting action individually", () => {
  let localServer: Awaited<ReturnType<typeof startLocalServer>>;

  test.beforeAll(async () => {
    localServer = await startLocalServer(await reserveLocalPort());
  });

  test.afterAll(async () => {
    await localServer.stop();
  });

  test("number format: general clears the saved format range", async ({ page }) => {
    const documentId = createToolbarActionDocumentId("format-general");
    await prepareToolbarActionDocument(page, localServer.localServerUrl, documentId, {
      value: 1234.5,
      setup: async (sessionId) => {
        await sendAgentRequest(localServer.localServerUrl, {
          kind: "setRangeNumberFormat",
          id: `format:${Date.now()}`,
          sessionId,
          range: {
            sheetName: TOOLBAR_ACTION_CELL.sheetName,
            startAddress: TOOLBAR_ACTION_CELL.startAddress,
            endAddress: TOOLBAR_ACTION_CELL.endAddress,
          },
          format: {
            kind: "accounting",
            currency: "USD",
            decimals: 2,
            useGrouping: true,
            negativeStyle: "parentheses",
            zeroStyle: "dash",
          },
        });
      },
    });

    await expect(page.getByLabel("Number format")).toHaveValue("accounting");
    await page.getByLabel("Number format").selectOption("general");
    await expect(page.getByLabel("Number format")).toHaveValue("general");

    await expectToolbarSnapshotProjection(
      localServer.localServerUrl,
      documentId,
      (snapshot) =>
        findSingleCellFormatRange(
          snapshot,
          TOOLBAR_ACTION_CELL.sheetName,
          TOOLBAR_ACTION_CELL.address,
        ) ?? null,
      null,
    );
  });

  for (const [preset, expected] of [
    ["number", { kind: "number", code: "number:2:1" }],
    ["currency", { kind: "currency", code: "currency:USD:2:1:minus:zero" }],
    ["accounting", { kind: "accounting", code: "accounting:USD:2:1:parentheses:dash" }],
    ["percent", { kind: "percent", code: "percent:2" }],
    ["date", { kind: "date", code: "date:short" }],
    ["text", { kind: "text", code: "text" }],
  ] as const) {
    test(`number format: ${preset} persists the expected format record`, async ({ page }) => {
      const documentId = createToolbarActionDocumentId(`format-${preset}`);
      await prepareToolbarActionDocument(page, localServer.localServerUrl, documentId, {
        value: preset === "date" ? 45292 : 1234.5,
      });

      await page.getByLabel("Number format").selectOption(preset);
      await expect(page.getByLabel("Number format")).toHaveValue(preset);

      await expectToolbarSnapshotProjection(
        localServer.localServerUrl,
        documentId,
        (snapshot) => {
          const format = getSingleCellFormatRecordOrNull(
            snapshot,
            TOOLBAR_ACTION_CELL.sheetName,
            TOOLBAR_ACTION_CELL.address,
          );
          return format ? { kind: format.kind, code: format.code } : null;
        },
        expected,
      );
    });
  }

  test("decrease decimals reduces the saved decimals count", async ({ page }) => {
    const documentId = createToolbarActionDocumentId("decimals-decrease");
    await prepareToolbarActionDocument(page, localServer.localServerUrl, documentId, {
      value: 1234.5,
      setup: async (sessionId) => {
        await sendAgentRequest(localServer.localServerUrl, {
          kind: "setRangeNumberFormat",
          id: `format:${Date.now()}`,
          sessionId,
          range: {
            sheetName: TOOLBAR_ACTION_CELL.sheetName,
            startAddress: TOOLBAR_ACTION_CELL.startAddress,
            endAddress: TOOLBAR_ACTION_CELL.endAddress,
          },
          format: { kind: "number", decimals: 2, useGrouping: true },
        });
      },
    });

    await expect(page.getByLabel("Number format")).toHaveValue("number");
    await page.getByLabel("Decrease decimals").click();

    await expectToolbarSnapshotProjection(
      localServer.localServerUrl,
      documentId,
      (snapshot) => {
        const format = getSingleCellFormatRecordOrNull(
          snapshot,
          TOOLBAR_ACTION_CELL.sheetName,
          TOOLBAR_ACTION_CELL.address,
        );
        return format?.code ?? null;
      },
      "number:1:1",
    );
  });

  test("increase decimals increases the saved decimals count", async ({ page }) => {
    const documentId = createToolbarActionDocumentId("decimals-increase");
    await prepareToolbarActionDocument(page, localServer.localServerUrl, documentId, {
      value: 1234.5,
      setup: async (sessionId) => {
        await sendAgentRequest(localServer.localServerUrl, {
          kind: "setRangeNumberFormat",
          id: `format:${Date.now()}`,
          sessionId,
          range: {
            sheetName: TOOLBAR_ACTION_CELL.sheetName,
            startAddress: TOOLBAR_ACTION_CELL.startAddress,
            endAddress: TOOLBAR_ACTION_CELL.endAddress,
          },
          format: { kind: "number", decimals: 2, useGrouping: true },
        });
      },
    });

    await expect(page.getByLabel("Number format")).toHaveValue("number");
    await page.getByLabel("Increase decimals").click();

    await expectToolbarSnapshotProjection(
      localServer.localServerUrl,
      documentId,
      (snapshot) => {
        const format = getSingleCellFormatRecordOrNull(
          snapshot,
          TOOLBAR_ACTION_CELL.sheetName,
          TOOLBAR_ACTION_CELL.address,
        );
        return format?.code ?? null;
      },
      "number:3:1",
    );
  });

  test("toggle grouping flips grouping in the saved format record", async ({ page }) => {
    const documentId = createToolbarActionDocumentId("toggle-grouping");
    await prepareToolbarActionDocument(page, localServer.localServerUrl, documentId, {
      value: 1234.5,
      setup: async (sessionId) => {
        await sendAgentRequest(localServer.localServerUrl, {
          kind: "setRangeNumberFormat",
          id: `format:${Date.now()}`,
          sessionId,
          range: {
            sheetName: TOOLBAR_ACTION_CELL.sheetName,
            startAddress: TOOLBAR_ACTION_CELL.startAddress,
            endAddress: TOOLBAR_ACTION_CELL.endAddress,
          },
          format: { kind: "number", decimals: 2, useGrouping: true },
        });
      },
    });

    await expect(page.getByLabel("Number format")).toHaveValue("number");
    await page.getByLabel("Toggle grouping").click();

    await expectToolbarSnapshotProjection(
      localServer.localServerUrl,
      documentId,
      (snapshot) => {
        const format = getSingleCellFormatRecordOrNull(
          snapshot,
          TOOLBAR_ACTION_CELL.sheetName,
          TOOLBAR_ACTION_CELL.address,
        );
        return format?.code ?? null;
      },
      "number:2:0",
    );
  });

  for (const [label, value, expectedFamily] of [
    ["Font family", "Georgia", "Georgia"],
    ["Font family", "Courier New", "Courier New"],
  ] as const) {
    test(`font family option ${expectedFamily} persists to the saved style`, async ({ page }) => {
      const documentId = createToolbarActionDocumentId(
        `font-family-${expectedFamily.replace(/\s+/g, "-").toLowerCase()}`,
      );
      await prepareToolbarActionDocument(page, localServer.localServerUrl, documentId, {
        value: "toolbar",
      });

      await page.getByLabel(label).selectOption(value);
      await expect(page.getByLabel(label)).toHaveValue(value);

      await expectToolbarSnapshotProjection(
        localServer.localServerUrl,
        documentId,
        (snapshot) =>
          getSingleCellStyleRecordOrNull(
            snapshot,
            TOOLBAR_ACTION_CELL.sheetName,
            TOOLBAR_ACTION_CELL.address,
          )?.font?.family ?? null,
        expectedFamily,
      );
    });
  }

  test("font size persists to the saved style", async ({ page }) => {
    const documentId = createToolbarActionDocumentId("font-size");
    await prepareToolbarActionDocument(page, localServer.localServerUrl, documentId, {
      value: "toolbar",
    });

    await page.getByLabel("Font size").selectOption("14");
    await expect(page.getByLabel("Font size")).toHaveValue("14");

    await expectToolbarSnapshotProjection(
      localServer.localServerUrl,
      documentId,
      (snapshot) =>
        getSingleCellStyleRecordOrNull(
          snapshot,
          TOOLBAR_ACTION_CELL.sheetName,
          TOOLBAR_ACTION_CELL.address,
        )?.font?.size ?? null,
      14,
    );
  });

  for (const [label, path, expected] of [
    ["Bold", ["font", "bold"], true],
    ["Italic", ["font", "italic"], true],
    ["Underline", ["font", "underline"], true],
    ["Wrap", ["alignment", "wrap"], true],
  ] as const) {
    test(`${label} persists to the saved style`, async ({ page }) => {
      const documentId = createToolbarActionDocumentId(label.toLowerCase().replace(/\s+/g, "-"));
      await prepareToolbarActionDocument(page, localServer.localServerUrl, documentId, {
        value: "toolbar",
      });

      await page.getByLabel(label).click();
      if (label === "Wrap") {
        await expect(page.getByLabel(label)).toHaveAttribute("aria-pressed", "true");
      } else {
        await expect(page.getByLabel(label)).toHaveClass(/bg-\[#e6f4ea\]/);
      }

      await expectToolbarSnapshotProjection(
        localServer.localServerUrl,
        documentId,
        (snapshot) => {
          const style = getSingleCellStyleRecordOrNull(
            snapshot,
            TOOLBAR_ACTION_CELL.sheetName,
            TOOLBAR_ACTION_CELL.address,
          );
          if (!style) {
            return null;
          }
          if (path[0] === "font") {
            return style.font?.[path[1]] ?? null;
          }
          return style.alignment?.[path[1] as "wrap"] ?? null;
        },
        expected,
      );
    });
  }

  for (const [label, color, projector] of [
    [
      "Fill color",
      "#dbeafe",
      (snapshot: WorkbookSnapshotLike) =>
        getSingleCellStyleRecordOrNull(
          snapshot,
          TOOLBAR_ACTION_CELL.sheetName,
          TOOLBAR_ACTION_CELL.address,
        )?.fill?.backgroundColor ?? null,
    ],
    [
      "Text color",
      "#7c2d12",
      (snapshot: WorkbookSnapshotLike) =>
        getSingleCellStyleRecordOrNull(
          snapshot,
          TOOLBAR_ACTION_CELL.sheetName,
          TOOLBAR_ACTION_CELL.address,
        )?.font?.color ?? null,
    ],
  ] as const) {
    test(`${label} persists the selected color`, async ({ page }) => {
      const documentId = createToolbarActionDocumentId(label.toLowerCase().replace(/\s+/g, "-"));
      await prepareToolbarActionDocument(page, localServer.localServerUrl, documentId, {
        value: "toolbar",
      });

      await setToolbarCustomColor(page, label, color);
      await expectToolbarColor(page.getByLabel(label), color);

      await expectToolbarSnapshotProjection(
        localServer.localServerUrl,
        documentId,
        projector,
        color,
      );
    });
  }

  for (const [label, swatchLabel, color, projector] of [
    [
      "Fill color",
      "light cornflower blue 3",
      "#c9daf8",
      (snapshot: WorkbookSnapshotLike) =>
        getSingleCellStyleRecordOrNull(
          snapshot,
          TOOLBAR_ACTION_CELL.sheetName,
          TOOLBAR_ACTION_CELL.address,
        )?.fill?.backgroundColor ?? null,
    ],
    [
      "Text color",
      "dark blue 1",
      "#3d85c6",
      (snapshot: WorkbookSnapshotLike) =>
        getSingleCellStyleRecordOrNull(
          snapshot,
          TOOLBAR_ACTION_CELL.sheetName,
          TOOLBAR_ACTION_CELL.address,
        )?.font?.color ?? null,
    ],
  ] as const) {
    test(`${label} preset swatch persists the selected color`, async ({ page }) => {
      const documentId = createToolbarActionDocumentId(
        `${label.toLowerCase().replace(/\s+/g, "-")}-preset`,
      );
      await prepareToolbarActionDocument(page, localServer.localServerUrl, documentId, {
        value: "toolbar",
      });

      await pickToolbarPresetColor(page, label, swatchLabel);
      await expectToolbarColor(page.getByLabel(label), color);

      await expectToolbarSnapshotProjection(
        localServer.localServerUrl,
        documentId,
        projector,
        color,
      );
    });
  }

  for (const [label, expectedHorizontal] of [
    ["Align left", "left"],
    ["Align center", "center"],
    ["Align right", "right"],
  ] as const) {
    test(`${label} persists horizontal alignment`, async ({ page }) => {
      const documentId = createToolbarActionDocumentId(label.toLowerCase().replace(/\s+/g, "-"));
      await prepareToolbarActionDocument(page, localServer.localServerUrl, documentId, {
        value: 1234.5,
      });

      await page.getByLabel(label).click();
      await expect(page.getByLabel(label)).toHaveClass(/bg-\[#e6f4ea\]/);

      await expectToolbarSnapshotProjection(
        localServer.localServerUrl,
        documentId,
        (snapshot) =>
          getSingleCellStyleRecordOrNull(
            snapshot,
            TOOLBAR_ACTION_CELL.sheetName,
            TOOLBAR_ACTION_CELL.address,
          )?.alignment?.horizontal ?? null,
        expectedHorizontal,
      );
    });
  }

  test("bottom border persists the expected border side", async ({ page }) => {
    const documentId = createToolbarActionDocumentId("border-bottom");
    await prepareToolbarActionDocument(page, localServer.localServerUrl, documentId, {
      value: "toolbar",
    });

    await page.getByLabel("Bottom border").click();

    await expectToolbarSnapshotProjection(
      localServer.localServerUrl,
      documentId,
      (snapshot) =>
        getSingleCellStyleRecordOrNull(
          snapshot,
          TOOLBAR_ACTION_CELL.sheetName,
          TOOLBAR_ACTION_CELL.address,
        )?.borders?.bottom ?? null,
      { style: "double", weight: "medium", color: "#111827" },
    );
  });

  test("all borders persists all four border sides", async ({ page }) => {
    const documentId = createToolbarActionDocumentId("border-all");
    await prepareToolbarActionDocument(page, localServer.localServerUrl, documentId, {
      value: "toolbar",
    });

    await page.getByLabel("All borders").click();

    await expectToolbarSnapshotProjection(
      localServer.localServerUrl,
      documentId,
      (snapshot) => {
        const borders =
          getSingleCellStyleRecordOrNull(
            snapshot,
            TOOLBAR_ACTION_CELL.sheetName,
            TOOLBAR_ACTION_CELL.address,
          )?.borders ?? null;
        if (!borders) {
          return null;
        }
        return BORDER_SIDES.map((side) => borders[side] ?? null);
      },
      [
        { style: "solid", weight: "thin", color: "#111827" },
        { style: "solid", weight: "thin", color: "#111827" },
        { style: "solid", weight: "thin", color: "#111827" },
        { style: "solid", weight: "thin", color: "#111827" },
      ],
    );
  });

  test("no borders removes the saved border sides", async ({ page }) => {
    const documentId = createToolbarActionDocumentId("border-none");
    await prepareToolbarActionDocument(page, localServer.localServerUrl, documentId, {
      value: "toolbar",
      setup: async (sessionId) => {
        await sendAgentRequest(localServer.localServerUrl, {
          kind: "setRangeStyle",
          id: `style:${Date.now()}`,
          sessionId,
          range: {
            sheetName: TOOLBAR_ACTION_CELL.sheetName,
            startAddress: TOOLBAR_ACTION_CELL.startAddress,
            endAddress: TOOLBAR_ACTION_CELL.endAddress,
          },
          patch: {
            borders: {
              top: { style: "solid", weight: "thin", color: "#111827" },
              right: { style: "solid", weight: "thin", color: "#111827" },
              bottom: { style: "solid", weight: "thin", color: "#111827" },
              left: { style: "solid", weight: "thin", color: "#111827" },
            },
          },
        });
      },
    });

    await expect
      .poll(async () => {
        const snapshot = await exportWorkbookSnapshot(localServer.localServerUrl, documentId);
        return Boolean(
          getSingleCellStyleRecordOrNull(
            snapshot,
            TOOLBAR_ACTION_CELL.sheetName,
            TOOLBAR_ACTION_CELL.address,
          )?.borders,
        );
      })
      .toBe(true);
    await page.getByLabel("No borders").click();

    await expectToolbarSnapshotProjection(
      localServer.localServerUrl,
      documentId,
      (snapshot) =>
        findSingleCellStyleRange(
          snapshot,
          TOOLBAR_ACTION_CELL.sheetName,
          TOOLBAR_ACTION_CELL.address,
        ) ?? null,
      null,
    );
  });

  test("clear style removes the saved style range", async ({ page }) => {
    const documentId = createToolbarActionDocumentId("clear-style");
    await prepareToolbarActionDocument(page, localServer.localServerUrl, documentId, {
      value: "toolbar",
      setup: async (sessionId) => {
        await sendAgentRequest(localServer.localServerUrl, {
          kind: "setRangeStyle",
          id: `style:${Date.now()}`,
          sessionId,
          range: {
            sheetName: TOOLBAR_ACTION_CELL.sheetName,
            startAddress: TOOLBAR_ACTION_CELL.startAddress,
            endAddress: TOOLBAR_ACTION_CELL.endAddress,
          },
          patch: {
            fill: { backgroundColor: "#fef3c7" },
            font: { bold: true, family: "Georgia", size: 14, color: "#7c2d12" },
            alignment: { horizontal: "right", wrap: true },
          },
        });
      },
    });

    await expect(page.getByLabel("Font family")).toHaveValue("Georgia");
    await expect(page.getByLabel("Font size")).toHaveValue("14");
    await expect(page.getByLabel("Bold")).toHaveClass(/bg-\[#e6f4ea\]/);
    await expect(page.getByLabel("Wrap")).toHaveAttribute("aria-pressed", "true");
    await page.getByLabel("Clear style").click();
    await expect(page.getByLabel("Bold")).not.toHaveClass(/bg-\[#e6f4ea\]/);
    await expect(page.getByLabel("Wrap")).toHaveAttribute("aria-pressed", "false");

    await expectToolbarSnapshotProjection(
      localServer.localServerUrl,
      documentId,
      (snapshot) =>
        findSingleCellStyleRange(
          snapshot,
          TOOLBAR_ACTION_CELL.sheetName,
          TOOLBAR_ACTION_CELL.address,
        ) ?? null,
      null,
    );
  });

  test("styled cell remains persisted and hydrates again after selecting another cell", async ({
    page,
  }) => {
    const documentId = createToolbarActionDocumentId("click-away-persistence");
    await prepareToolbarActionDocument(page, localServer.localServerUrl, documentId, {
      value: "ledger",
    });

    await page.getByLabel("Font family").selectOption("Georgia");
    await page.getByLabel("Bold").click();
    await setToolbarCustomColor(page, "Fill color", "#dbeafe");
    await page.getByLabel("All borders").click();

    await expectToolbarSnapshotProjection(
      localServer.localServerUrl,
      documentId,
      (snapshot) => {
        const style = getSingleCellStyleRecordOrNull(
          snapshot,
          TOOLBAR_ACTION_CELL.sheetName,
          TOOLBAR_ACTION_CELL.address,
        );
        return {
          fontFamily: style?.font?.family ?? null,
          bold: style?.font?.bold ?? false,
          fill: style?.fill?.backgroundColor ?? null,
          borders: Boolean(
            style?.borders?.top &&
            style?.borders?.right &&
            style?.borders?.bottom &&
            style?.borders?.left,
          ),
        };
      },
      {
        fontFamily: "Georgia",
        bold: true,
        fill: "#dbeafe",
        borders: true,
      },
    );

    await clickProductCell(page, 2, 2);
    await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!C3");
    await expect(page.getByLabel("Bold")).not.toHaveClass(/bg-\[#e6f4ea\]/);

    await expectToolbarSnapshotProjection(
      localServer.localServerUrl,
      documentId,
      (snapshot) => {
        const style = getSingleCellStyleRecordOrNull(
          snapshot,
          TOOLBAR_ACTION_CELL.sheetName,
          TOOLBAR_ACTION_CELL.address,
        );
        return {
          fontFamily: style?.font?.family ?? null,
          bold: style?.font?.bold ?? false,
          fill: style?.fill?.backgroundColor ?? null,
          borders: Boolean(
            style?.borders?.top &&
            style?.borders?.right &&
            style?.borders?.bottom &&
            style?.borders?.left,
          ),
        };
      },
      {
        fontFamily: "Georgia",
        bold: true,
        fill: "#dbeafe",
        borders: true,
      },
    );

    await clickProductCell(page, TOOLBAR_ACTION_CELL.columnIndex, TOOLBAR_ACTION_CELL.rowIndex);
    await expect(page.getByLabel("Font family")).toHaveValue("Georgia");
    await expect(page.getByLabel("Bold")).toHaveClass(/bg-\[#e6f4ea\]/);
    await expectToolbarColor(page.getByLabel("Fill color"), "#dbeafe");
  });

  test("range styling persists after selecting another cell", async ({ page }) => {
    const documentId = createToolbarActionDocumentId("range-click-away-persistence");
    await prepareToolbarActionDocument(page, localServer.localServerUrl, documentId, {
      value: "seed",
      setup: async (sessionId) => {
        await sendAgentRequest(localServer.localServerUrl, {
          kind: "writeRange",
          id: `write:${Date.now()}`,
          sessionId,
          range: { sheetName: "Sheet1", startAddress: "B2", endAddress: "C3" },
          values: [
            ["North", "South"],
            ["10", "20"],
          ],
        });
      },
    });

    await dragProductBodySelection(page, 1, 1, 2, 2);
    await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!B2:C3");

    await page.getByLabel("Bold").click();
    await setToolbarCustomColor(page, "Fill color", "#dbeafe");
    await page.getByLabel("All borders").click();

    await expectToolbarSnapshotProjection(
      localServer.localServerUrl,
      documentId,
      (snapshot) => {
        const sheet = getSheetSnapshot(snapshot, "Sheet1");
        const styledCells = ["B2", "C2", "B3", "C3"].map((address) => {
          const style = getStyleRecordAtCell(snapshot, "Sheet1", address);
          return {
            address,
            bold: style?.font?.bold ?? false,
            fill: style?.fill?.backgroundColor ?? null,
            borders: Boolean(
              style?.borders?.top &&
              style?.borders?.right &&
              style?.borders?.bottom &&
              style?.borders?.left,
            ),
          };
        });
        return {
          styleRangeCount: sheet.metadata?.styleRanges?.length ?? 0,
          styledCells,
        };
      },
      {
        styleRangeCount: 1,
        styledCells: [
          { address: "B2", bold: true, fill: "#dbeafe", borders: true },
          { address: "C2", bold: true, fill: "#dbeafe", borders: true },
          { address: "B3", bold: true, fill: "#dbeafe", borders: true },
          { address: "C3", bold: true, fill: "#dbeafe", borders: true },
        ],
      },
    );

    await clickProductCell(page, 4, 4);
    await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!E5");
    await expect(page.getByLabel("Bold")).not.toHaveClass(/bg-\[#e6f4ea\]/);

    await expectToolbarSnapshotProjection(
      localServer.localServerUrl,
      documentId,
      (snapshot) => {
        const style = getStyleRecordAtCell(snapshot, "Sheet1", "B2");
        return {
          bold: style?.font?.bold ?? false,
          fill: style?.fill?.backgroundColor ?? null,
          borders: Boolean(
            style?.borders?.top &&
            style?.borders?.right &&
            style?.borders?.bottom &&
            style?.borders?.left,
          ),
        };
      },
      { bold: true, fill: "#dbeafe", borders: true },
    );

    await clickProductCell(page, 1, 1);
    await expect(page.getByLabel("Bold")).toHaveClass(/bg-\[#e6f4ea\]/);
    await expectToolbarColor(page.getByLabel("Fill color"), "#dbeafe");
  });

  test("fill color visibly repaints the grid cell after clicking away", async ({ page }) => {
    const documentId = createToolbarActionDocumentId("fill-visual");
    await prepareToolbarActionDocument(page, localServer.localServerUrl, documentId);

    const grid = await getGridBox(page);
    const center = getProductCellClientPoint(
      grid,
      TOOLBAR_ACTION_CELL.columnIndex,
      TOOLBAR_ACTION_CELL.rowIndex,
      0.55,
      0.55,
    );

    await setToolbarCustomColor(page, "Fill color", "#dbeafe");
    await clickProductCell(page, 4, 4);

    const after = await sampleGridPixel(page, center.x, center.y);
    expect(colorDistance(after, { r: 219, g: 234, b: 254 })).toBeLessThan(30);
  });

  test("all borders visibly remain on the grid after clicking away", async ({ page }) => {
    const documentId = createToolbarActionDocumentId("border-visual");
    await prepareToolbarActionDocument(page, localServer.localServerUrl, documentId, {
      value: "paint",
      setup: async (sessionId) => {
        await sendAgentRequest(localServer.localServerUrl, {
          kind: "writeRange",
          id: `write:${Date.now()}`,
          sessionId,
          range: { sheetName: "Sheet1", startAddress: "B2", endAddress: "C3" },
          values: [
            ["North", "South"],
            ["10", "20"],
          ],
        });
      },
    });

    await dragProductBodySelection(page, 1, 1, 2, 2);
    await page.getByLabel("All borders").click();
    await clickProductCell(page, 4, 4);
    await expectAllBordersVisibleForRange(page, 1, 1, 2, 2);
  });

  test("all borders renders unique edge segments instead of doubled shared lines", async ({
    page,
  }) => {
    const documentId = createToolbarActionDocumentId("all-borders-no-double-stroke");
    await prepareToolbarActionDocument(page, localServer.localServerUrl, documentId);

    await dragProductBodySelection(page, 1, 1, 2, 2);
    await page.getByLabel("All borders").click();
    await clickProductCell(page, 5, 5);

    await waitForRenderedBorders(page, 12);
    await expect(page.getByTestId("grid-border-overlay-segment")).toHaveCount(12);
  });

  test("outer borders keep only the perimeter after clicking away", async ({ page }) => {
    const documentId = createToolbarActionDocumentId("outer-border-visual");
    await prepareToolbarActionDocument(page, localServer.localServerUrl, documentId);

    await dragProductBodySelection(page, 1, 1, 5, 15);
    await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!B2:F16");

    await pickToolbarPresetColor(page, "Fill color", "light red 3");
    await page.getByLabel("Outer borders").click();
    await clickProductCell(page, 8, 20);

    await waitForRenderedBorders(page);

    const grid = await getGridBox(page);
    const leftEdge = getProductCellClientPoint(grid, 1, 1, 0.03, 0.5);
    const topEdge = getProductCellClientPoint(grid, 1, 1, 0.5, 0.08);
    const rightEdge = getProductCellClientPoint(grid, 5, 1, 0.97, 0.5);
    const bottomEdge = getProductCellClientPoint(grid, 1, 15, 0.5, 0.92);
    const innerVertical = getProductCellClientPoint(grid, 1, 8, 0.97, 0.5);
    const innerHorizontal = getProductCellClientPoint(grid, 3, 1, 0.5, 0.92);

    await expect
      .poll(
        async () =>
          await borderVisibilityMatches(page, [
            { x: leftEdge.x, y: leftEdge.y, mode: "present", tolerance: 12 },
            { x: topEdge.x, y: topEdge.y, mode: "present", tolerance: 12 },
            { x: rightEdge.x, y: rightEdge.y, mode: "present", tolerance: 12 },
            { x: bottomEdge.x, y: bottomEdge.y, mode: "present", tolerance: 12 },
            { x: innerVertical.x, y: innerVertical.y, mode: "absent" },
            { x: innerHorizontal.x, y: innerHorizontal.y, mode: "absent" },
          ]),
        {
          message: "outer-border overlay should settle on the perimeter only",
        },
      )
      .toBe(true);
  });

  test("large blank range keeps all borders after clicking a far away cell", async ({ page }) => {
    const documentId = createToolbarActionDocumentId("large-range-border-click-away");
    await prepareToolbarActionDocument(page, localServer.localServerUrl, documentId);

    await dragProductBodySelection(page, 1, 1, 5, 15);
    await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!B2:F16");

    await page.getByLabel("All borders").click();
    await clickProductCell(page, 7, 2);
    await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!H3");

    await expectToolbarSnapshotProjection(
      localServer.localServerUrl,
      documentId,
      (snapshot) => {
        const sheet = getSheetSnapshot(snapshot, "Sheet1");
        return {
          styleRangeCount: sheet.metadata?.styleRanges?.length ?? 0,
          topLeftBorders: Boolean(
            getStyleRecordAtCell(snapshot, "Sheet1", "B2")?.borders?.top &&
            getStyleRecordAtCell(snapshot, "Sheet1", "B2")?.borders?.right &&
            getStyleRecordAtCell(snapshot, "Sheet1", "B2")?.borders?.bottom &&
            getStyleRecordAtCell(snapshot, "Sheet1", "B2")?.borders?.left,
          ),
          bottomRightBorders: Boolean(
            getStyleRecordAtCell(snapshot, "Sheet1", "F16")?.borders?.top &&
            getStyleRecordAtCell(snapshot, "Sheet1", "F16")?.borders?.right &&
            getStyleRecordAtCell(snapshot, "Sheet1", "F16")?.borders?.bottom &&
            getStyleRecordAtCell(snapshot, "Sheet1", "F16")?.borders?.left,
          ),
        };
      },
      {
        styleRangeCount: 1,
        topLeftBorders: true,
        bottomRightBorders: true,
      },
    );

    await expectAllBordersVisibleForRange(page, 1, 1, 5, 15);
  });
});

test("web app reflects a local-server agent write in the rendered spreadsheet", async ({
  page,
}) => {
  test.slow();
  const port = await reserveLocalPort();
  const documentId = `playwright-${Date.now()}`;
  const localServer = await startLocalServer(port);

  try {
    await page.goto(
      `/?document=${encodeURIComponent(documentId)}&server=${encodeURIComponent(localServer.localServerUrl)}`,
    );

    const nameBox = page.getByTestId("name-box");
    const formulaInput = page.getByTestId("formula-input");

    await expect(nameBox).toHaveValue("A1");
    await expect(formulaInput).toHaveValue("");

    await expect
      .poll(
        async () => {
          const documentState = await fetchDocumentState(localServer.localServerUrl, documentId);
          return documentState.browserSessions.length > 0;
        },
        {
          message: "browser should attach to the local-server document session",
        },
      )
      .toBe(true);

    await withAgentSession(
      localServer.localServerUrl,
      documentId,
      `playwright-agent:${Date.now()}`,
      async (sessionId) => {
        await sendAgentRequest(localServer.localServerUrl, {
          kind: "writeRange",
          id: `write:${Date.now()}`,
          sessionId,
          range: {
            sheetName: "Sheet1",
            startAddress: "A1",
            endAddress: "A1",
          },
          values: [[42]],
        });
      },
    );

    await expect(formulaInput).toHaveValue("42");
  } catch (error) {
    throw new Error(
      `${error instanceof Error ? error.message : String(error)}\nLocal-server logs:\n${localServer.getLogs()}`,
      { cause: error },
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

  await grid.click({
    position: {
      x: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH + Math.floor(PRODUCT_COLUMN_WIDTH / 2),
      y: Math.floor(PRODUCT_HEADER_HEIGHT / 2),
    },
  });
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!B:B");

  await grid.click({
    position: {
      x: Math.floor(PRODUCT_ROW_MARKER_WIDTH / 2),
      y: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2),
    },
  });
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

test("web app keeps the active focus inside the Glide grid when clicking a cell", async ({
  page,
}) => {
  await page.goto("/");

  await clickProductCell(page, 2, 2);
  await expect(page.getByTestId("name-box")).toHaveValue("C3");

  const activeElementState = await page.evaluate(() => {
    const active = document.activeElement;
    return {
      testId: active?.getAttribute("data-testid") ?? null,
      insideSheetGrid: Boolean(active?.closest('[data-testid="sheet-grid"]')),
    };
  });

  expect(activeElementState.insideSheetGrid).toBe(true);
  expect(activeElementState.testId).not.toBe("sheet-grid");
});

test("web app maps clicks in the upper half of a cell to that same visible cell", async ({
  page,
}) => {
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
  await formulaInput.fill('=A1="HELLO"');
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
  const targetY =
    grid.y + PRODUCT_HEADER_HEIGHT + 4 * PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2);
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

test("web app supports rectangular clipboard copy and external paste", async ({
  page,
  context,
}) => {
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

  await expect
    .poll(() => page.evaluate(() => navigator.clipboard.readText()))
    .toBe("11\t12\n13\t14");

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

test("web app relocates formulas when using rectangular clipboard paste", async ({
  page,
  context,
}) => {
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
