import { createHash } from "node:crypto";
import { writeFile } from "node:fs/promises";
import { expect, test, type Locator, type Page, type TestInfo } from "@playwright/test";
import fc from "fast-check";
import { runProperty, shouldRunFuzzSuite } from "../../packages/test-fuzz/src/index.ts";

const PRODUCT_ROW_MARKER_WIDTH = 46;
const PRODUCT_COLUMN_WIDTH = 104;
const PRODUCT_HEADER_HEIGHT = 24;
const PRODUCT_ROW_HEIGHT = 22;
const PRIMARY_MODIFIER = process.platform === "darwin" ? "Meta" : "Control";
const fuzzBrowserEnabled = process.env["BILIG_FUZZ_BROWSER"] === "1";

type BrowserSelectionAction =
  | { kind: "click"; row: number; col: number }
  | { kind: "shiftClick"; row: number; col: number }
  | { kind: "key"; key: "ArrowLeft" | "ArrowRight" | "ArrowUp" | "ArrowDown"; shift: boolean };

interface ToolbarSyncAction {
  readonly label: string;
  readonly apply: (page: Page) => Promise<void>;
}

function delay(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
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
  const { sourceX, sourceY, targetX, targetY } = await getProductFillHandleDragPoints(
    page,
    sourceCol,
    sourceRow,
    targetCol,
    targetRow,
  );

  await page.mouse.move(sourceX, sourceY);
  await page.mouse.down();
  await page.mouse.move(targetX, targetY, {
    steps: 10,
  });
  await page.mouse.up();
}

async function getProductFillHandleDragPoints(
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

  return {
    sourceX: sourceLeft + sourceWidth - 3,
    sourceY: sourceTop + PRODUCT_ROW_HEIGHT - 3,
    targetX: targetLeft + targetWidth - 3,
    targetY: targetTop + PRODUCT_ROW_HEIGHT - 3,
  };
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
  await expect(locator).toBeVisible();
  const box = await locator.boundingBox();
  if (!box) {
    throw new Error("locator is not visible");
  }
  return box;
}

async function clickProductCell(
  page: Parameters<typeof test>[0]["page"],
  columnIndex: number,
  rowIndex: number,
  options?: {
    shift?: boolean;
  },
) {
  const gridLocator = page.getByTestId("sheet-grid");
  await expect(gridLocator).toBeVisible();
  const grid = await gridLocator.boundingBox();
  if (!grid) {
    throw new Error("sheet grid is not visible");
  }

  const columnLeft = await getProductColumnLeft(page, columnIndex);
  const columnWidth = await getProductColumnWidth(page, columnIndex);
  if (options?.shift) {
    await page.keyboard.down("Shift");
  }
  try {
    await page.mouse.click(
      grid.x + columnLeft + Math.floor(columnWidth / 2),
      grid.y +
        PRODUCT_HEADER_HEIGHT +
        rowIndex * PRODUCT_ROW_HEIGHT +
        Math.floor(PRODUCT_ROW_HEIGHT / 2),
    );
  } finally {
    if (options?.shift) {
      await page.keyboard.up("Shift");
    }
  }
}

async function clickSelectionFuzzCell(
  page: Parameters<typeof test>[0]["page"],
  columnIndex: number,
  rowIndex: number,
  shift = false,
) {
  const gridLocator = page.getByTestId("sheet-grid");
  await expect(gridLocator).toBeVisible();
  const grid = await gridLocator.boundingBox();
  if (!grid) {
    throw new Error("sheet grid is not visible");
  }
  const columnLeft = await getProductColumnLeft(page, columnIndex);
  const columnWidth = await getProductColumnWidth(page, columnIndex);
  const scrollLeft = await gridLocator.evaluate(
    (node, target) => {
      const scrollViewport = node.querySelector('[aria-hidden="true"]');
      if (!(scrollViewport instanceof HTMLElement)) {
        return 0;
      }
      const targetCenter = target.columnLeft + target.columnWidth / 2;
      const visibleStart = scrollViewport.scrollLeft;
      const visibleEnd = visibleStart + scrollViewport.clientWidth;
      if (targetCenter < visibleStart || targetCenter > visibleEnd) {
        scrollViewport.scrollLeft = Math.max(0, targetCenter - scrollViewport.clientWidth / 2);
      }
      return scrollViewport.scrollLeft;
    },
    {
      columnLeft,
      columnWidth,
    },
  );
  const point = {
    x: grid.x + columnLeft - scrollLeft + columnWidth / 2,
    y: grid.y + PRODUCT_HEADER_HEIGHT + rowIndex * PRODUCT_ROW_HEIGHT + PRODUCT_ROW_HEIGHT / 2,
  };
  if (shift) {
    await page.keyboard.down("Shift");
  }
  try {
    await page.mouse.click(point.x, point.y);
  } finally {
    if (shift) {
      await page.keyboard.up("Shift");
    }
  }
}

async function runSelectionFuzzActions(
  page: Parameters<typeof test>[0]["page"],
  grid: Locator,
  actions: readonly BrowserSelectionAction[],
  index = 0,
): Promise<void> {
  const action = actions[index];
  if (!action) {
    return;
  }

  if (action.kind === "click") {
    await clickSelectionFuzzCell(page, action.col, action.row);
  } else if (action.kind === "shiftClick") {
    await clickSelectionFuzzCell(page, action.col, action.row, true);
  } else {
    await grid.press(action.shift ? `Shift+${action.key}` : action.key);
  }

  const selection = await page.getByTestId("status-selection").textContent();
  expect(selection).toMatch(
    /^Sheet1!(?:[A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?|[A-Z]+:[A-Z]+|[0-9]+:[0-9]+|All)$/,
  );

  const focusInsideShell = await page.evaluate(() => {
    const active = document.activeElement;
    return Boolean(
      active?.closest('[data-testid="sheet-grid"]') ||
      active?.closest('[data-testid="formula-bar"]') ||
      active?.closest('[role="toolbar"]'),
    );
  });
  expect(focusInsideShell).toBe(true);

  await runSelectionFuzzActions(page, grid, actions, index + 1);
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

async function dragProductSelectionBorder(
  page: Parameters<typeof test>[0]["page"],
  startColumn: number,
  startRow: number,
  endColumn: number,
  endRow: number,
  targetColumn: number,
  targetRow: number,
) {
  const gridLocator = page.getByTestId("sheet-grid");
  await expect(gridLocator).toBeVisible();
  const grid = await gridLocator.boundingBox();
  if (!grid) {
    throw new Error("sheet grid is not visible");
  }

  const startLeft = await getProductColumnLeft(page, startColumn);
  const rangeTop = grid.y + PRODUCT_HEADER_HEIGHT + startRow * PRODUCT_ROW_HEIGHT;
  const sourceX = grid.x + startLeft + 3;
  const sourceY = rangeTop + 2;
  const targetLeft = await getProductColumnLeft(page, targetColumn);
  const targetWidth = await getProductColumnWidth(page, targetColumn);
  const targetX = grid.x + targetLeft + Math.floor(targetWidth / 2);
  const targetY =
    grid.y +
    PRODUCT_HEADER_HEIGHT +
    targetRow * PRODUCT_ROW_HEIGHT +
    Math.floor(PRODUCT_ROW_HEIGHT / 2);

  await page.mouse.move(sourceX, sourceY);
  await page.mouse.down();
  await page.mouse.move(targetX, targetY, { steps: 12 });
  await page.mouse.up();
}

type ToolbarPage = Parameters<typeof test>[0]["page"];

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

function getToolbarButton(page: ToolbarPage, label: string): Locator {
  return page.getByRole("button", { name: label, exact: true });
}

function getToolbarSelect(page: ToolbarPage, label: string): Locator {
  return page.getByRole("combobox", { name: label, exact: true });
}

async function expectToolbarColor(locator: Locator, value: string) {
  await expect(locator).toHaveAttribute("data-current-color", value.toLowerCase());
}

async function expectToolbarSelectValue(page: ToolbarPage, label: string, value: string) {
  await expect(getToolbarSelect(page, label)).toHaveAttribute("data-current-value", value);
}

async function selectToolbarOption(
  page: ToolbarPage,
  label: string,
  optionLabel: string,
  expectedValue = optionLabel,
) {
  await getToolbarSelect(page, label).click();
  await page.getByRole("option", { name: optionLabel, exact: true }).click();
  await expectToolbarSelectValue(page, label, expectedValue);
}

async function setToolbarCustomColor(
  page: ToolbarPage,
  controlLabel: "Fill color" | "Text color",
  value: string,
) {
  await getToolbarButton(page, controlLabel).click();
  await page.getByLabel(`Open custom ${controlLabel.toLowerCase()} picker`).click();
  await setColorInput(
    page.getByLabel(controlLabel === "Fill color" ? "Custom fill color" : "Custom text color", {
      exact: true,
    }),
    value,
  );
}

async function pickToolbarPresetColor(
  page: ToolbarPage,
  controlLabel: "Fill color" | "Text color",
  swatchLabel: string,
) {
  await getToolbarButton(page, controlLabel).click();
  await page.getByLabel(`${controlLabel} ${swatchLabel}`).click();
}

async function pickToolbarBorderPreset(
  page: ToolbarPage,
  presetLabel:
    | "All borders"
    | "Outer borders"
    | "Left border"
    | "Top border"
    | "Right border"
    | "Bottom border"
    | "Clear borders",
) {
  await getToolbarButton(page, "Borders").click();
  await page.getByRole("button", { name: presetLabel }).click();
}

const TOOLBAR_SYNC_ACTIONS: readonly ToolbarSyncAction[] = [
  {
    label: "number-format-accounting",
    apply: async (activePage) =>
      await selectToolbarOption(activePage, "Number format", "Accounting", "accounting"),
  },
  {
    label: "font-size-14",
    apply: async (activePage) => await selectToolbarOption(activePage, "Font size", "14"),
  },
  { label: "bold", apply: async (activePage) => await activePage.getByLabel("Bold").click() },
  {
    label: "italic",
    apply: async (activePage) => await activePage.getByLabel("Italic").click(),
  },
  {
    label: "underline",
    apply: async (activePage) => await activePage.getByLabel("Underline").click(),
  },
  {
    label: "fill-color",
    apply: async (activePage) => await setToolbarCustomColor(activePage, "Fill color", "#dbeafe"),
  },
  {
    label: "text-color",
    apply: async (activePage) => await setToolbarCustomColor(activePage, "Text color", "#7c2d12"),
  },
  {
    label: "align-left",
    apply: async (activePage) => await activePage.getByLabel("Align left").click(),
  },
  {
    label: "align-center",
    apply: async (activePage) => await activePage.getByLabel("Align center").click(),
  },
  {
    label: "align-right",
    apply: async (activePage) => await activePage.getByLabel("Align right").click(),
  },
  {
    label: "border-all",
    apply: async (activePage) => await pickToolbarBorderPreset(activePage, "All borders"),
  },
  {
    label: "border-outer",
    apply: async (activePage) => await pickToolbarBorderPreset(activePage, "Outer borders"),
  },
  {
    label: "border-left",
    apply: async (activePage) => await pickToolbarBorderPreset(activePage, "Left border"),
  },
  {
    label: "border-top",
    apply: async (activePage) => await pickToolbarBorderPreset(activePage, "Top border"),
  },
  {
    label: "border-right",
    apply: async (activePage) => await pickToolbarBorderPreset(activePage, "Right border"),
  },
  {
    label: "border-bottom",
    apply: async (activePage) => await pickToolbarBorderPreset(activePage, "Bottom border"),
  },
  {
    label: "border-clear",
    apply: async (activePage) => await pickToolbarBorderPreset(activePage, "Clear borders"),
  },
  { label: "wrap", apply: async (activePage) => await activePage.getByLabel("Wrap").click() },
  {
    label: "clear-style",
    apply: async (activePage) => await activePage.getByLabel("Clear style").click(),
  },
  {
    label: "number-format-general",
    apply: async (activePage) =>
      await selectToolbarOption(activePage, "Number format", "General", "general"),
  },
];

async function selectToolbarActionRange(page: Page) {
  await clickProductCell(page, 1, 1);
  await clickProductCell(page, 2, 2, { shift: true });
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!B2:C3");
}

async function seedToolbarActionRange(page: Page) {
  const nameBox = page.getByTestId("name-box");
  const formulaInput = page.getByTestId("formula-input");
  await nameBox.fill("B2");
  await nameBox.press("Enter");
  await formulaInput.fill("1234.5");
  await formulaInput.press("Enter");

  await nameBox.fill("C2");
  await nameBox.press("Enter");
  await formulaInput.fill("6789.125");
  await formulaInput.press("Enter");

  await nameBox.fill("B3");
  await nameBox.press("Enter");
  await formulaInput.fill("42.25");
  await formulaInput.press("Enter");

  await nameBox.fill("C3");
  await nameBox.press("Enter");
  await formulaInput.fill("-7.5");
  await formulaInput.press("Enter");
}

async function captureGridRangeScreenshot(
  page: Page,
  startColumn: number,
  startRow: number,
  endColumn: number,
  endRow: number,
) {
  await page.bringToFront();
  const gridLocator = page.getByTestId("sheet-grid");
  await expect(gridLocator).toBeVisible();
  const grid = await gridLocator.boundingBox();
  if (!grid) {
    throw new Error("sheet grid is not visible");
  }

  const minColumn = Math.min(startColumn, endColumn);
  const maxColumn = Math.max(startColumn, endColumn);
  const minRow = Math.min(startRow, endRow);
  const maxRow = Math.max(startRow, endRow);
  const startLeft = await getProductColumnLeft(page, minColumn);
  const endLeft = await getProductColumnLeft(page, maxColumn);
  const endWidth = await getProductColumnWidth(page, maxColumn);

  return await page.screenshot({
    animations: "disabled",
    caret: "hide",
    clip: {
      x: Math.round(grid.x + startLeft),
      y: Math.round(grid.y + PRODUCT_HEADER_HEIGHT + minRow * PRODUCT_ROW_HEIGHT),
      width: Math.round(endLeft + endWidth - startLeft),
      height: Math.round((maxRow - minRow + 1) * PRODUCT_ROW_HEIGHT),
    },
  });
}

async function compareScreenshotPixels(page: Page, left: Buffer, right: Buffer) {
  return await page.evaluate(
    async ({ leftDataUrl, rightDataUrl, channelTolerance }) => {
      const [leftImage, rightImage] = await Promise.all([
        new Promise<HTMLImageElement>((resolve, reject) => {
          const image = new Image();
          image.addEventListener("load", () => resolve(image), { once: true });
          image.addEventListener(
            "error",
            () => reject(new Error("Failed to decode left screenshot data URL")),
            { once: true },
          );
          image.src = leftDataUrl;
        }),
        new Promise<HTMLImageElement>((resolve, reject) => {
          const image = new Image();
          image.addEventListener("load", () => resolve(image), { once: true });
          image.addEventListener(
            "error",
            () => reject(new Error("Failed to decode right screenshot data URL")),
            { once: true },
          );
          image.src = rightDataUrl;
        }),
      ]);
      if (
        leftImage.naturalWidth !== rightImage.naturalWidth ||
        leftImage.naturalHeight !== rightImage.naturalHeight
      ) {
        return {
          equal: false,
          diffPixels: Number.POSITIVE_INFINITY,
          width: leftImage.naturalWidth,
          height: leftImage.naturalHeight,
        };
      }

      const width = leftImage.naturalWidth;
      const height = leftImage.naturalHeight;
      const canvas = document.createElement("canvas");
      canvas.width = width;
      canvas.height = height;
      const context = canvas.getContext("2d");
      if (!context) {
        throw new Error("Missing 2d context for screenshot comparison");
      }

      context.clearRect(0, 0, width, height);
      context.drawImage(leftImage, 0, 0);
      const leftPixels = context.getImageData(0, 0, width, height).data;
      context.clearRect(0, 0, width, height);
      context.drawImage(rightImage, 0, 0);
      const rightPixels = context.getImageData(0, 0, width, height).data;

      let diffPixels = 0;
      for (let index = 0; index < leftPixels.length; index += 4) {
        if (
          Math.abs(leftPixels[index] - rightPixels[index]) > channelTolerance ||
          Math.abs(leftPixels[index + 1] - rightPixels[index + 1]) > channelTolerance ||
          Math.abs(leftPixels[index + 2] - rightPixels[index + 2]) > channelTolerance ||
          Math.abs(leftPixels[index + 3] - rightPixels[index + 3]) > channelTolerance
        ) {
          diffPixels += 1;
        }
      }

      return { equal: diffPixels === 0, diffPixels, width, height };
    },
    {
      leftDataUrl: `data:image/png;base64,${left.toString("base64")}`,
      rightDataUrl: `data:image/png;base64,${right.toString("base64")}`,
      channelTolerance: 2,
    },
  );
}

async function pollMatchingGridRangeScreenshots(
  primaryPage: Page,
  mirrorPage: Page,
  startedAt: number,
  timeoutMs: number,
  maxDiffPixels: number,
  startColumn: number,
  startRow: number,
  endColumn: number,
  endRow: number,
): Promise<{
  primaryBuffer: Buffer;
  mirrorBuffer: Buffer;
  diffPixels: number;
  matched: boolean;
}> {
  const [primaryBuffer, mirrorBuffer] = await Promise.all([
    captureGridRangeScreenshot(primaryPage, startColumn, startRow, endColumn, endRow),
    captureGridRangeScreenshot(mirrorPage, startColumn, startRow, endColumn, endRow),
  ]);
  const comparison = await compareScreenshotPixels(primaryPage, primaryBuffer, mirrorBuffer);
  const matched = comparison.equal || comparison.diffPixels <= maxDiffPixels;
  if (matched || Date.now() - startedAt > timeoutMs) {
    return {
      primaryBuffer,
      mirrorBuffer,
      diffPixels: comparison.diffPixels,
      matched,
    };
  }

  await delay(50);
  return await pollMatchingGridRangeScreenshots(
    primaryPage,
    mirrorPage,
    startedAt,
    timeoutMs,
    maxDiffPixels,
    startColumn,
    startRow,
    endColumn,
    endRow,
  );
}

async function expectMatchingGridRangeScreenshots(
  primaryPage: Page,
  mirrorPage: Page,
  actionLabel: string,
  testInfo: TestInfo,
  startColumn: number,
  startRow: number,
  endColumn: number,
  endRow: number,
  timeoutMs = 1_500,
  maxDiffPixels = 8,
) {
  const startedAt = Date.now();
  const result = await pollMatchingGridRangeScreenshots(
    primaryPage,
    mirrorPage,
    startedAt,
    timeoutMs,
    maxDiffPixels,
    startColumn,
    startRow,
    endColumn,
    endRow,
  );
  if (result.matched) {
    return Date.now() - startedAt;
  }

  const primaryHash = createHash("sha256").update(result.primaryBuffer).digest("hex");
  const mirrorHash = createHash("sha256").update(result.mirrorBuffer).digest("hex");
  await writeFile(
    testInfo.outputPath(`multiplayer-${actionLabel}-range-primary.png`),
    result.primaryBuffer,
  );
  await writeFile(
    testInfo.outputPath(`multiplayer-${actionLabel}-range-mirror.png`),
    result.mirrorBuffer,
  );

  throw new Error(
    `multiplayer grid range screenshots diverged for ${actionLabel} after ${timeoutMs}ms (primary=${primaryHash}, mirror=${mirrorHash}, diffPixels=${result.diffPixels}, maxDiffPixels=${maxDiffPixels})`,
  );
}

async function openZeroWorkbookPage(page: Page, documentId: string) {
  await page.goto(`/?document=${encodeURIComponent(documentId)}`);
  await waitForWorkbookReady(page);
  await selectToolbarActionRange(page);
}

async function waitForWorkbookReady(page: Page) {
  await expect(page.getByTestId("formula-bar")).toBeVisible({ timeout: 15_000 });
  await expect(page.getByTestId("sheet-grid")).toBeVisible({ timeout: 15_000 });
  await expect(page.getByTestId("status-sync")).toHaveText("Ready", { timeout: 15_000 });
}

async function runToolbarSyncActions(
  page: Page,
  mirrorPage: Page,
  actions: readonly ToolbarSyncAction[],
  testInfo: TestInfo,
  index = 0,
): Promise<void> {
  const action = actions[index];
  if (!action) {
    return;
  }

  await action.apply(page);
  await selectToolbarActionRange(page);
  await selectToolbarActionRange(mirrorPage);
  const elapsed = await expectMatchingGridRangeScreenshots(
    page,
    mirrorPage,
    action.label,
    testInfo,
    1,
    1,
    2,
    2,
    1_500,
  );
  expect(elapsed).toBeLessThanOrEqual(1_500);
  await runToolbarSyncActions(page, mirrorPage, actions, testInfo, index + 1);
}

test("web app renders the minimal product shell without legacy demo chrome", async ({ page }) => {
  await page.goto("/");
  await waitForWorkbookReady(page);

  await expect(page.getByTestId("formula-bar")).toBeVisible();
  await expect(page.getByTestId("name-box")).toBeVisible();
  await expect(page.getByTestId("sheet-grid")).toBeVisible();
  await expect(page.getByRole("tab", { name: "Sheet1" })).toBeVisible();
  await expect(page.getByRole("heading", { name: "bilig-demo" })).toHaveCount(0);

  await expect(page.getByTestId("preset-strip")).toHaveCount(0);
  await expect(page.getByTestId("metrics-panel")).toHaveCount(0);
  await expect(page.getByTestId("replica-panel")).toHaveCount(0);

  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A1");
  await expect(page.getByTestId("status-sync")).toHaveText("Ready", {
    timeout: 15_000,
  });
  await expect(page.locator(".formula-result-shell")).toHaveCount(0);
});

test("web app keeps toolbar controls aligned and consistently sized", async ({ page }) => {
  await page.goto("/");
  await waitForWorkbookReady(page);

  const toolbar = page.getByRole("toolbar", { name: "Formatting toolbar" });
  await expect(toolbar).toBeVisible();

  const controls = [
    page.getByLabel("Number format"),
    page.getByLabel("Font size"),
    page.getByLabel("Bold"),
    page.getByLabel("Italic"),
    page.getByLabel("Underline"),
    page.getByLabel("Fill color"),
    page.getByLabel("Text color"),
    page.getByLabel("Align left"),
    page.getByLabel("Align center"),
    page.getByLabel("Align right"),
    page.getByLabel("Borders"),
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
  await waitForWorkbookReady(page);

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
  await waitForWorkbookReady(page);

  const nameBox = page.getByTestId("name-box");
  const formulaFrame = page.getByTestId("formula-input-frame");

  const [nameBoxBox, formulaFrameBox] = await Promise.all([getBox(nameBox), getBox(formulaFrame)]);

  expect(Math.abs(nameBoxBox.height - formulaFrameBox.height)).toBeLessThanOrEqual(1);
  expect(Math.abs(nameBoxBox.y - formulaFrameBox.y)).toBeLessThanOrEqual(1);
  expect(
    Math.abs(nameBoxBox.y + nameBoxBox.height - (formulaFrameBox.y + formulaFrameBox.height)),
  ).toBeLessThanOrEqual(1);
});

test("web app keeps shell controls on one height and radius system", async ({ page }) => {
  await page.goto("/");
  await waitForWorkbookReady(page);

  const locators = [
    page.getByLabel("Number format"),
    page.getByTestId("name-box"),
    page.getByTestId("formula-input-frame"),
    page.getByTestId("status-mode"),
    page.getByRole("tab", { name: "Sheet1" }),
  ];

  const metrics = await Promise.all(
    locators.map(async (locator) => ({
      height: Math.round((await getBox(locator)).height),
      radius: await locator.evaluate((element) => getComputedStyle(element).borderRadius),
    })),
  );

  const heights = metrics.map(({ height }) => height);
  expect(Math.max(...heights) - Math.min(...heights)).toBeLessThanOrEqual(1);
  expect(new Set(metrics.map(({ radius }) => radius)).size).toBe(1);
});

test("web app keeps the toolbar compact on narrow viewports", async ({ page }) => {
  await page.setViewportSize({ width: 620, height: 760 });
  await page.goto("/");
  await waitForWorkbookReady(page);

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
  await waitForWorkbookReady(page);

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
  await waitForWorkbookReady(page);

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
  await expect(page.getByRole("button", { name: "Show fill color swatches" })).toBeVisible();
});

test("web app applies preset swatch colors directly from the palette", async ({ page }) => {
  await page.goto("/");
  await waitForWorkbookReady(page);

  await pickToolbarPresetColor(page, "Fill color", "light cornflower blue 3");
  await expectToolbarColor(getToolbarButton(page, "Fill color"), "#c9daf8");

  await pickToolbarPresetColor(page, "Text color", "dark blue 1");
  await expectToolbarColor(getToolbarButton(page, "Text color"), "#3d85c6");
});

test("web app propagates content and styling changes across live zero tabs", async ({
  page,
}, testInfo) => {
  test.slow();
  const documentId = `playwright-zero-style-multiplayer-${Date.now()}`;
  const mirrorPage = await page.context().newPage();
  const viewport = page.viewportSize();
  if (viewport) {
    await mirrorPage.setViewportSize(viewport);
  }

  try {
    await Promise.all([
      openZeroWorkbookPage(page, documentId),
      openZeroWorkbookPage(mirrorPage, documentId),
    ]);

    await clickProductCell(page, 1, 1);
    await page.keyboard.type("relay");
    await page.keyboard.press("Enter");
    await selectToolbarActionRange(page);
    await selectToolbarActionRange(mirrorPage);

    const contentElapsed = await expectMatchingGridRangeScreenshots(
      page,
      mirrorPage,
      "zero-content-relay",
      testInfo,
      1,
      1,
      2,
      2,
      1_500,
      8,
    );
    expect(contentElapsed).toBeLessThanOrEqual(1_500);

    await page.getByLabel("Bold").click();
    await pickToolbarPresetColor(page, "Fill color", "light cornflower blue 3");
    await pickToolbarBorderPreset(page, "All borders");
    await selectToolbarActionRange(page);
    await selectToolbarActionRange(mirrorPage);

    const styleElapsed = await expectMatchingGridRangeScreenshots(
      page,
      mirrorPage,
      "zero-style-relay",
      testInfo,
      1,
      1,
      2,
      2,
      1_500,
      8,
    );
    expect(styleElapsed).toBeLessThanOrEqual(1_500);
  } finally {
    await mirrorPage.close().catch(() => undefined);
  }
});

test("web app keeps two live zero tabs visually converged across toolbar actions", async ({
  page,
}, testInfo) => {
  test.slow();
  const documentId = `playwright-zero-toolbar-multiplayer-${Date.now()}`;
  const mirrorPage = await page.context().newPage();
  const viewport = page.viewportSize();
  if (viewport) {
    await mirrorPage.setViewportSize(viewport);
  }

  try {
    await Promise.all([
      openZeroWorkbookPage(page, documentId),
      openZeroWorkbookPage(mirrorPage, documentId),
    ]);
    await seedToolbarActionRange(page);
    await selectToolbarActionRange(page);
    await selectToolbarActionRange(mirrorPage);

    const initialElapsed = await expectMatchingGridRangeScreenshots(
      page,
      mirrorPage,
      "zero-toolbar-initial",
      testInfo,
      1,
      1,
      2,
      2,
      5_000,
      8,
    );
    expect(initialElapsed).toBeLessThanOrEqual(5_000);

    await runToolbarSyncActions(page, mirrorPage, TOOLBAR_SYNC_ACTIONS, testInfo);
  } finally {
    await mirrorPage.close().catch(() => undefined);
  }
});

test("web app preserves an in-progress local draft when another tab edits the same cell", async ({
  page,
}) => {
  test.slow();
  const documentId = `playwright-zero-same-cell-draft-${Date.now()}`;
  const mirrorPage = await page.context().newPage();
  const viewport = page.viewportSize();
  if (viewport) {
    await mirrorPage.setViewportSize(viewport);
  }

  try {
    await Promise.all([
      openZeroWorkbookPage(page, documentId),
      openZeroWorkbookPage(mirrorPage, documentId),
    ]);
    await expect(page.getByTestId("worker-error")).toHaveCount(0);
    await expect(mirrorPage.getByTestId("worker-error")).toHaveCount(0);

    const formulaInput = page.getByTestId("formula-input");
    const mirrorFormulaInput = mirrorPage.getByTestId("formula-input");

    await clickProductCell(page, 0, 0);
    await clickProductCell(mirrorPage, 0, 0);
    await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A1");
    await expect(mirrorPage.getByTestId("status-selection")).toHaveText("Sheet1!A1");

    await formulaInput.focus();
    await formulaInput.selectText();
    await page.keyboard.type("local-draft");
    await expect(formulaInput).toHaveValue("local-draft");

    await mirrorFormulaInput.focus();
    await mirrorFormulaInput.selectText();
    await mirrorPage.keyboard.type("remote");
    await mirrorFormulaInput.press("Enter");
    await expect(mirrorFormulaInput).toHaveValue("remote");

    await expect(formulaInput).toHaveValue("local-draft");

    await formulaInput.press("Escape");
    await expect(formulaInput).toHaveValue("remote");
  } finally {
    await mirrorPage.close().catch(() => undefined);
  }
});

test("web app compares and applies a stale same-cell draft without losing local work", async ({
  page,
}) => {
  test.slow();
  const documentId = `playwright-zero-same-cell-conflict-${Date.now()}`;
  const mirrorPage = await page.context().newPage();
  const viewport = page.viewportSize();
  if (viewport) {
    await mirrorPage.setViewportSize(viewport);
  }

  try {
    await Promise.all([
      openZeroWorkbookPage(page, documentId),
      openZeroWorkbookPage(mirrorPage, documentId),
    ]);
    await expect(page.getByTestId("worker-error")).toHaveCount(0);
    await expect(mirrorPage.getByTestId("worker-error")).toHaveCount(0);

    const formulaInput = page.getByTestId("formula-input");
    const mirrorFormulaInput = mirrorPage.getByTestId("formula-input");

    await clickProductCell(page, 0, 0);
    await clickProductCell(mirrorPage, 0, 0);
    await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A1");
    await expect(mirrorPage.getByTestId("status-selection")).toHaveText("Sheet1!A1");

    await formulaInput.focus();
    await formulaInput.selectText();
    await page.keyboard.type("local-draft");
    await expect(formulaInput).toHaveValue("local-draft");

    await mirrorFormulaInput.focus();
    await mirrorFormulaInput.selectText();
    await mirrorPage.keyboard.type("remote");
    await mirrorFormulaInput.press("Enter");
    await expect(mirrorFormulaInput).toHaveValue("remote");

    await expect(page.getByTestId("editor-conflict-banner")).toContainText(
      "Remote update detected in Sheet1!A1 while you were editing.",
    );
    await expect(formulaInput).toHaveValue("local-draft");

    await formulaInput.press("Enter");
    await expect(page.getByTestId("editor-conflict-apply-mine")).toBeVisible();

    await page.getByTestId("editor-conflict-apply-mine").click();

    await expect(page.getByTestId("editor-conflict-banner")).toHaveCount(0);
    await expect(formulaInput).toHaveValue("local-draft");
    await expect(mirrorFormulaInput).toHaveValue("local-draft");
  } finally {
    await mirrorPage.close().catch(() => undefined);
  }
});

test("web app reverts an authoritative change from the changes pane", async ({ page }) => {
  const documentId = `playwright-zero-change-revert-${Date.now()}`;
  await openZeroWorkbookPage(page, documentId);

  const formulaInput = page.getByTestId("formula-input");
  const changesTab = page.getByTestId("workbook-side-rail-tab-changes");

  await clickProductCell(page, 0, 0);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A1");
  await formulaInput.fill("seed");
  await formulaInput.press("Enter");
  await expect(formulaInput).toHaveValue("seed");

  await expect(changesTab).toContainText("1");
  await changesTab.click();

  const changeRows = page.getByTestId("workbook-change-row");
  await expect(changeRows).toHaveCount(1);
  await expect(changeRows.first()).toContainText("Updated Sheet1!A1");

  await page.getByTestId("workbook-change-revert").click();

  await expect(formulaInput).toHaveValue("");
  await expect(changesTab).toContainText("2");
  await expect(changeRows).toHaveCount(2);
  await expect(changeRows.first()).toContainText("Reverted r1: Updated Sheet1!A1");
  await expect(changeRows.nth(1)).toContainText("Reverted in r2");
});

test("web app restores persisted workbook state after a full reload", async ({ page }) => {
  const documentId = `playwright-zero-reload-persist-${Date.now()}`;
  const formulaInput = page.getByTestId("formula-input");
  const resolvedValue = page.getByTestId("formula-resolved-value");

  await openZeroWorkbookPage(page, documentId);
  await expect(page.getByTestId("worker-error")).toHaveCount(0);

  await clickProductCell(page, 0, 0);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A1");
  await formulaInput.fill("17");
  await formulaInput.press("Enter");
  await expect(formulaInput).toHaveValue("17");
  await expect(resolvedValue).toHaveText("17");

  await page.reload({ waitUntil: "domcontentloaded" });
  await waitForWorkbookReady(page);
  await expect(page.getByTestId("worker-error")).toHaveCount(0);

  await clickProductCell(page, 0, 0);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A1");
  await expect(formulaInput).toHaveValue("17");
  await expect(resolvedValue).toHaveText("17");
});

test("web app keeps sheet tabs and status bar visible in a short viewport", async ({ page }) => {
  await page.setViewportSize({ width: 2048, height: 220 });
  await page.goto("/");
  await waitForWorkbookReady(page);

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
  await waitForWorkbookReady(page);

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
  await waitForWorkbookReady(page);

  await dragProductHeaderSelection(page, "column", 1, 3);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!B:D");

  await dragProductHeaderSelection(page, "row", 1, 3);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!2:4");
});

test("web app supports rectangular drag selection", async ({ page }) => {
  await page.goto("/");
  await waitForWorkbookReady(page);

  await dragProductBodySelection(page, 1, 1, 3, 3);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!B2:D4");
});

test("web app keeps the active focus inside the sheet grid when clicking a cell", async ({
  page,
}) => {
  await page.goto("/");
  await waitForWorkbookReady(page);

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
  await waitForWorkbookReady(page);

  await clickProductCellUpperHalf(page, 4, 11);
  await expect(page.getByTestId("name-box")).toHaveValue("E12");
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!E12");

  await clickProductCellUpperHalf(page, 2, 4);
  await expect(page.getByTestId("name-box")).toHaveValue("C5");
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!C5");
});

test("web app supports column resize without breaking hit testing", async ({ page }) => {
  await page.goto("/");
  await waitForWorkbookReady(page);

  await clickProductBodyOffset(page, 82, 0);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A1");

  await dragProductColumnResize(page, 0, -36);

  await clickProductBodyOffset(page, 82, 0);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!B1");
});

test("web app supports column edge double-click autofit", async ({ page }) => {
  await page.goto("/");
  await waitForWorkbookReady(page);

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
  await waitForWorkbookReady(page);

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
  await waitForWorkbookReady(page);

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
  await expect(cellEditor).toBeHidden();

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
  await expect(cellEditor).toBeHidden();

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
  await waitForWorkbookReady(page);

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
  await clickProductCell(page, 0, 0);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A1");
  await expect(formulaInput).toHaveValue("123");

  await clickProductCell(page, 1, 0);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!B1");
  await grid.press("4");
  await expect(cellEditor).toBeVisible();
  await expect(cellEditor).toHaveValue("4");
});

test("web app right-aligns numeric in-cell editing like numeric view state", async ({ page }) => {
  await page.goto("/");
  await waitForWorkbookReady(page);

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
  await waitForWorkbookReady(page);

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

  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A2");
  await clickProductCell(page, 0, 0);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A1");
  await expect(formulaInput).toHaveValue("123");
});

test("web app supports F2 edit in the product shell", async ({ page }) => {
  await page.goto("/");
  await waitForWorkbookReady(page);
  await waitForWorkbookReady(page);

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
  await waitForWorkbookReady(page);

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
  await waitForWorkbookReady(page);

  const nameBox = page.getByTestId("name-box");

  await nameBox.fill("C5");
  await nameBox.press("Enter");
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!C5");

  await clickProductSelectedCellTopBorder(page, 2, 4);
  await expect(nameBox).toHaveValue("C5");
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!C5");
});

test("web app keeps selected text cells visible when clicked", async ({ page }) => {
  await page.goto("/");
  await waitForWorkbookReady(page);

  const nameBox = page.getByTestId("name-box");
  const formulaInput = page.getByTestId("formula-input");
  const textOverlay = page.getByTestId("grid-text-overlay");
  const sampleText = "visible text sample";

  await nameBox.fill("C5");
  await nameBox.press("Enter");
  await formulaInput.fill(sampleText);
  await formulaInput.press("Enter");

  await clickProductCell(page, 0, 0);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A1");

  await clickProductCell(page, 2, 4);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!C5");
  const spilledText = textOverlay.getByText(sampleText, { exact: true });
  await expect(spilledText).toBeVisible();
  await expect(formulaInput).toHaveValue(sampleText);
  await expect
    .poll(() =>
      spilledText.evaluate((element) => {
        if (!(element instanceof HTMLElement)) {
          return 0;
        }
        return Math.round(element.getBoundingClientRect().width);
      }),
    )
    .toBeGreaterThan(PRODUCT_COLUMN_WIDTH);
});

test("web app supports fill-handle propagation", async ({ page }) => {
  await page.goto("/");
  await waitForWorkbookReady(page);

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

test("web app previews and fills rightward autofill like Sheets", async ({ page }) => {
  await page.goto("/");
  await waitForWorkbookReady(page);

  const nameBox = page.getByTestId("name-box");
  const formulaInput = page.getByTestId("formula-input");
  const resolvedValue = page.getByTestId("formula-resolved-value");
  const selectionStatus = page.getByTestId("status-selection");
  const fillPreview = page.locator("[data-grid-fill-preview='true']");

  await nameBox.fill("F6");
  await nameBox.press("Enter");
  await formulaInput.fill("7");
  await formulaInput.press("Enter");

  const { sourceX, sourceY, targetX, targetY } = await getProductFillHandleDragPoints(
    page,
    5,
    5,
    7,
    5,
  );
  await page.mouse.move(sourceX, sourceY);
  await page.mouse.down();
  await page.mouse.move(targetX, targetY, { steps: 10 });

  await expect(fillPreview).toBeVisible();
  await expect(fillPreview).toHaveCSS("border-top-style", "dashed");

  await page.mouse.up();

  await expect(selectionStatus).toContainText("!F6:H6");

  await nameBox.fill("H6");
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
  await waitForWorkbookReady(page);

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
  await waitForWorkbookReady(page);

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

  await nameBox.fill("D2");
  await nameBox.press("Enter");
  await expect(formulaInput).toHaveValue("3");
  await expect(resolvedValue).toHaveText("3");

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
  await waitForWorkbookReady(page);

  const baselineWidth = await getProductColumnWidth(page, 0);
  await dragProductColumnResize(page, 0, 48);
  await expect.poll(() => getProductColumnWidth(page, 0)).toBeGreaterThan(baselineWidth + 30);
});

test("web app shows #VALUE! for invalid formulas", async ({ page }) => {
  const documentId = `playwright-invalid-formula-${Date.now()}`;
  await openZeroWorkbookPage(page, documentId);

  const nameBox = page.getByTestId("name-box");
  const formulaInput = page.getByTestId("formula-input");
  const resolvedValue = page.getByTestId("formula-resolved-value");

  await nameBox.fill("A1");
  await nameBox.press("Enter");
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A1");

  await formulaInput.fill("=1+");
  await formulaInput.press("Enter");

  await expect(formulaInput).toHaveValue("#VALUE!");
  await expect(resolvedValue).toHaveText("#VALUE!");
});

test("web app commits in-cell string edits when clicking away", async ({ page }) => {
  await page.goto("/");
  await waitForWorkbookReady(page);

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

test("web app drags a selected range by its border with a grab cursor", async ({ page }) => {
  await page.goto("/");
  await waitForWorkbookReady(page);

  const nameBox = page.getByTestId("name-box");
  const formulaInput = page.getByTestId("formula-input");
  const resolvedValue = page.getByTestId("formula-resolved-value");

  await nameBox.fill("B2");
  await nameBox.press("Enter");
  await formulaInput.fill("left");
  await formulaInput.press("Enter");

  await nameBox.fill("C2");
  await nameBox.press("Enter");
  await formulaInput.fill("right");
  await formulaInput.press("Enter");

  await dragProductBodySelection(page, 1, 1, 2, 1);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!B2:C2");

  await dragProductSelectionBorder(page, 1, 1, 2, 1, 3, 3);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!D4:E4");

  await nameBox.fill("B2");
  await nameBox.press("Enter");
  await expect(formulaInput).toHaveValue("");
  await expect(resolvedValue).toHaveText("∅");

  await nameBox.fill("C2");
  await nameBox.press("Enter");
  await expect(formulaInput).toHaveValue("");
  await expect(resolvedValue).toHaveText("∅");

  await nameBox.fill("D4");
  await nameBox.press("Enter");
  await expect(formulaInput).toHaveValue("left");
  await expect(resolvedValue).toHaveText("left");

  await nameBox.fill("E4");
  await nameBox.press("Enter");
  await expect(formulaInput).toHaveValue("right");
  await expect(resolvedValue).toHaveText("right");
});

test("web app applies core formatting shortcuts from the keyboard", async ({ page }) => {
  await page.goto("/");
  await waitForWorkbookReady(page);

  const grid = page.getByTestId("sheet-grid");
  await clickProductCell(page, 0, 0);
  await grid.press(`${PRIMARY_MODIFIER}+B`);
  await expect(page.getByLabel("Bold")).toHaveClass(/bg-\[var\(--wb-accent-soft\)\]/);
  await grid.press(`${PRIMARY_MODIFIER}+I`);
  await expect(page.getByLabel("Italic")).toHaveClass(/bg-\[var\(--wb-accent-soft\)\]/);
  await grid.press(`${PRIMARY_MODIFIER}+U`);
  await expect(page.getByLabel("Underline")).toHaveClass(/bg-\[var\(--wb-accent-soft\)\]/);
  await grid.press(`${PRIMARY_MODIFIER}+Shift+E`);
  await expect(page.getByLabel("Align center")).toHaveClass(/bg-\[var\(--wb-accent-soft\)\]/);
  await grid.press(`${PRIMARY_MODIFIER}+Shift+R`);
  await expect(page.getByLabel("Align right")).toHaveClass(/bg-\[var\(--wb-accent-soft\)\]/);
  await grid.press(`${PRIMARY_MODIFIER}+Shift+L`);
  await expect(page.getByLabel("Align left")).toHaveClass(/bg-\[var\(--wb-accent-soft\)\]/);
});

test("web app supports row, column, and full-sheet selection shortcuts", async ({ page }) => {
  await page.goto("/");
  await waitForWorkbookReady(page);

  const grid = page.getByTestId("sheet-grid");
  await clickProductCell(page, 2, 4);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!C5");

  await grid.press("Shift+Space");
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!5:5");

  await grid.press(`${PRIMARY_MODIFIER}+Space`);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!C:C");

  await grid.press(`${PRIMARY_MODIFIER}+Shift+Space`);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!All");

  await grid.press(`${PRIMARY_MODIFIER}+A`);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!All");
});

test("web app expands the active range with repeated shift arrows", async ({ page }) => {
  await page.goto("/");
  await waitForWorkbookReady(page);

  const grid = page.getByTestId("sheet-grid");
  await clickProductCell(page, 2, 4);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!C5");

  await grid.press("Shift+ArrowRight");
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!C5:D5");

  await grid.press("Shift+ArrowRight");
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!C5:E5");

  await grid.press("Shift+ArrowDown");
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!C5:E6");
});

test("web app expands the active range with shift-click", async ({ page }) => {
  await page.goto("/");
  await waitForWorkbookReady(page);

  await clickProductCell(page, 1, 1);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!B2");

  await clickProductCell(page, 4, 5, { shift: true });
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!B2:E6");
});

for (const key of ["Delete", "Backspace"] as const) {
  test(`web app clears the full selected range with ${key.toLowerCase()}`, async ({
    page,
    context,
  }) => {
    await context.grantPermissions(["clipboard-read", "clipboard-write"]);
    await page.goto("/");
    await waitForWorkbookReady(page);

    const grid = page.getByTestId("sheet-grid");
    const formulaInput = page.getByTestId("formula-input");

    await clickProductCell(page, 1, 1);
    await page.evaluate(() => navigator.clipboard.writeText("11\t12\n13\t14"));
    await grid.press(`${PRIMARY_MODIFIER}+V`);
    await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!B2");

    await dragProductBodySelection(page, 1, 1, 2, 2);
    await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!B2:C3");

    await grid.press(key);

    await clickProductCell(page, 1, 1);
    await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!B2");
    await expect(formulaInput).toHaveValue("");
    await clickProductCell(page, 2, 1);
    await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!C2");
    await expect(formulaInput).toHaveValue("");
    await clickProductCell(page, 1, 2);
    await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!B3");
    await expect(formulaInput).toHaveValue("");
    await clickProductCell(page, 2, 2);
    await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!C3");
    await expect(formulaInput).toHaveValue("");
  });
}

test("web app ignores right gutter clicks", async ({ page }) => {
  await page.goto("/");
  await waitForWorkbookReady(page);

  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A1");
  await clickGridRightEdge(page, 3);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A1");
});

test("@fuzz-browser web app preserves valid selection geometry and focus under generated selection actions", async ({
  page,
}) => {
  test.skip(
    !fuzzBrowserEnabled || !shouldRunFuzzSuite("browser/grid-selection-focus", "browser"),
    "browser fuzz runs only in fuzz mode",
  );

  await runProperty({
    suite: "browser/grid-selection-focus",
    kind: "browser",
    arbitrary: fc.array(
      fc.oneof<BrowserSelectionAction>(
        fc.record({
          kind: fc.constant<"click">("click"),
          row: fc.integer({ min: 0, max: 8 }),
          col: fc.integer({ min: 0, max: 8 }),
        }),
        fc.record({
          kind: fc.constant<"shiftClick">("shiftClick"),
          row: fc.integer({ min: 0, max: 8 }),
          col: fc.integer({ min: 0, max: 8 }),
        }),
        fc.record({
          kind: fc.constant<"key">("key"),
          key: fc.constantFrom("ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown"),
          shift: fc.boolean(),
        }),
      ),
      { minLength: 6, maxLength: 10 },
    ),
    parameters: {
      interruptAfterTimeLimit: 40_000,
    },
    predicate: async (actions) => {
      await page.goto("/");
      await waitForWorkbookReady(page);
      const grid = page.getByTestId("sheet-grid");
      const nameBox = page.getByTestId("name-box");
      await expect(grid).toBeVisible({ timeout: 15_000 });
      await nameBox.fill("C5");
      await nameBox.press("Enter");
      await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!C5");
      await runSelectionFuzzActions(page, grid, actions);
    },
  });
});
