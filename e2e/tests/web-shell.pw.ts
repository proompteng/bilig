import { expect, test } from "@playwright/test";

const PRODUCT_ROW_MARKER_WIDTH = 46;
const PRODUCT_COLUMN_WIDTH = 104;
const PRODUCT_HEADER_HEIGHT = 24;
const PRODUCT_ROW_HEIGHT = 22;

async function dragProductHeaderSelection(
  page: Parameters<typeof test>[0]["page"],
  axis: "column" | "row",
  startIndex: number,
  endIndex: number
) {
  const grid = await page.getByTestId("sheet-grid").boundingBox();
  if (!grid) {
    throw new Error("sheet grid is not visible");
  }

  const startX = axis === "column"
    ? grid.x + PRODUCT_ROW_MARKER_WIDTH + (startIndex * PRODUCT_COLUMN_WIDTH) + Math.floor(PRODUCT_COLUMN_WIDTH / 2)
    : grid.x + Math.floor(PRODUCT_ROW_MARKER_WIDTH / 2);
  const startY = axis === "column"
    ? grid.y + Math.floor(PRODUCT_HEADER_HEIGHT / 2)
    : grid.y + PRODUCT_HEADER_HEIGHT + (startIndex * PRODUCT_ROW_HEIGHT) + Math.floor(PRODUCT_ROW_HEIGHT / 2);
  const endX = axis === "column"
    ? grid.x + PRODUCT_ROW_MARKER_WIDTH + (endIndex * PRODUCT_COLUMN_WIDTH) + Math.floor(PRODUCT_COLUMN_WIDTH / 2)
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
  const grid = await page.getByTestId("sheet-grid").boundingBox();
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
  const grid = await page.getByTestId("sheet-grid").boundingBox();
  if (!grid) {
    throw new Error("sheet grid is not visible");
  }

  const sourceLeft = grid.x + PRODUCT_ROW_MARKER_WIDTH + (sourceCol * PRODUCT_COLUMN_WIDTH);
  const sourceTop = grid.y + PRODUCT_HEADER_HEIGHT + (sourceRow * PRODUCT_ROW_HEIGHT);
  const targetLeft = grid.x + PRODUCT_ROW_MARKER_WIDTH + (targetCol * PRODUCT_COLUMN_WIDTH);
  const targetTop = grid.y + PRODUCT_HEADER_HEIGHT + (targetRow * PRODUCT_ROW_HEIGHT);

  await page.mouse.move(sourceLeft + PRODUCT_COLUMN_WIDTH - 3, sourceTop + PRODUCT_ROW_HEIGHT - 3);
  await page.mouse.down();
  await page.mouse.move(targetLeft + PRODUCT_COLUMN_WIDTH - 3, targetTop + PRODUCT_ROW_HEIGHT - 3, { steps: 10 });
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

  await expect(page.getByTestId("status-mode")).toHaveText("Live");
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A1");
  await expect(page.getByTestId("status-sync")).toHaveText("Ready");
  await expect(page.locator(".formula-result-shell")).toHaveCount(0);
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

test("web app accepts string values and string comparison formulas", async ({ page }) => {
  await page.goto("/");

  const nameBox = page.getByTestId("name-box");
  const formulaInput = page.getByTestId("formula-input");
  const resolvedValue = page.getByTestId("formula-resolved-value");

  await formulaInput.fill("hello");
  await formulaInput.press("Enter");
  await expect(nameBox).toHaveValue("A1");
  await expect(resolvedValue).toHaveText("hello");

  await nameBox.fill("A2");
  await nameBox.press("Enter");
  await formulaInput.fill("=A1=\"HELLO\"");
  await formulaInput.press("Enter");
  await expect(resolvedValue).toHaveText("TRUE");
});

test("web app supports fill-handle propagation", async ({ page }) => {
  await page.goto("/");

  const nameBox = page.getByTestId("name-box");
  const formulaInput = page.getByTestId("formula-input");
  const resolvedValue = page.getByTestId("formula-resolved-value");

  await formulaInput.fill("7");
  await formulaInput.press("Enter");

  await dragProductFillHandle(page, 0, 0, 0, 2);

  await nameBox.fill("A3");
  await nameBox.press("Enter");
  await expect(formulaInput).toHaveValue("7");
  await expect(resolvedValue).toHaveText("7");
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

  await grid.focus();
  await grid.press("F2");
  await expect(page.getByTestId("cell-editor-input")).toBeVisible();
  await page.keyboard.type("hello");
  await grid.click({
    position: {
      x: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH + Math.floor(PRODUCT_COLUMN_WIDTH / 2),
      y: PRODUCT_HEADER_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2)
    }
  });

  await expect(nameBox).toHaveValue("B1");
  await nameBox.focus();
  await nameBox.selectText();
  await page.keyboard.type("A1");
  await nameBox.press("Enter");
  await expect(formulaInput).toHaveValue("hello");
  await expect(resolvedValue).toHaveText("hello");
});

test("web app ignores right gutter clicks", async ({ page }) => {
  await page.goto("/");

  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A1");
  await clickGridRightEdge(page, 3);
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A1");
});
