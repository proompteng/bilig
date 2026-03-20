import { expect, test } from "@playwright/test";

const PRODUCT_ROW_MARKER_WIDTH = 46;
const PRODUCT_COLUMN_WIDTH = 104;
const PRODUCT_HEADER_HEIGHT = 24;
const PRODUCT_ROW_HEIGHT = 22;
const PRIMARY_MODIFIER = process.platform === "darwin" ? "Meta" : "Control";

async function getProductColumnWidth(page: Parameters<typeof test>[0]["page"], columnIndex: number) {
  const grid = page.getByTestId("sheet-grid");
  const [defaultWidthRaw, overridesRaw] = await Promise.all([
    grid.getAttribute("data-default-column-width"),
    grid.getAttribute("data-column-width-overrides")
  ]);
  const defaultWidth = Number(defaultWidthRaw ?? String(PRODUCT_COLUMN_WIDTH));
  const overrides = overridesRaw ? JSON.parse(overridesRaw) as Record<string, number> : {};
  return overrides[String(columnIndex)] ?? defaultWidth;
}

async function getProductColumnLeft(page: Parameters<typeof test>[0]["page"], columnIndex: number) {
  let offset = PRODUCT_ROW_MARKER_WIDTH;
  for (let index = 0; index < columnIndex; index += 1) {
    offset += await getProductColumnWidth(page, index);
  }
  return offset;
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

async function dragProductColumnDivider(
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

  const dividerX = grid.x + await getProductColumnLeft(page, columnIndex) + await getProductColumnWidth(page, columnIndex);
  const y = grid.y + Math.floor(PRODUCT_HEADER_HEIGHT / 2);

  await page.mouse.move(dividerX - 1, y);
  await page.mouse.down();
  await page.mouse.move(dividerX - 1 + deltaX, y, { steps: 8 });
  await page.mouse.up();
}

async function doubleClickProductColumnDivider(
  page: Parameters<typeof test>[0]["page"],
  columnIndex: number
) {
  const gridLocator = page.getByTestId("sheet-grid");
  await expect(gridLocator).toBeVisible();
  const grid = await gridLocator.boundingBox();
  if (!grid) {
    throw new Error("sheet grid is not visible");
  }

  const dividerX = grid.x + await getProductColumnLeft(page, columnIndex) + await getProductColumnWidth(page, columnIndex);
  const y = grid.y + Math.floor(PRODUCT_HEADER_HEIGHT / 2);
  await page.mouse.dblclick(dividerX - 1, y);
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
