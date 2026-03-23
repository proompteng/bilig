import { expect, test, type Page } from "@playwright/test";

async function clearWorkspace(page: Page) {
  await page.goto("/");
  await page.evaluate(async () => {
    window.localStorage.clear();
    await new Promise<void>((resolve) => {
      const request = window.indexedDB.deleteDatabase("bilig-playground");
      request.addEventListener("success", () => resolve(), { once: true });
      request.addEventListener("error", () => resolve(), { once: true });
      request.addEventListener("blocked", () => resolve(), { once: true });
    });
  });
  await page.reload();
  await loadPreset(page, "starter", "Starter Demo");
  await expect(page.getByTestId("selection-chip")).toHaveText("Sheet1!A1");
  await expect(page.getByTestId("formula-input")).toHaveValue("10");
}

async function jumpTo(page: Page, target: string) {
  const nameBox = page.getByTestId("name-box");
  await nameBox.fill(target);
  await nameBox.press("Enter");
}

async function loadPreset(page: Page, presetId: string, label?: string) {
  await page.getByTestId(`preset-${presetId}`).click();
  const loadingBanner = page.getByTestId("preset-loading");
  await page.waitForTimeout(50);
  if ((await loadingBanner.count()) > 0) {
    await expect(loadingBanner).toHaveCount(0);
  }
  if (label) {
    await expect(page.getByTestId("status-active-preset")).toHaveText(label);
  }
}

async function clickVisibleCell(page: Page, colIndex: number, rowIndex: number) {
  const grid = await page.getByTestId("sheet-grid").boundingBox();
  if (!grid) {
    throw new Error("sheet grid is not visible");
  }

  const rowMarkerWidth = 60;
  const headerHeight = 30;
  const columnWidth = 120;
  const rowHeight = 28;
  const x = grid.x + rowMarkerWidth + colIndex * columnWidth + 24;
  const y = grid.y + headerHeight + rowIndex * rowHeight + Math.floor(rowHeight / 2);
  await page.mouse.click(x, y);
}

async function dragVisibleSelection(
  page: Page,
  startCol: number,
  startRow: number,
  endCol: number,
  endRow: number,
) {
  const grid = await page.getByTestId("sheet-grid").boundingBox();
  if (!grid) {
    throw new Error("sheet grid is not visible");
  }

  const rowMarkerWidth = 60;
  const headerHeight = 30;
  const columnWidth = 120;
  const rowHeight = 28;
  const startX = grid.x + rowMarkerWidth + startCol * columnWidth + 24;
  const startY = grid.y + headerHeight + startRow * rowHeight + Math.floor(rowHeight / 2);
  const endX = grid.x + rowMarkerWidth + endCol * columnWidth + 24;
  const endY = grid.y + headerHeight + endRow * rowHeight + Math.floor(rowHeight / 2);

  await page.mouse.move(startX, startY);
  await page.mouse.down();
  await page.mouse.move(endX, endY, { steps: 8 });
  await page.mouse.up();
}

async function clickGridRightEdge(page: Page, rowIndex = 2) {
  const grid = await page.getByTestId("sheet-grid").boundingBox();
  if (!grid) {
    throw new Error("sheet grid is not visible");
  }

  const headerHeight = 30;
  const rowHeight = 28;
  const x = grid.x + grid.width - 3;
  const y = grid.y + headerHeight + rowIndex * rowHeight + Math.floor(rowHeight / 2);
  await page.mouse.click(x, y);
}

test("playground shell supports formula-bar navigation, in-grid editing, and Excel-scale presets", async ({
  page,
}) => {
  await clearWorkspace(page);

  await expect(page.getByRole("heading", { name: "bilig-demo" })).toBeVisible();
  await expect(page.getByTestId("selection-chip")).toHaveText("Sheet1!A1");
  await expect(page.getByTestId("name-box")).toHaveValue("A1");
  await expect(page.getByTestId("formula-input")).toHaveValue("10");

  const gridHost = page.getByTestId("sheet-grid");
  await expect(gridHost).toBeVisible();
  await gridHost.focus();
  await gridHost.press("F2");

  const overlay = page.getByTestId("cell-editor-overlay");
  await expect(overlay).toBeVisible();
  const overlayInput = overlay.locator("input");
  await overlayInput.fill("12");
  await overlayInput.press("Enter");

  await expect(overlay).toHaveCount(0);
  await expect(page.getByTestId("selection-chip")).toHaveText("Sheet1!A2");
  await expect(page.getByTestId("formula-input")).toHaveValue("5");

  await jumpTo(page, "B1");
  await expect(page.getByTestId("selection-chip")).toHaveText("Sheet1!B1");
  await expect(page.getByTestId("formula-input")).toHaveValue("=A1*2");
  await expect(page.getByTestId("formula-resolved-value")).toContainText("24");

  await jumpTo(page, "C1");
  await expect(page.getByTestId("formula-resolved-value")).toContainText("17");
  await jumpTo(page, "D1");
  await expect(page.getByTestId("formula-resolved-value")).toContainText("22");
  const statusBar = page.locator(".workbook-status");
  await expect(statusBar.getByTestId("metric-js")).toContainText("Fallback 0");
  await expect(statusBar.getByTestId("metric-wasm")).not.toContainText("WASM 0");
  await expect(page.getByTestId("replica-value")).toHaveText("22");

  await page.getByRole("tab", { name: "Sheet2" }).click();
  await expect(page.getByTestId("selection-chip")).toHaveText("Sheet2!A1");
  await expect(page.getByTestId("formula-resolved-value")).toContainText("25");

  await loadPreset(page, "million-surface", "Million-Row Surface");
  await expect(page.getByTestId("status-active-preset")).toHaveText("Million-Row Surface");
  await expect(page.getByText("1,048,576 rows x 16,384 columns")).toBeVisible();

  await jumpTo(page, "A1048576");
  await expect(page.getByTestId("formula-input")).toHaveValue("1048576");
  await expect(page.getByTestId("formula-resolved-value")).toContainText("1048576");

  await jumpTo(page, "XFD1048576");
  await expect(page.getByTestId("formula-input")).toHaveValue("=B1048576+1");
  await expect(page.getByTestId("formula-resolved-value")).toContainText("2097153");
});

test("cross-sheet formulas recalculate end to end after editing the source sheet", async ({
  page,
}) => {
  await clearWorkspace(page);

  const gridHost = page.getByTestId("sheet-grid");
  await gridHost.focus();
  await gridHost.press("F2");

  const overlay = page.getByTestId("cell-editor-overlay");
  await expect(overlay).toBeVisible();
  await page.getByTestId("cell-editor-input").fill("12");
  await page.getByTestId("cell-editor-input").press("Enter");

  await expect(page.getByTestId("selection-chip")).toHaveText("Sheet1!A2");
  await page.getByRole("tab", { name: "Sheet2" }).click();
  await expect(page.getByTestId("selection-chip")).toHaveText("Sheet2!A1");
  await expect(page.getByTestId("formula-input")).toHaveValue(
    "=IF(Sheet1!B1>20,Sheet1!B1+1,Sheet1!B2-1)",
  );
  await expect(page.getByTestId("formula-resolved-value")).toContainText("25");
});

test("pointer clicks select the visible cell instead of an offset row", async ({ page }) => {
  await clearWorkspace(page);

  await clickVisibleCell(page, 0, 0);
  await expect(page.getByTestId("selection-chip")).toHaveText("Sheet1!A1");

  await clickVisibleCell(page, 0, 4);
  await expect(page.getByTestId("selection-chip")).toHaveText("Sheet1!A5");

  await clickVisibleCell(page, 1, 1);
  await expect(page.getByTestId("selection-chip")).toHaveText("Sheet1!B2");
});

test("dragging selects a rectangular range like a spreadsheet", async ({ page }) => {
  await clearWorkspace(page);

  await dragVisibleSelection(page, 0, 0, 2, 2);
  await expect(page.getByTestId("selection-chip")).toHaveText("Sheet1!A1:C3");

  await dragVisibleSelection(page, 1, 1, 3, 4);
  await expect(page.getByTestId("selection-chip")).toHaveText("Sheet1!B2:D5");
});

test("keyboard shortcuts work for navigation, editing, cancel, delete, and formula bar commit", async ({
  page,
}) => {
  await clearWorkspace(page);

  const gridHost = page.getByTestId("sheet-grid");
  await gridHost.focus();

  await gridHost.press("ArrowRight");
  await expect(page.getByTestId("selection-chip")).toHaveText("Sheet1!B1");

  await gridHost.press("ArrowDown");
  await expect(page.getByTestId("selection-chip")).toHaveText("Sheet1!B2");

  await gridHost.press("ArrowLeft");
  await expect(page.getByTestId("selection-chip")).toHaveText("Sheet1!A2");

  await gridHost.press("ArrowUp");
  await expect(page.getByTestId("selection-chip")).toHaveText("Sheet1!A1");

  await gridHost.press("Tab");
  await expect(page.getByTestId("selection-chip")).toHaveText("Sheet1!B1");

  await gridHost.press("Shift+Tab");
  await expect(page.getByTestId("selection-chip")).toHaveText("Sheet1!A1");

  await gridHost.press("9");
  const overlay = page.getByTestId("cell-editor-overlay");
  await expect(overlay).toBeVisible();
  await expect(page.getByTestId("cell-editor-input")).toHaveValue("9");
  await page.getByTestId("cell-editor-input").press("Escape");
  await expect(overlay).toHaveCount(0);
  await expect(page.getByTestId("formula-input")).toHaveValue("10");

  await gridHost.focus();
  await gridHost.press("Delete");
  await expect(page.getByTestId("formula-input")).toHaveValue("");

  await gridHost.press("F2");
  await expect(overlay).toBeVisible();
  await page.getByTestId("cell-editor-input").fill("11");
  await page.getByTestId("cell-editor-input").press("Tab");
  await expect(overlay).toHaveCount(0);
  await expect(page.getByTestId("selection-chip")).toHaveText("Sheet1!B1");

  await jumpTo(page, "A1");
  await expect(page.getByTestId("formula-input")).toHaveValue("11");

  const formulaInput = page.getByTestId("formula-input");
  await formulaInput.focus();
  await formulaInput.press("Control+A");
  await formulaInput.fill("=7*6");
  await formulaInput.press("Enter");
  await expect(page.getByTestId("formula-input")).toHaveValue("=7*6");
  await expect(page.getByTestId("formula-resolved-value")).toContainText("42");

  await jumpTo(page, "B1");
  await expect(page.getByTestId("formula-input")).toHaveValue("=A1*2");
  await expect(page.getByTestId("formula-resolved-value")).toContainText("84");
});

test("clicking the right scrollbar gutter does not select the last visible column", async ({
  page,
}) => {
  await clearWorkspace(page);

  await expect(page.getByTestId("selection-chip")).toHaveText("Sheet1!A1");
  await clickGridRightEdge(page, 3);
  await expect(page.getByTestId("selection-chip")).toHaveText("Sheet1!A1");
});

test("paused relay queue survives reload and resumes replication", async ({ page }) => {
  await clearWorkspace(page);

  await expect(page.getByTestId("replica-status")).toHaveText("Live");
  await expect(page.getByTestId("replica-value")).toHaveText("10");

  await page.getByRole("button", { name: "Pause sync" }).click();
  await expect(page.getByTestId("replica-status")).toHaveText("Paused");

  const formulaInput = page.getByTestId("formula-input");
  await formulaInput.fill("18");
  await page.getByRole("button", { name: "Commit" }).click();

  await expect(page.getByTestId("selection-chip")).toHaveText("Sheet1!A1");
  await expect(page.getByTestId("formula-input")).toHaveValue("18");
  await expect(page.getByTestId("replica-value")).toHaveText("10");
  await expect(page.getByTestId("replica-queued")).not.toHaveText("0");

  await page.reload();

  await expect(page.getByTestId("replica-status")).toHaveText("Paused");
  await expect(page.getByTestId("formula-input")).toHaveValue("18");
  await expect(page.getByTestId("replica-value")).toHaveText("10");
  await expect(page.getByTestId("replica-queued")).not.toHaveText("0");

  await page.getByRole("button", { name: "Resume sync" }).click();

  await expect(page.getByTestId("replica-status")).toHaveText("Live");
  await expect(page.getByTestId("replica-value")).toHaveText("18");
  await expect(page.getByTestId("replica-queued")).toHaveText("0");
});

test("large presets persist through reload without storage quota failures", async ({ page }) => {
  await clearWorkspace(page);

  await loadPreset(page, "load-250k", "250k Materialized");
  await expect(page.getByTestId("preset-error")).toHaveCount(0);

  await jumpTo(page, "A125000");
  await expect(page.getByTestId("formula-input")).toHaveValue("125000");

  await jumpTo(page, "B125000");
  await expect(page.getByTestId("formula-input")).toHaveValue("=A125000*2");

  await page.reload();

  await expect(page.getByTestId("preset-error")).toHaveCount(0);
  await expect(page.getByTestId("status-active-preset")).toHaveText("Restored workspace");
  await jumpTo(page, "A125000");
  await expect(page.getByTestId("formula-input")).toHaveValue("125000");
  await jumpTo(page, "B125000");
  await expect(page.getByTestId("formula-input")).toHaveValue("=A125000*2");
});
