import { expect, test, type Page } from "@playwright/test";

async function clearWorkspace(page: Page) {
  await page.goto("/");
  await page.evaluate(() => {
    window.localStorage.clear();
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

test("playground shell supports formula-bar navigation, in-grid editing, and Excel-scale presets", async ({ page }) => {
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
  await expect(statusBar.getByTestId("metric-js")).not.toContainText("JS 0");
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
