import { expect, test } from "@playwright/test";

test("playground smoke exercises the custom renderer and wasm-backed recalculation", async ({ page }) => {
  await page.goto("/");
  await page.evaluate(() => {
    window.localStorage.clear();
  });
  await page.reload();

  await expect(page.getByRole("heading", { name: /custom reconciler playground/i })).toBeVisible();
  await expect(page.getByRole("button", { name: /^Cell B1$/ })).toHaveText("20");
  await expect(page.getByRole("button", { name: /^Cell C1$/ })).toHaveText("15");
  await expect(page.getByRole("button", { name: /^Cell D1$/ })).toHaveText("20");

  const formulaInput = page.getByLabel("Formula");
  await expect(formulaInput).toHaveValue("10");

  await formulaInput.fill("12");
  await page.getByRole("button", { name: "Commit" }).click();

  await expect(page.getByRole("button", { name: /^Cell B1$/ })).toHaveText("24");
  await expect(page.getByRole("button", { name: /^Cell C1$/ })).toHaveText("17");
  await expect(page.getByRole("button", { name: /^Cell D1$/ })).toHaveText("22");
  await expect(page.getByTestId("metric-js")).not.toHaveText("0");
  await expect(page.getByTestId("metric-wasm")).not.toHaveText("0");
  await expect(page.getByTestId("replica-value")).toHaveText("12");
  await page.getByRole("tab", { name: "Sheet2" }).click();
  await expect(page.getByRole("button", { name: /^Cell A1$/ })).toHaveText("25");
  const grid = page.getByTestId("sheet-grid");
  await grid.focus();
  await grid.press("ArrowRight");
  await expect(page.getByTestId("selection-chip")).toHaveText("Sheet2!B1");
  const cellB1 = page.getByRole("button", { name: /^Cell B1$/ });
  await cellB1.dblclick();
  await expect(page.getByTestId("cell-editor-overlay")).toBeVisible();
  const overlayInput = page.getByLabel("Sheet2!B1 editor");
  await overlayInput.fill("44");
  await overlayInput.press("Enter");
  await expect(page.getByTestId("cell-editor-overlay")).toHaveCount(0);
  await expect(cellB1).toHaveText("44");
  await cellB1.dblclick();
  await expect(page.getByTestId("cell-editor-overlay")).toBeVisible();
  await overlayInput.fill("99");
  await overlayInput.press("Escape");
  await expect(page.getByTestId("cell-editor-overlay")).toHaveCount(0);
  await expect(cellB1).toHaveText("44");

  await page.reload();
  await expect(page.getByRole("button", { name: /^Cell A1$/ })).toHaveText("12");
  await expect(page.getByRole("button", { name: /^Cell B1$/ })).toHaveText("24");
  await expect(page.getByRole("button", { name: /^Cell C1$/ })).toHaveText("17");
  await expect(page.getByRole("button", { name: /^Cell D1$/ })).toHaveText("22");
  await expect(page.getByTestId("replica-value")).toHaveText("12");
});

test("paused relay queue survives reload and resumes replication", async ({ page }) => {
  await page.goto("/");
  await page.evaluate(() => {
    window.localStorage.clear();
  });
  await page.reload();

  const formulaInput = page.getByLabel("Formula");

  await expect(page.getByTestId("replica-status")).toHaveText("Live");
  await expect(page.getByTestId("replica-value")).toHaveText("10");

  await page.getByRole("button", { name: "Pause sync" }).click();
  await expect(page.getByTestId("replica-status")).toHaveText("Paused");

  await formulaInput.fill("15");
  await page.getByRole("button", { name: "Commit" }).click();

  await expect(page.getByRole("button", { name: /^Cell A1$/ })).toHaveText("15");
  await expect(page.getByTestId("replica-value")).toHaveText("10");
  await expect(page.getByTestId("replica-queued")).toHaveText("1");

  await formulaInput.fill("18");
  await page.getByRole("button", { name: "Commit" }).click();

  await expect(page.getByRole("button", { name: /^Cell A1$/ })).toHaveText("18");
  await expect(page.getByTestId("replica-value")).toHaveText("10");
  await expect(page.getByTestId("replica-queued")).toHaveText("1");

  await page.reload();

  await expect(page.getByTestId("replica-status")).toHaveText("Paused");
  await expect(page.getByRole("button", { name: /^Cell A1$/ })).toHaveText("18");
  await expect(page.getByTestId("replica-value")).toHaveText("10");
  await expect(page.getByTestId("replica-queued")).toHaveText("1");

  await page.getByRole("button", { name: "Resume sync" }).click();

  await expect(page.getByTestId("replica-status")).toHaveText("Live");
  await expect(page.getByTestId("replica-value")).toHaveText("18");
  await expect(page.getByTestId("replica-queued")).toHaveText("0");
});
