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
  await expect(page.getByTestId("replica-value")).toHaveText("12");
  await page.getByRole("tab", { name: "Sheet2" }).click();
  await expect(page.getByRole("button", { name: /^Cell A1$/ })).toHaveText("25");
  await page.getByTestId("sheet-grid").focus();
  await page.keyboard.press("ArrowRight");
  await expect(page.getByTestId("selection-chip")).toHaveText("Sheet2!B1");
  await expect(page.getByTestId("metric-js")).not.toHaveText("0");
  await expect(page.getByTestId("metric-wasm")).not.toHaveText("0");

  await page.reload();
  await expect(page.getByRole("button", { name: /^Cell A1$/ })).toHaveText("12");
  await expect(page.getByRole("button", { name: /^Cell B1$/ })).toHaveText("24");
  await expect(page.getByRole("button", { name: /^Cell C1$/ })).toHaveText("17");
  await expect(page.getByRole("button", { name: /^Cell D1$/ })).toHaveText("22");
  await expect(page.getByTestId("replica-value")).toHaveText("12");
});
