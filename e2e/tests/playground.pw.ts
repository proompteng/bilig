import { expect, test } from "@playwright/test";

test("playground smoke exercises the custom renderer and wasm-backed recalculation", async ({ page }) => {
  await page.goto("/");

  await expect(page.getByRole("heading", { name: /custom reconciler playground/i })).toBeVisible();
  await expect(page.getByRole("button", { name: /^Cell B1$/ })).toHaveText("20");

  const formulaInput = page.getByLabel("Formula");
  await expect(formulaInput).toHaveValue("10");

  await formulaInput.fill("12");
  await page.getByRole("button", { name: "Commit" }).click();

  await expect(page.getByRole("button", { name: /^Cell B1$/ })).toHaveText("24");
  await expect(page.getByTestId("metric-wasm")).not.toHaveText("0");
});
