import { expect, test } from "@playwright/test";

test("web app renders the minimal product shell without playground chrome", async ({ page }) => {
  await page.goto("/");

  await expect(page.getByTestId("formula-bar")).toBeVisible();
  await expect(page.getByTestId("name-box")).toBeVisible();
  await expect(page.getByTestId("sheet-grid")).toBeVisible();
  await expect(page.getByRole("tab", { name: "Sheet1" })).toBeVisible();

  await expect(page.getByTestId("preset-strip")).toHaveCount(0);
  await expect(page.getByTestId("metrics-panel")).toHaveCount(0);
  await expect(page.getByTestId("replica-panel")).toHaveCount(0);

  await expect(page.getByTestId("status-mode")).toHaveText("Live");
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!A1");
  await expect(page.getByTestId("status-sync")).toHaveText("Ready");
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
