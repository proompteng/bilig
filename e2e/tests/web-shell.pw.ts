import { expect, test } from "@playwright/test";

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
});

test("web app supports column and row header selection", async ({ page }) => {
  await page.goto("/");

  const grid = page.getByTestId("sheet-grid");

  await grid.click({ position: { x: 60 + 120 + 60, y: 15 } });
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!B:B");

  await grid.click({ position: { x: 30, y: 30 + 28 + 14 } });
  await expect(page.getByTestId("status-selection")).toHaveText("Sheet1!2:2");
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

test("web app shows #VALUE! for invalid formulas", async ({ page }) => {
  await page.goto("/");

  const formulaInput = page.getByTestId("formula-input");
  const resolvedValue = page.getByTestId("formula-resolved-value");

  await formulaInput.fill("=1+");
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
  await grid.click({ position: { x: 60 + 120 + 60, y: 30 + 14 } });

  await expect(nameBox).toHaveValue("B1");
  await nameBox.fill("A1");
  await nameBox.press("Enter");
  await expect(formulaInput).toHaveValue("hello");
  await expect(resolvedValue).toHaveText("hello");
});
