const { expect, test } = require("@playwright/test");

test("loads roster and completes the core selector flow", async ({ page }) => {
  await page.goto("/");

  await expect(page.getByRole("heading", { name: "Random Student Selector" })).toBeVisible();

  const classSelect = page.locator("[data-action='class']");
  await expect(classSelect).toBeVisible();
  await expect(classSelect.locator("option").first()).toHaveCSS("color", "rgb(7, 17, 31)");
  await expect(classSelect.locator("option").first()).toHaveCSS("background-color", "rgb(255, 255, 255)");
  const options = await classSelect.locator("option").allTextContents();
  const className = options.find((option) => option !== "Select a Class");
  expect(className).toBeTruthy();

  await classSelect.selectOption({ label: className });
  await expect(page.getByLabel("Selection stage").getByText(/\d+ left \| 0 graded/)).toBeVisible();

  await page.getByRole("button", { name: "Attendance" }).click();
  await expect(page.getByRole("dialog", { name: "Roll call" })).toBeVisible();
  await page.keyboard.press("Enter");
  await page.keyboard.press("KeyA");
  await page.getByLabel("Close").click();

  await page.getByRole("button", { name: "Slot Effect" }).click();
  const startButton = page.getByRole("button", { name: "START SELECTION" });
  await startButton.hover();
  await expect(startButton).toHaveCSS("color", "rgb(255, 250, 243)");
  await expect(startButton).toHaveCSS("background-color", "rgb(133, 77, 34)");
  await startButton.click();

  await expect(page.locator("[data-current-name]")).toHaveText("Get ready");
  await expect(page.getByRole("button", { name: /A\*|Excellent/ })).toBeVisible({ timeout: 7_000 });
  await page.getByRole("button", { name: /A Strong/ }).click();
  await expect(page.getByRole("dialog", { name: "Feedback" })).toBeVisible();
  await page.getByRole("button", { name: "Return To Dock" }).click();
  await page.getByRole("button", { name: "View Summary" }).click();
  await expect(page.getByRole("dialog", { name: "Session summary" })).toBeVisible();
  await expect(page.locator(".selector-summary-row")).toHaveCount(1);
});

test("exposes the reusable lesson overlay API", async ({ page }) => {
  await page.goto("/");

  const hasApi = await page.evaluate(() => {
    return Boolean(window.StudentSelector?.mount && window.StudentSelector?.open);
  });

  expect(hasApi).toBe(true);

  await page.evaluate(() => {
    window.__selectorOverlay = window.StudentSelector.open();
  });

  await expect(page.locator(".selector-overlay-host")).toBeVisible();
  await expect(page.locator(".selector-overlay-host").getByRole("heading", { name: "Random Student Selector" })).toBeVisible();

  await page.evaluate(() => window.__selectorOverlay.close());
  await expect(page.locator(".selector-overlay-host")).toHaveCount(0);
});
