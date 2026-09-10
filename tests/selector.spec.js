const { expect, test } = require("@playwright/test");

test("starts muted and uses reduced audio volume", async ({ page }) => {
  await page.addInitScript(() => {
    const NativeAudio = window.Audio;
    window.__selectorAudioInstances = [];
    window.Audio = function TrackedAudio(...args) {
      const audio = new NativeAudio(...args);
      window.__selectorAudioInstances.push(audio);
      return audio;
    };
    window.Audio.prototype = NativeAudio.prototype;
  });

  await page.goto("/");

  const soundToggle = page.getByRole("button", { name: "Sound" });
  await expect(soundToggle).not.toHaveClass(/is-active/);
  await expect(page.getByText("Muted", { exact: true })).toBeVisible();

  await soundToggle.click();
  await page.getByRole("button", { name: "Play Intro" }).click();
  await expect.poll(() => page.evaluate(() => window.__selectorAudioInstances[0]?.volume)).toBe(0.15);
});

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

test("loads the 2026-27 S3 Economics rosters", async ({ page }) => {
  await page.goto("/");

  const expectedRosters = [
    { className: "S3.3", count: 38, firstStudent: "柏一鹏 Liam" },
    { className: "S3.4", count: 40, firstStudent: "许睿宁 Ryan" },
    { className: "S3.6", count: 33, firstStudent: "陈邢晔 Ron" }
  ];

  for (const { className, count, firstStudent } of expectedRosters) {
    await page.locator("[data-action='class']").selectOption(className);
    await expect(page.getByLabel("Selection stage").getByText(`${count} left | 0 graded`, { exact: true })).toBeVisible();

    await page.getByRole("button", { name: "Attendance" }).click();
    await expect(page.getByRole("dialog", { name: "Roll call" })).toContainText(`Student 1 of ${count}`);
    await expect(page.locator(".selector-rollcall-name")).toHaveText(firstStudent);
    await page.getByLabel("Close").click();
  }
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
