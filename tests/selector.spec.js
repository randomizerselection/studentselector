const { expect, test } = require("@playwright/test");

const DEMO_ROSTER = "class,student\nDemo Class,Student One\nDemo Class,Student Two\nDemo Class,Student Three\n";

async function loadRoster(page, csv = DEMO_ROSTER) {
  await page.getByLabel("Load roster CSV", { exact: true }).setInputFiles({
    name: "students.csv",
    mimeType: "text/csv",
    buffer: Buffer.from(csv)
  });
  await expect(page.getByLabel("Load roster CSV", { exact: true })).toBeEnabled();
}

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
  await loadRoster(page);

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

test("imports locally without any roster network request and survives refresh", async ({ page }) => {
  const rosterRequests = [];
  page.on("request", (request) => {
    if (request.url().includes("students.csv")) rosterRequests.push(request.url());
  });
  await page.goto("/");
  await expect(page.locator("[data-action='class'] option")).toHaveCount(1);
  await loadRoster(page);
  await page.locator("[data-action='class']").selectOption("Demo Class");
  await expect(page.getByLabel("Selection stage").getByText("3 left | 0 graded", { exact: true })).toBeVisible();
  await page.reload();
  await expect(page.locator("[data-action='class']")).toHaveValue("Demo Class");
  await expect(page.getByLabel("Selection stage").getByText("3 left | 0 graded", { exact: true })).toBeVisible();
  expect(rosterRequests).toEqual([]);
});

test("handles invalid files and repeated imports without losing a valid roster", async ({ page }) => {
  await page.goto("/");
  await loadRoster(page);
  await page.locator("[data-action='class']").selectOption("Demo Class");
  await loadRoster(page, "class,student\n");
  await expect(page.getByRole("dialog", { name: "Unable to load roster" })).toBeVisible();
  await page.getByLabel("Close", { exact: true }).click();
  await expect(page.locator("[data-action='class']")).toHaveValue("Demo Class");
  await loadRoster(page);
  await expect(page.locator("[data-action='class']")).toHaveValue("Demo Class");
  await loadRoster(page, "class,student\nOther Class,Student Four\n");
  await expect(page.locator("[data-action='class']")).toHaveValue("");
  await expect(page.locator("[data-action='class'] option")).toHaveCount(2);
  await page.locator("[data-action='class']").selectOption("Other Class");
  await expect(page.getByLabel("Selection stage").getByText("1 left | 0 graded", { exact: true })).toBeVisible();
});

test("uses imported roster in a reopened overlay", async ({ page }) => {
  await page.goto("/");
  await loadRoster(page);
  await page.evaluate(() => { window.__selectorOverlay = window.StudentSelector.open(); });
  const overlay = page.locator(".selector-overlay-host");
  await expect(overlay.locator("[data-action='class'] option")).toHaveCount(2);
  await overlay.getByLabel("Close selector").click();
  await expect(overlay).toHaveCount(0);
  await page.evaluate(() => { window.__selectorOverlay = window.StudentSelector.open(); });
  await expect(overlay.locator("[data-action='class'] option")).toHaveCount(2);
  await overlay.getByLabel("Close selector").click();
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

