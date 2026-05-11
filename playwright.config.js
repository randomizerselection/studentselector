const { defineConfig } = require("@playwright/test");

module.exports = defineConfig({
  testDir: "./tests",
  timeout: 30_000,
  use: {
    baseURL: "http://127.0.0.1:8770",
    trace: "on-first-retry"
  },
  webServer: {
    command: "python -m http.server 8770 --bind 127.0.0.1",
    url: "http://127.0.0.1:8770",
    reuseExistingServer: true,
    stdout: "ignore",
    stderr: "pipe"
  }
});
