import { existsSync } from 'node:fs';
import { defineConfig } from '@playwright/test';

const chrome = [
  process.env.VGHTPE_TEST_BROWSER,
  (process.env.ProgramFiles || 'C:/Program Files') + '/Google/Chrome/Application/chrome.exe',
  (process.env['ProgramFiles(x86)'] || 'C:/Program Files (x86)') + '/Google/Chrome/Application/chrome.exe',
].find((candidate): candidate is string => Boolean(candidate && existsSync(candidate)));

export default defineConfig({
  testDir: './e2e',
  globalSetup: './e2e/setup.ts',
  fullyParallel: false,
  workers: 1,
  reporter: 'list',
  timeout: 30_000,
  use: {
    baseURL: 'http://127.0.0.1:18865',
    viewport: { width: 1280, height: 900 },
    launchOptions: process.env.CI ? {} : { executablePath: chrome },
    trace: 'retain-on-failure',
    screenshot: 'only-on-failure',
  },
});
