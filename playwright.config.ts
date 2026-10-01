import { defineConfig, devices } from 'playwright/test';

/**
 * End-to-end tests: the React demo in a real browser, driven like a user
 * (mouse, keyboard, file chooser, downloads).
 *
 * The specs in `e2e/` boot the demo's Vite dev server, which resolves both
 * workspace packages to their source, so no build is needed first. Set
 * `XLSX_E2E_PORT` to run on another port (two checkouts at once), and
 * `XLSX_E2E_CHANNEL` to pick the browser: the default is the installed
 * Google Chrome, so no browser download is required; use `chromium` for
 * Playwright's own build (`pnpm exec playwright install chromium`).
 */
const PORT = Number(process.env.XLSX_E2E_PORT ?? 4183);
const CHANNEL = process.env.XLSX_E2E_CHANNEL ?? 'chrome';
const isCI = Boolean(process.env.CI);

export default defineConfig({
	testDir: './e2e',
	testIgnore: ['**/support/**', '**/fixtures/**'],
	timeout: 60_000,
	expect: { timeout: 10_000 },
	fullyParallel: true,
	forbidOnly: isCI,
	retries: isCI ? 2 : 0,
	reporter: isCI ? [['list'], ['html', { open: 'never' }]] : 'list',
	use: {
		baseURL: `http://127.0.0.1:${PORT}`,
		trace: 'retain-on-failure',
		actionTimeout: 10_000,
	},
	projects: [
		{
			name: 'chrome',
			use: {
				...devices['Desktop Chrome'],
				channel: CHANNEL,
				viewport: { width: 1440, height: 900 },
				locale: 'en-US',
			},
		},
	],
	webServer: {
		command: `pnpm --filter xlsx-react-demo exec vite --port ${PORT} --strictPort --host 127.0.0.1`,
		url: `http://127.0.0.1:${PORT}`,
		reuseExistingServer: !isCI,
		timeout: 120_000,
	},
});
