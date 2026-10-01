import { resolve } from 'node:path';

import react from '@vitejs/plugin-react';
import { defineConfig } from 'vitest/config';

export default defineConfig({
	plugins: [react()],
	resolve: {
		alias: {
			'xlsx-viewer-core': resolve(__dirname, '../core/src/index.ts'),
		},
	},
	test: {
		include: ['src/**/*.test.ts', 'src/**/*.test.tsx'],
		environment: 'jsdom',
		setupFiles: ['./src/test-setup.ts'],
		testTimeout: 20000,
	},
});
