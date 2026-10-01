import { resolve } from 'node:path';

import { defineConfig } from 'vite';

// Library build: one ESM entry, dependencies left external so consumers dedupe them.
export default defineConfig({
	build: {
		lib: {
			entry: resolve(__dirname, 'src/index.ts'),
			formats: ['es'],
			fileName: () => 'index.js',
		},
		outDir: 'dist',
		emptyOutDir: true,
		sourcemap: true,
		target: 'es2022',
		rollupOptions: {
			external: ['jszip'],
		},
	},
});
