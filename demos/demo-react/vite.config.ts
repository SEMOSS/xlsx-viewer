import path from 'node:path';

import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import { defineConfig } from 'vite';

// The demo resolves both workspace packages to their *source*, so edits in
// packages/core/src and packages/react/src are live on reload without a build.
export default defineConfig({
	root: __dirname,
	base: process.env.DEMO_BASE ?? '/',
	plugins: [react(), tailwindcss()],
	publicDir: path.resolve(__dirname, 'public'),
	server: {
		port: 4183,
		strictPort: true,
	},
	preview: {
		port: 4184,
		strictPort: true,
	},
	build: {
		chunkSizeWarningLimit: 2500,
	},
	resolve: {
		alias: {
			'xlsx-viewer-core': path.resolve(__dirname, '../../packages/core/src/index.ts'),
			'xlsx-react-viewer': path.resolve(__dirname, '../../packages/react/src/index.ts'),
		},
	},
});
