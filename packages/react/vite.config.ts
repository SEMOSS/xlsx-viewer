import { resolve } from 'node:path';

import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import { defineConfig } from 'vite';

// Library build: JS entry plus the Tailwind-compiled stylesheet (dist/xlsx-viewer.css).
// React, ReactDOM and the core engine stay external so the host app dedupes them.
export default defineConfig({
	plugins: [react(), tailwindcss()],
	build: {
		lib: {
			entry: resolve(__dirname, 'src/index.ts'),
			formats: ['es'],
			fileName: () => 'index.js',
			cssFileName: 'xlsx-viewer',
		},
		outDir: 'dist',
		emptyOutDir: true,
		sourcemap: true,
		target: 'es2022',
		rollupOptions: {
			external: [
				'react',
				'react-dom',
				'react/jsx-runtime',
				'react-dom/client',
				'xlsx-viewer-core',
				'jszip',
				'lucide-react',
				'clsx',
			],
		},
	},
});
