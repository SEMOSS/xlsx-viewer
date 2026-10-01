import { afterEach, beforeEach, vi } from 'vitest';

/**
 * jsdom has no layout: give the grid a ResizeObserver and the sheet tabs
 * `scrollIntoView` so a whole `<ExcelViewer>` can render in tests.
 */
export function stubLayout(): void {
	const scrollIntoView = Element.prototype.scrollIntoView;
	beforeEach(() => {
		vi.stubGlobal(
			'ResizeObserver',
			class {
				observe() {}
				unobserve() {}
				disconnect() {}
			},
		);
		Element.prototype.scrollIntoView = () => undefined;
	});
	afterEach(() => {
		vi.unstubAllGlobals();
		Element.prototype.scrollIntoView = scrollIntoView;
	});
}
