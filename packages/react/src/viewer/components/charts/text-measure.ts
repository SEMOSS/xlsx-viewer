/**
 * Canvas text measurement for chart layout, so labels and legends are sized
 * with the fonts the browser actually renders. Falls back to the core's
 * deterministic Calibri metrics where no canvas is available (SSR, jsdom).
 */
import type { ChartTextMeasurer } from 'xlsx-viewer-core';

let shared: ChartTextMeasurer | null | undefined;

function create(): ChartTextMeasurer | null {
	if (
		typeof document === 'undefined' ||
		typeof navigator === 'undefined' ||
		/jsdom/i.test(navigator.userAgent)
	) {
		return null;
	}
	let ctx: CanvasRenderingContext2D | null = null;
	try {
		ctx = document.createElement('canvas').getContext('2d');
	} catch {
		return null;
	}
	if (!ctx || typeof ctx.measureText !== 'function') {
		return null;
	}
	const context = ctx;
	const cache = new Map<string, number>();
	return (text, font) => {
		const css = `${font.italic ? 'italic ' : ''}${font.bold ? 'bold ' : ''}${font.size}px "${font.family}", Calibri, Carlito, "Segoe UI", Arial, sans-serif`;
		const key = `${css}\u0000${text}`;
		let width = cache.get(key);
		if (width === undefined) {
			context.font = css;
			width = context.measureText(text).width;
			if (cache.size > 5000) {
				cache.clear();
			}
			cache.set(key, width);
		}
		return width;
	};
}

/** A shared canvas measurer, or `undefined` when the environment has no canvas. */
export function canvasTextMeasurer(): ChartTextMeasurer | undefined {
	if (shared === undefined) {
		shared = create();
	}
	return shared ?? undefined;
}
