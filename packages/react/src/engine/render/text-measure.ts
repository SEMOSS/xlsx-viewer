/**
 * Text width measurement with a shared offscreen canvas and an LRU-ish
 * cache. Falls back to an average-width estimate where canvas is not
 * available (tests, SSR).
 */
let ctx: CanvasRenderingContext2D | null | undefined;
const cache = new Map<string, number>();
const MAX_CACHE = 20000;

function context(): CanvasRenderingContext2D | null {
	if (ctx !== undefined) {
		return ctx;
	}
	try {
		if (typeof document === 'undefined') {
			ctx = null;
			return ctx;
		}
		const canvas = document.createElement('canvas');
		ctx = typeof canvas.getContext === 'function' ? canvas.getContext('2d') : null;
	} catch {
		ctx = null;
	}
	return ctx;
}

function estimate(text: string, font: string): number {
	const m = /(\d+(?:\.\d+)?)px/.exec(font);
	const size = m ? Number(m[1]) : 14;
	const bold = /\b(700|bold)\b/.test(font) ? 1.06 : 1;
	let units = 0;
	for (const ch of text) {
		if (/[il.,:;'|!]/.test(ch)) {
			units += 0.3;
		} else if (/[mwMW@]/.test(ch)) {
			units += 0.85;
		} else if (/[A-Z0-9]/.test(ch)) {
			units += 0.62;
		} else if (ch.charCodeAt(0) > 0x2e80) {
			units += 1;
		} else {
			units += 0.52;
		}
	}
	return units * size * bold;
}

/** Width of `text` in CSS pixels for a canvas font string. */
export function measureText(text: string, font: string): number {
	if (text === '') {
		return 0;
	}
	const key = `${font}\u0000${text}`;
	const hit = cache.get(key);
	if (hit !== undefined) {
		return hit;
	}
	const c = context();
	let width: number;
	if (c) {
		c.font = font;
		width = c.measureText(text).width;
	} else {
		width = estimate(text, font);
	}
	if (cache.size > MAX_CACHE) {
		cache.clear();
	}
	cache.set(key, width);
	return width;
}

/** Widest line of a multi-line text. */
export function measureLines(text: string, font: string): number {
	let max = 0;
	for (const line of text.split('\n')) {
		max = Math.max(max, measureText(line, font));
	}
	return max;
}

/** Override the measurer (tests). */
export function setMeasureContextForTests(next: CanvasRenderingContext2D | null): void {
	ctx = next;
	cache.clear();
}
