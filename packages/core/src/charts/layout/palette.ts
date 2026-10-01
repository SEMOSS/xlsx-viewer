/**
 * Excel's automatic chart formatting: series / point colors for the built-in
 * chart styles and the automatic marker cycle.
 */
import type { ChartMarker } from '../../types/chart';
import type { DrawingColor } from '../../types/drawing';

/**
 * Luminance variations applied to the six accents on each pass through the
 * palette (Office "colorful palette": accents, darker 60%, lighter 80% + 20%, ...).
 */
const VARIATIONS: ReadonlyArray<ReadonlyArray<{ name: string; val: number }>> = [
	[],
	[{ name: 'lumMod', val: 60000 }],
	[
		{ name: 'lumMod', val: 80000 },
		{ name: 'lumOff', val: 20000 },
	],
	[{ name: 'lumMod', val: 80000 }],
	[
		{ name: 'lumMod', val: 60000 },
		{ name: 'lumOff', val: 40000 },
	],
	[{ name: 'lumMod', val: 50000 }],
	[
		{ name: 'lumMod', val: 70000 },
		{ name: 'lumOff', val: 30000 },
	],
	[{ name: 'lumMod', val: 70000 }],
	[
		{ name: 'lumMod', val: 50000 },
		{ name: 'lumOff', val: 50000 },
	],
];

/** Gray ramp used by the monochrome-black styles (1, 9, 17, ...). */
const GRAYS: ReadonlyArray<ReadonlyArray<{ name: string; val: number }>> = [
	[
		{ name: 'lumMod', val: 65000 },
		{ name: 'lumOff', val: 35000 },
	],
	[
		{ name: 'lumMod', val: 50000 },
		{ name: 'lumOff', val: 50000 },
	],
	[
		{ name: 'lumMod', val: 85000 },
		{ name: 'lumOff', val: 15000 },
	],
	[
		{ name: 'lumMod', val: 35000 },
		{ name: 'lumOff', val: 65000 },
	],
	[
		{ name: 'lumMod', val: 75000 },
		{ name: 'lumOff', val: 25000 },
	],
	[
		{ name: 'lumMod', val: 25000 },
		{ name: 'lumOff', val: 75000 },
	],
];

/** Normalise `c:style` / `c14:style` (101-148) to the 1-48 gallery index. */
export function chartStyleIndex(style: number | undefined): number {
	if (style === undefined || !Number.isFinite(style)) {
		return 2;
	}
	const s = style > 100 ? style - 100 : style;
	return s >= 1 && s <= 48 ? Math.round(s) : 2;
}

/**
 * Automatic color for series (or point, with varyColors) `index` of `count`.
 * Style 2 (the default) cycles accent1..accent6 with luminance variations;
 * monochrome styles shade one accent from dark to light.
 */
export function autoChartColor(index: number, count: number, style?: number): DrawingColor {
	const column = (chartStyleIndex(style) - 1) % 8;
	const i = Math.max(0, index);
	if (column === 0) {
		return { kind: 'scheme', value: 'tx1', transforms: [...GRAYS[i % GRAYS.length]] };
	}
	if (column === 1) {
		const pass = Math.floor(i / 6) % VARIATIONS.length;
		const color: DrawingColor = { kind: 'scheme', value: `accent${(i % 6) + 1}` };
		if (VARIATIONS[pass].length) {
			color.transforms = [...VARIATIONS[pass]];
		}
		return color;
	}
	return monochromeColor(`accent${column - 1}`, i, count);
}

/** Accent `accent` with the luminance variation of pass `pass` through a colorful palette. */
export function accentVariation(accent: string, pass: number): DrawingColor {
	const variation = VARIATIONS[Math.max(0, pass) % VARIATIONS.length];
	const color: DrawingColor = { kind: 'scheme', value: accent };
	if (variation.length) {
		color.transforms = [...variation];
	}
	return color;
}

/**
 * Monochromatic palettes: shades (darker) through tints (lighter) of one
 * scheme color, spread evenly over `count` series.
 */
export function monochromeColor(scheme: string, index: number, count: number): DrawingColor {
	const n = Math.max(1, count);
	if (n === 1) {
		return { kind: 'scheme', value: scheme };
	}
	const t = (Math.max(0, index) % n) / (n - 1);
	const shade = t < 0.5;
	const amount = Math.round((shade ? 0.5 + t : 1.5 - t) * 100000);
	return {
		kind: 'scheme',
		value: scheme,
		transforms: [{ name: shade ? 'shade' : 'tint', val: Math.min(100000, amount) }],
	};
}

/** Excel's automatic marker sequence for line, scatter and radar series. */
const AUTO_MARKERS: ReadonlyArray<NonNullable<ChartMarker['symbol']>> = [
	'diamond',
	'square',
	'triangle',
	'x',
	'star',
	'circle',
	'plus',
];

export function autoMarkerSymbol(index: number): NonNullable<ChartMarker['symbol']> {
	return AUTO_MARKERS[Math.max(0, index) % AUTO_MARKERS.length];
}
