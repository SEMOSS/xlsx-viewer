import type { DrawingColor } from '../types/drawing';
import type { ColorSpec } from '../types/style';
import type { Theme, ThemeColorScheme } from '../types/theme';
import {
	applyDrawingTransform,
	applyTint,
	parseArgbAlpha,
	parseHex,
	toCss,
	toHex,
	type Rgba,
} from './color-math';
import { DEFAULT_INDEXED_COLORS, SYSTEM_BACKGROUND, SYSTEM_FOREGROUND } from './indexed-palette';

/**
 * SpreadsheetML theme index -> color scheme slot. Note the swapped first two
 * pairs: index 0 is lt1 ("Background 1") and index 1 is dk1 ("Text 1").
 */
const THEME_INDEX_SLOTS: ReadonlyArray<keyof ThemeColorScheme> = [
	'lt1',
	'dk1',
	'lt2',
	'dk2',
	'accent1',
	'accent2',
	'accent3',
	'accent4',
	'accent5',
	'accent6',
	'hlink',
	'folHlink',
];

/** Theme color (`RRGGBB`) for a SpreadsheetML theme index. */
export function themeColorByIndex(theme: Theme, index: number): string | undefined {
	const slot = THEME_INDEX_SLOTS[index];
	return slot ? (theme.colors[slot] as string) : undefined;
}

export interface ColorContext {
	theme: Theme;
	/** Custom indexed palette from styles.xml, if any. */
	indexedColors?: readonly string[];
}

/**
 * Resolve a SpreadsheetML color to `RRGGBB`. Returns `undefined` for
 * automatic colors (callers pick the context default, e.g. black text).
 */
export function resolveColorHex(
	color: ColorSpec | undefined,
	ctx: ColorContext,
): string | undefined {
	if (!color || color.auto) {
		return undefined;
	}
	let base: string | undefined;
	if (color.rgb) {
		base = color.rgb.length === 8 ? color.rgb.slice(2) : color.rgb;
	} else if (color.theme !== undefined) {
		base = themeColorByIndex(ctx.theme, color.theme);
	} else if (color.indexed !== undefined) {
		if (color.indexed === 64) {
			base = SYSTEM_FOREGROUND;
		} else if (color.indexed === 65) {
			base = SYSTEM_BACKGROUND;
		} else {
			const palette = ctx.indexedColors ?? DEFAULT_INDEXED_COLORS;
			const entry = palette[color.indexed];
			base = entry ? (entry.length === 8 ? entry.slice(2) : entry) : undefined;
		}
	}
	if (!base) {
		return undefined;
	}
	const rgb = parseHex(base);
	if (!rgb) {
		return undefined;
	}
	return toHex(color.tint ? applyTint(rgb, color.tint) : rgb);
}

/** Resolve a SpreadsheetML color to a CSS color, or `fallback` for automatic colors. */
export function resolveColorCss(
	color: ColorSpec | undefined,
	ctx: ColorContext,
	fallback?: string,
): string | undefined {
	const hex = resolveColorHex(color, ctx);
	if (!hex) {
		return fallback;
	}
	// ARGB alpha is ignored by Excel for cell colors except fully transparent '00' which it treats as opaque.
	return `#${hex}`;
}

/** DrawingML scheme color name -> theme slot, using Excel's default color map. */
const SCHEME_SLOTS: Record<string, keyof ThemeColorScheme> = {
	bg1: 'lt1',
	tx1: 'dk1',
	bg2: 'lt2',
	tx2: 'dk2',
	lt1: 'lt1',
	dk1: 'dk1',
	lt2: 'lt2',
	dk2: 'dk2',
	accent1: 'accent1',
	accent2: 'accent2',
	accent3: 'accent3',
	accent4: 'accent4',
	accent5: 'accent5',
	accent6: 'accent6',
	hlink: 'hlink',
	folHlink: 'folHlink',
};

/** A subset of DrawingML preset color names whose CSS name differs or that CSS lacks. */
const PRESET_COLORS: Record<string, string> = {
	black: '000000',
	white: 'FFFFFF',
	red: 'FF0000',
	green: '008000',
	blue: '0000FF',
	yellow: 'FFFF00',
	cyan: '00FFFF',
	magenta: 'FF00FF',
	gray: '808080',
	grey: '808080',
	orange: 'FFA500',
	purple: '800080',
	brown: 'A52A2A',
	pink: 'FFC0CB',
	navy: '000080',
	teal: '008080',
	olive: '808000',
	maroon: '800000',
	silver: 'C0C0C0',
	lime: '00FF00',
	gold: 'FFD700',
	dkBlue: '00008B',
	dkRed: '8B0000',
	dkGreen: '006400',
	dkGray: 'A9A9A9',
	ltGray: 'D3D3D3',
	ltBlue: 'ADD8E6',
	ltGreen: '90EE90',
	ltYellow: 'FFFFE0',
};

/**
 * Resolve a DrawingML color (with transforms) to RGBA. `placeholder` supplies
 * the color for `phClr` (style-matrix placeholders).
 */
export function resolveDrawingColor(
	color: DrawingColor | undefined,
	theme: Theme,
	placeholder?: string,
): Rgba | undefined {
	if (!color) {
		return undefined;
	}
	let base: string | undefined;
	switch (color.kind) {
		case 'srgb':
			base = color.value;
			break;
		case 'scheme': {
			if (color.value === 'phClr') {
				base = placeholder;
			} else {
				const slot = SCHEME_SLOTS[color.value];
				base = slot ? (theme.colors[slot] as string) : undefined;
			}
			break;
		}
		case 'sys':
			base = /^[0-9a-fA-F]{6}$/.test(color.value)
				? color.value
				: color.value === 'window'
					? 'FFFFFF'
					: '000000';
			break;
		case 'preset':
			base = PRESET_COLORS[color.value] ?? PRESET_COLORS[color.value.toLowerCase()];
			break;
		case 'hsl': {
			const [h, s, l] = color.value.split(',').map(Number);
			base = toHexFromHsl(h, s, l);
			break;
		}
		case 'scrgb': {
			const [r, g, b] = color.value.split(',').map((v) => Number(v) / 100000);
			base = toHex({ r, g, b });
			break;
		}
	}
	if (!base) {
		return undefined;
	}
	const rgb = parseHex(base);
	if (!rgb) {
		return undefined;
	}
	let c: Rgba = { ...rgb, a: base.length === 8 ? parseArgbAlpha(base) : 1 };
	for (const t of color.transforms ?? []) {
		c = applyDrawingTransform(c, t.name, t.val);
	}
	return c;
}

function toHexFromHsl(hue60k: number, satPct: number, lumPct: number): string {
	const h = (hue60k / 21600000) % 1;
	const s = satPct / 100000;
	const l = lumPct / 100000;
	const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
	const p = 2 * l - q;
	const conv = (t: number) => {
		let tt = t;
		if (tt < 0) {
			tt += 1;
		}
		if (tt > 1) {
			tt -= 1;
		}
		if (tt < 1 / 6) {
			return p + (q - p) * 6 * tt;
		}
		if (tt < 1 / 2) {
			return q;
		}
		if (tt < 2 / 3) {
			return p + (q - p) * (2 / 3 - tt) * 6;
		}
		return p;
	};
	return toHex(
		s === 0 ? { r: l, g: l, b: l } : { r: conv(h + 1 / 3), g: conv(h), b: conv(h - 1 / 3) },
	);
}

/** Resolve a DrawingML color straight to a CSS string. */
export function resolveDrawingColorCss(
	color: DrawingColor | undefined,
	theme: Theme,
	placeholder?: string,
): string | undefined {
	const c = resolveDrawingColor(color, theme, placeholder);
	return c ? toCss(c) : undefined;
}
