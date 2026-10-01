/** Color space helpers. All channels are floats in [0, 1] unless noted. */

export interface Rgb {
	r: number;
	g: number;
	b: number;
}

export interface Hsl {
	/** Hue in [0, 1). */
	h: number;
	s: number;
	l: number;
}

/** Parse `RRGGBB`, `#RRGGBB`, `AARRGGBB` (alpha ignored) or `RGB`. Returns `null` if invalid. */
export function parseHex(hex: string): Rgb | null {
	let h = hex.trim().replace(/^#/, '');
	if (h.length === 8) {
		h = h.slice(2);
	} else if (h.length === 3) {
		h = h
			.split('')
			.map((c) => c + c)
			.join('');
	}
	if (!/^[0-9a-fA-F]{6}$/.test(h)) {
		return null;
	}
	return {
		r: parseInt(h.slice(0, 2), 16) / 255,
		g: parseInt(h.slice(2, 4), 16) / 255,
		b: parseInt(h.slice(4, 6), 16) / 255,
	};
}

/** Alpha from an `AARRGGBB` string (1 when absent). */
export function parseArgbAlpha(hex: string): number {
	const h = hex.trim().replace(/^#/, '');
	if (h.length === 8) {
		return parseInt(h.slice(0, 2), 16) / 255;
	}
	return 1;
}

function clamp01(v: number): number {
	return v < 0 ? 0 : v > 1 ? 1 : v;
}

/** `RRGGBB` (upper case, no `#`). */
export function toHex(c: Rgb): string {
	const ch = (v: number) =>
		Math.round(clamp01(v) * 255)
			.toString(16)
			.padStart(2, '0');
	return `${ch(c.r)}${ch(c.g)}${ch(c.b)}`.toUpperCase();
}

export function rgbToHsl({ r, g, b }: Rgb): Hsl {
	const max = Math.max(r, g, b);
	const min = Math.min(r, g, b);
	const l = (max + min) / 2;
	if (max === min) {
		return { h: 0, s: 0, l };
	}
	const d = max - min;
	const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
	let h: number;
	if (max === r) {
		h = (g - b) / d + (g < b ? 6 : 0);
	} else if (max === g) {
		h = (b - r) / d + 2;
	} else {
		h = (r - g) / d + 4;
	}
	return { h: h / 6, s, l };
}

function hueToRgb(p: number, q: number, t: number): number {
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
}

export function hslToRgb({ h, s, l }: Hsl): Rgb {
	if (s === 0) {
		return { r: l, g: l, b: l };
	}
	const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
	const p = 2 * l - q;
	return {
		r: hueToRgb(p, q, h + 1 / 3),
		g: hueToRgb(p, q, h),
		b: hueToRgb(p, q, h - 1 / 3),
	};
}

/**
 * SpreadsheetML tint (`<color theme="4" tint="-0.25"/>`): negative values
 * darken, positive values lighten, applied to HSL luminance.
 */
export function applyTint(rgb: Rgb, tint: number): Rgb {
	if (!tint) {
		return rgb;
	}
	const hsl = rgbToHsl(rgb);
	if (tint < 0) {
		hsl.l = hsl.l * (1 + tint);
	} else {
		hsl.l = hsl.l * (1 - tint) + tint;
	}
	hsl.l = clamp01(hsl.l);
	return hslToRgb(hsl);
}

function toLinear(v: number): number {
	return v <= 0.04045 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4;
}

function fromLinear(v: number): number {
	return v <= 0.0031308 ? v * 12.92 : 1.055 * v ** (1 / 2.4) - 0.055;
}

export interface Rgba extends Rgb {
	a: number;
}

/**
 * Apply one DrawingML color transform. `val` is in thousandths of a percent
 * (100000 = 100%) for percentage transforms and 60000ths of a degree for hue.
 */
export function applyDrawingTransform(c: Rgba, name: string, val = 0): Rgba {
	const pct = val / 100000;
	switch (name) {
		case 'alpha':
			return { ...c, a: clamp01(pct) };
		case 'alphaMod':
			return { ...c, a: clamp01(c.a * pct) };
		case 'alphaOff':
			return { ...c, a: clamp01(c.a + pct) };
		case 'lumMod': {
			const hsl = rgbToHsl(c);
			hsl.l = clamp01(hsl.l * pct);
			return { ...hslToRgb(hsl), a: c.a };
		}
		case 'lumOff': {
			const hsl = rgbToHsl(c);
			hsl.l = clamp01(hsl.l + pct);
			return { ...hslToRgb(hsl), a: c.a };
		}
		case 'lum': {
			const hsl = rgbToHsl(c);
			hsl.l = clamp01(pct);
			return { ...hslToRgb(hsl), a: c.a };
		}
		case 'satMod': {
			const hsl = rgbToHsl(c);
			hsl.s = clamp01(hsl.s * pct);
			return { ...hslToRgb(hsl), a: c.a };
		}
		case 'satOff': {
			const hsl = rgbToHsl(c);
			hsl.s = clamp01(hsl.s + pct);
			return { ...hslToRgb(hsl), a: c.a };
		}
		case 'sat': {
			const hsl = rgbToHsl(c);
			hsl.s = clamp01(pct);
			return { ...hslToRgb(hsl), a: c.a };
		}
		case 'hueMod': {
			const hsl = rgbToHsl(c);
			hsl.h = (hsl.h * pct) % 1;
			return { ...hslToRgb(hsl), a: c.a };
		}
		case 'hueOff': {
			const hsl = rgbToHsl(c);
			hsl.h = (((hsl.h + val / 21600000) % 1) + 1) % 1;
			return { ...hslToRgb(hsl), a: c.a };
		}
		case 'hue': {
			const hsl = rgbToHsl(c);
			hsl.h = (((val / 21600000) % 1) + 1) % 1;
			return { ...hslToRgb(hsl), a: c.a };
		}
		case 'tint': {
			// Lighten toward white in linear RGB.
			const f = (v: number) => fromLinear(toLinear(v) * pct + (1 - pct));
			return { r: f(c.r), g: f(c.g), b: f(c.b), a: c.a };
		}
		case 'shade': {
			// Darken toward black in linear RGB.
			const f = (v: number) => fromLinear(toLinear(v) * pct);
			return { r: f(c.r), g: f(c.g), b: f(c.b), a: c.a };
		}
		case 'inv':
			return { r: 1 - c.r, g: 1 - c.g, b: 1 - c.b, a: c.a };
		case 'comp': {
			const hsl = rgbToHsl(c);
			hsl.h = (hsl.h + 0.5) % 1;
			return { ...hslToRgb(hsl), a: c.a };
		}
		case 'gray': {
			const y = 0.2126 * c.r + 0.7152 * c.g + 0.0722 * c.b;
			return { r: y, g: y, b: y, a: c.a };
		}
		case 'red':
			return { ...c, r: clamp01(pct) };
		case 'green':
			return { ...c, g: clamp01(pct) };
		case 'blue':
			return { ...c, b: clamp01(pct) };
		case 'redMod':
			return { ...c, r: clamp01(c.r * pct) };
		case 'greenMod':
			return { ...c, g: clamp01(c.g * pct) };
		case 'blueMod':
			return { ...c, b: clamp01(c.b * pct) };
		case 'redOff':
			return { ...c, r: clamp01(c.r + pct) };
		case 'greenOff':
			return { ...c, g: clamp01(c.g + pct) };
		case 'blueOff':
			return { ...c, b: clamp01(c.b + pct) };
		default:
			return c;
	}
}

/** CSS color string: `#RRGGBB` when opaque, otherwise `rgba(...)`. */
export function toCss(c: Rgba): string {
	if (c.a >= 0.999) {
		return `#${toHex(c)}`;
	}
	const ch = (v: number) => Math.round(clamp01(v) * 255);
	return `rgba(${ch(c.r)}, ${ch(c.g)}, ${ch(c.b)}, ${Math.round(c.a * 1000) / 1000})`;
}

/** Relative luminance (WCAG) of an sRGB color. */
export function relativeLuminance(c: Rgb): number {
	return 0.2126 * toLinear(c.r) + 0.7152 * toLinear(c.g) + 0.0722 * toLinear(c.b);
}
