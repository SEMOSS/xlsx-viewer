import { toHex, type Rgb } from '../color/color-math';

/** Linear interpolation between two colors in RGB space (`t` clamped to [0, 1]). */
export function lerpRgb(a: Rgb, b: Rgb, t: number): Rgb {
	const f = t <= 0 ? 0 : t >= 1 ? 1 : t;
	return {
		r: a.r + (b.r - a.r) * f,
		g: a.g + (b.g - a.g) * f,
		b: a.b + (b.b - a.b) * f,
	};
}

function segmentColor(value: number, t0: number, c0: Rgb, t1: number, c1: Rgb): Rgb {
	if (value <= t0) {
		return c0;
	}
	if (value >= t1) {
		return c1;
	}
	return lerpRgb(c0, c1, (value - t0) / (t1 - t0));
}

/**
 * Color of a 2- or 3-point color scale for `value`. Values at or below the
 * first threshold take the first color, values at or above the last take the
 * last color; in between, the two colors of the enclosing segment are mixed
 * linearly in RGB. When all thresholds coincide (every value equal) the
 * first color is used.
 */
export function interpolateColorScale(
	value: number,
	thresholds: readonly number[],
	colors: readonly Rgb[],
): Rgb {
	const n = Math.min(thresholds.length, colors.length);
	if (n === 0) {
		return { r: 1, g: 1, b: 1 };
	}
	if (n === 1) {
		return colors[0];
	}
	let i = 0;
	while (i < n - 2 && value > thresholds[i + 1]) {
		i++;
	}
	return segmentColor(value, thresholds[i], colors[i], thresholds[i + 1], colors[i + 1]);
}

/** `#RRGGBB` for a color with float channels. */
export function rgbToCss(c: Rgb): string {
	return `#${toHex(c)}`;
}
