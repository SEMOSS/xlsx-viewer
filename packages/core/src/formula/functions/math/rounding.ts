/**
 * Precision helpers that reproduce Excel's rounding behaviour on IEEE-754
 * doubles. Excel reasons about numbers as 15 significant decimal digits, so
 * `ROUND(2.675, 2)` is 2.68 even though the double closest to 2.675 is
 * slightly below it, and `INT((0.1 + 0.7) * 10)` is 8.
 */

/** Relative tolerance treated as binary noise (a few units in the last place). */
export const NOISE = 4 * Number.EPSILON;

/** Tighter tolerance for {@link cleanNoise}, which must not swallow genuine 16th / 17th digits. */
const CLEAN_NOISE = 2 * Number.EPSILON;

/** Values at or above this magnitude are always integers in binary64. */
const INTEGRAL_LIMIT = 2 ** 52;

/** Decimal exponent of a finite non-zero number (`1234` -> 3, `0.05` -> -2). */
export function decimalExponent(x: number): number {
	const text = Math.abs(x).toExponential();
	return Number(text.slice(text.indexOf('e') + 1));
}

/**
 * Multiply by `10^places` exactly, by shifting the exponent of the
 * shortest decimal representation instead of multiplying in binary.
 */
export function shiftDecimal(x: number, places: number): number {
	if (x === 0 || places === 0 || !Number.isFinite(x)) {
		return x;
	}
	const text = x.toExponential();
	const e = text.indexOf('e');
	return Number(`${text.slice(0, e)}e${Number(text.slice(e + 1)) + places}`);
}

/** `q` snapped to the nearest integer when it is within binary noise of it. */
export function snapToInteger(q: number): number {
	if (!Number.isFinite(q)) {
		return q;
	}
	const r = Math.round(q);
	return r !== q && r !== 0 && Math.abs(q - r) <= Math.abs(r) * NOISE ? r : q;
}

/** Floor that ignores binary noise (`approxFloor(2.9999999999999996)` is 3). */
export function approxFloor(q: number): number {
	return Math.floor(snapToInteger(q));
}

/** Ceiling that ignores binary noise (`approxCeil(3.0000000000000004)` is 3). */
export function approxCeil(q: number): number {
	return Math.ceil(snapToInteger(q));
}

/**
 * Remove binary noise from a computed result: its 15-significant-digit form
 * when that lies within about two ULPs (`0.1 + 0.2` becomes 0.3). `scale`
 * is the magnitude the noise is relative to (the inputs of a subtraction).
 */
export function cleanNoise(x: number, scale = Math.abs(x)): number {
	if (x === 0 || !Number.isFinite(x)) {
		return x;
	}
	const c = Number(x.toPrecision(15));
	return Math.abs(c - x) <= scale * CLEAN_NOISE ? c : x;
}

function clampDigits(digits: number): number {
	return Math.max(-340, Math.min(340, Math.trunc(digits)));
}

function signed(x: number, magnitude: number): number {
	if (magnitude === 0) {
		return 0;
	}
	return x < 0 ? -magnitude : magnitude;
}

/**
 * Excel `ROUND`: half away from zero, applied to the number's
 * 15-significant-digit decimal form (`ROUND(2.675, 2)` is 2.68,
 * `ROUND(-0.5, 0)` is -1). Negative `digits` round left of the point.
 */
export function roundHalfAway(x: number, digits: number): number {
	if (!Number.isFinite(x) || x === 0) {
		return x;
	}
	const d = clampDigits(digits);
	const ax = Math.abs(x);
	const exponent = decimalExponent(ax);
	if (exponent + d < -1) {
		return 0;
	}
	const base = exponent + d < 15 ? Number(ax.toPrecision(15)) : ax;
	const scaled = shiftDecimal(base, d);
	if (!(scaled < INTEGRAL_LIMIT)) {
		return x;
	}
	const floor = Math.floor(scaled);
	const rounded = scaled - floor >= 0.5 ? floor + 1 : floor;
	return signed(x, shiftDecimal(rounded, -d));
}

/**
 * Excel `ROUNDUP` (`away` = true, away from zero) and `ROUNDDOWN` / `TRUNC`
 * (toward zero). Values within binary noise of a boundary are not pushed
 * across it (`ROUNDUP(0.1 + 0.2, 1)` is 0.3).
 */
export function roundDirected(x: number, digits: number, away: boolean): number {
	if (!Number.isFinite(x) || x === 0) {
		return x;
	}
	const d = clampDigits(digits);
	const scaled = shiftDecimal(Math.abs(x), d);
	if (!(scaled < INTEGRAL_LIMIT)) {
		return x;
	}
	const snapped = snapToInteger(scaled);
	const rounded = away ? Math.ceil(snapped) : Math.floor(snapped);
	return signed(x, shiftDecimal(rounded, -d));
}

/** Multiply an integer count of steps by a step size, removing binary noise from the product. */
export function multipleOf(steps: number, step: number): number {
	const product = steps * step;
	return product === 0 ? 0 : cleanNoise(product);
}
