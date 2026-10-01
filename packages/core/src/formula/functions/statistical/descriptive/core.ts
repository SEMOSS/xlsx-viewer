/** Numeric kernels shared by the descriptive statistics functions. */

/** Sum in input order, the way Excel accumulates. */
export function sum(xs: readonly number[]): number {
	let s = 0;
	for (const x of xs) {
		s += x;
	}
	return s;
}

export function mean(xs: readonly number[]): number {
	return sum(xs) / xs.length;
}

/** Sum of squared deviations from the mean (corrected two-pass algorithm). */
export function devsq(xs: readonly number[], m = mean(xs)): number {
	let ss = 0;
	let drift = 0;
	for (const x of xs) {
		const d = x - m;
		ss += d * d;
		drift += d;
	}
	return ss - (drift * drift) / xs.length;
}

/** Sample (n - 1) or population (n) variance; NaN when undefined. */
export function variance(xs: readonly number[], sample: boolean): number {
	const n = xs.length;
	if (n === 0 || (sample && n < 2)) {
		return NaN;
	}
	return devsq(xs) / (sample ? n - 1 : n);
}

/** Ascending copy. */
export function sortedAscending(xs: readonly number[]): number[] {
	// Typed-array sort is numeric (and much faster than a comparator sort).
	return Array.from(Float64Array.from(xs).sort());
}

/** Round to 15 significant digits, removing binary noise before floor / ceil. */
function snap(x: number): number {
	return Number.isFinite(x) && x !== 0 ? Number(x.toPrecision(15)) : x;
}

/** Floor that treats values within rounding noise of an integer as that integer. */
export function approxFloor(x: number): number {
	return Math.floor(snap(x));
}

/** Ceiling that treats values within rounding noise of an integer as that integer. */
export function approxCeil(x: number): number {
	return Math.ceil(snap(x));
}

/** Sum of ((x - m) / s)^power over the data. */
export function standardizedMoment(
	xs: readonly number[],
	m: number,
	s: number,
	power: 3 | 4,
): number {
	let acc = 0;
	for (const x of xs) {
		const z = (x - m) / s;
		acc += power === 3 ? z * z * z : z * z * z * z;
	}
	return acc;
}
