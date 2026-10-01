import { ERRORS, isCellError } from '../../../../values/errors';
import type { FunctionDefinition, Scalar } from '../../../types';
import { coerceArgs, columnOf, define, finish, numbersOf, valuesOf } from '../args';
import { sortedAscending } from './core';

/**
 * FREQUENCY(data, bins): counts per bin in the original bin order, each bin
 * holding the values above the next smaller bin and up to itself; one more
 * element counts the values above the largest bin.
 */
const FREQUENCY: FunctionDefinition = define({
	name: 'FREQUENCY',
	minArgs: 2,
	maxArgs: 2,
	params: ['any', 'any'],
	returnsArray: true,
	impl: (args, ctx) => {
		const data = numbersOf([args[0]], ctx);
		if (isCellError(data)) {
			return data;
		}
		const bins = numbersOf([args[1]], ctx);
		if (isCellError(bins)) {
			return bins;
		}
		const sorted = sortedAscending(data);
		// Stable sort of bin positions by value: the first of equal bins takes the count.
		const order = bins.map((_, i) => i).sort((a, b) => bins[a] - bins[b] || a - b);
		const counts = new Array<number>(bins.length + 1).fill(0);
		let i = 0;
		for (const position of order) {
			let c = 0;
			while (i < sorted.length && sorted[i] <= bins[position]) {
				c++;
				i++;
			}
			counts[position] = c;
		}
		counts[bins.length] = sorted.length - i;
		return columnOf(counts);
	},
});

/** PROB(x_range, prob_range, lower, [upper]): probability that a value lies in [lower, upper]. */
const PROB: FunctionDefinition = define({
	name: 'PROB',
	minArgs: 3,
	maxArgs: 4,
	params: ['any', 'any', 'value', 'value'],
	impl: (args, ctx) => {
		const xs = valuesOf(args[0], ctx);
		if (isCellError(xs)) {
			return xs;
		}
		const ps = valuesOf(args[1], ctx);
		if (isCellError(ps)) {
			return ps;
		}
		if (xs.length !== ps.length) {
			return ERRORS.NA;
		}
		const limits = coerceArgs(['n', { n: NaN }], args, ctx, 2);
		if (isCellError(limits)) {
			return limits;
		}
		const lower = limits[0];
		const upper = Number.isNaN(limits[1]) ? lower : limits[1];
		let total = 0;
		let inRange = 0;
		for (let i = 0; i < xs.length; i++) {
			const x: Scalar = xs[i];
			const p: Scalar = ps[i];
			if (isCellError(x)) {
				return x;
			}
			if (isCellError(p)) {
				return p;
			}
			if (typeof x !== 'number' || typeof p !== 'number') {
				continue;
			}
			if (p <= 0 || p > 1) {
				return ERRORS.NUM;
			}
			total += p;
			if (x >= lower && x <= upper) {
				inRange += p;
			}
		}
		if (Math.abs(total - 1) > 1e-7) {
			return ERRORS.NUM;
		}
		return finish(inRange);
	},
});

export const FREQUENCY_FUNCTIONS: FunctionDefinition[] = [FREQUENCY, PROB];
