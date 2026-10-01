import { ERRORS, isCellError } from '../../../../values/errors';
import type { FunctionDefinition } from '../../../types';
import { aggregateFunction, coerceArgs, define, numbersOf } from '../args';
import { A_OPTIONS } from './averages';
import { approxCeil, sortedAscending } from './core';

function max(xs: number[]): number {
	let m = -Infinity;
	for (const x of xs) {
		if (x > m) {
			m = x;
		}
	}
	return xs.length === 0 ? 0 : m;
}

function min(xs: number[]): number {
	let m = Infinity;
	for (const x of xs) {
		if (x < m) {
			m = x;
		}
	}
	return xs.length === 0 ? 0 : m;
}

/** LARGE / SMALL: the k-th largest or smallest number (k rounded up, as Excel does). */
function kth(name: string, largest: boolean): FunctionDefinition {
	return define({
		name,
		minArgs: 2,
		maxArgs: 2,
		params: ['any', 'value'],
		impl: (args, ctx) => {
			const xs = numbersOf([args[0]], ctx);
			if (isCellError(xs)) {
				return xs;
			}
			const v = coerceArgs(['n'], args, ctx, 1);
			if (isCellError(v)) {
				return v;
			}
			const k = approxCeil(v[0]);
			if (!(k >= 1) || k > xs.length) {
				return ERRORS.NUM;
			}
			const s = sortedAscending(xs);
			return largest ? s[s.length - k] : s[k - 1];
		},
	});
}

export const EXTREME_FUNCTIONS: FunctionDefinition[] = [
	aggregateFunction('MAX', max),
	aggregateFunction('MAXA', max, A_OPTIONS),
	aggregateFunction('MIN', min),
	aggregateFunction('MINA', min, A_OPTIONS),
	kth('LARGE', true),
	kth('SMALL', false),
];
