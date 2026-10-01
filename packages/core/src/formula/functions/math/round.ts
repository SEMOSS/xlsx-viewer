import type { CellError } from '../../../types/cell';
import { ERRORS } from '../../../values/errors';
import type { FunctionDefinition } from '../../types';
import { finite, numberArg, optionalNumber } from './args';
import { binaryMath, unaryMath, XLFN } from './define';
import { approxCeil, approxFloor, multipleOf, roundDirected, roundHalfAway } from './rounding';

type Result = number | CellError;

/**
 * CEILING (Excel 2010+ semantics): toward +infinity for a positive
 * significance, away from zero when both arguments are negative.
 */
function ceiling(n: number, s: number): Result {
	if (s === 0 || n === 0) {
		return 0;
	}
	if (n > 0 && s < 0) {
		return ERRORS.NUM;
	}
	return multipleOf(approxCeil(n / s), s);
}

/**
 * FLOOR: toward zero for positive numbers; for negative numbers away from
 * zero with a positive significance, toward zero with a negative one.
 */
function floor(n: number, s: number): Result {
	if (n === 0) {
		return 0;
	}
	if (s === 0) {
		return ERRORS.DIV0;
	}
	if (n > 0 && s < 0) {
		return ERRORS.NUM;
	}
	return multipleOf(approxFloor(n / s), s);
}

/**
 * CEILING.MATH / FLOOR.MATH: the sign of the significance is ignored;
 * `mode` (non-zero) reverses the direction for negative numbers.
 */
function mathRound(n: number, significance: number, mode: number, up: boolean): Result {
	const s = Math.abs(significance);
	if (s === 0 || n === 0) {
		return 0;
	}
	const q = n / s;
	if (n < 0 && mode !== 0) {
		// Ceiling rounds away from zero, floor toward zero.
		const steps = up ? -approxCeil(-q) : -approxFloor(-q);
		return multipleOf(steps, s);
	}
	return multipleOf(up ? approxCeil(q) : approxFloor(q), s);
}

function mathVariant(
	name: string,
	up: boolean,
	withMode: boolean,
	prefixed: boolean,
): FunctionDefinition {
	return {
		name,
		minArgs: 1,
		maxArgs: withMode ? 3 : 2,
		params: ['value', 'value', 'value'],
		...(prefixed ? { filePrefix: XLFN } : {}),
		impl: (args, ctx) => {
			const n = numberArg(args[0], ctx);
			if (typeof n !== 'number') {
				return n;
			}
			const s = optionalNumber(args, 1, ctx, 1);
			if (typeof s !== 'number') {
				return s;
			}
			const mode = withMode ? optionalNumber(args, 2, ctx, 0) : 0;
			if (typeof mode !== 'number') {
				return mode;
			}
			return finite(mathRound(n, s, mode, up));
		},
	};
}

function roundingFunction(
	name: string,
	round: (x: number, digits: number) => number,
): FunctionDefinition {
	return binaryMath(name, (x, digits) => round(x, digits));
}

const TRUNC: FunctionDefinition = {
	name: 'TRUNC',
	minArgs: 1,
	maxArgs: 2,
	params: ['value', 'value'],
	impl: (args, ctx) => {
		const x = numberArg(args[0], ctx);
		if (typeof x !== 'number') {
			return x;
		}
		const digits = optionalNumber(args, 1, ctx, 0);
		return typeof digits === 'number' ? finite(roundDirected(x, digits, false)) : digits;
	},
};

/** EVEN / ODD: round away from zero to the nearest even / odd integer. */
function parity(odd: boolean) {
	return (x: number): number => {
		let n = approxCeil(Math.abs(x));
		if ((n % 2 === 1) !== odd) {
			n += 1;
		}
		return x < 0 ? -n : n;
	};
}

function mround(n: number, m: number): Result {
	if (m === 0 || n === 0) {
		return 0;
	}
	if ((n > 0 && m < 0) || (n < 0 && m > 0)) {
		return ERRORS.NUM;
	}
	return multipleOf(roundHalfAway(n / m, 0), m);
}

export const ROUND_FUNCTIONS: FunctionDefinition[] = [
	roundingFunction('ROUND', roundHalfAway),
	roundingFunction('ROUNDUP', (x, d) => roundDirected(x, d, true)),
	roundingFunction('ROUNDDOWN', (x, d) => roundDirected(x, d, false)),
	TRUNC,
	unaryMath('INT', approxFloor),
	unaryMath('EVEN', parity(false)),
	unaryMath('ODD', parity(true)),
	binaryMath('MROUND', mround),
	binaryMath('CEILING', ceiling),
	binaryMath('FLOOR', floor),
	mathVariant('CEILING.MATH', true, true, true),
	mathVariant('CEILING.PRECISE', true, false, true),
	mathVariant('ISO.CEILING', true, false, false),
	mathVariant('FLOOR.MATH', false, true, true),
	mathVariant('FLOOR.PRECISE', false, false, true),
];
