import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import { forEachScalar, toNumber } from '../../coerce';
import type { FunctionArg, FunctionContext, FunctionDefinition } from '../../types';
import { isMissingArg } from './args';
import { binaryMath, unaryMath, XLFN } from './define';
import { approxFloor } from './rounding';

/** Integers at or above 2^53 are rejected by GCD / LCM. */
const MAX_INTEGER = 2 ** 53;

/** n! for a non-negative integer (Infinity on overflow). */
export function factorial(n: number): number {
	let r = 1;
	for (let i = 2; i <= n; i++) {
		r *= i;
		if (!Number.isFinite(r)) {
			return Infinity;
		}
	}
	return r;
}

/** Binomial coefficient C(n, k) for integers; 0 when k > n. */
export function binomial(n: number, k: number): number {
	if (k < 0 || k > n) {
		return 0;
	}
	const kk = Math.min(k, n - k);
	let r = 1;
	for (let i = 1; i <= kk; i++) {
		r = (r * (n - kk + i)) / i;
		if (!Number.isFinite(r)) {
			return Infinity;
		}
	}
	return r < MAX_INTEGER ? Math.round(r) : r;
}

function fact(x: number): number | CellError {
	if (x < 0) {
		return ERRORS.NUM;
	}
	return factorial(approxFloor(x));
}

function factDouble(x: number): number | CellError {
	if (x < 0) {
		return ERRORS.NUM;
	}
	let r = 1;
	for (let i = approxFloor(x); i > 1; i -= 2) {
		r *= i;
		if (!Number.isFinite(r)) {
			return ERRORS.NUM;
		}
	}
	return r;
}

function combin(n: number, k: number): number | CellError {
	const nn = Math.trunc(n);
	const kk = Math.trunc(k);
	if (nn < 0 || kk < 0 || nn < kk) {
		return ERRORS.NUM;
	}
	return binomial(nn, kk);
}

function combinA(n: number, k: number): number | CellError {
	const nn = Math.trunc(n);
	const kk = Math.trunc(k);
	if (nn < 0 || kk < 0) {
		return ERRORS.NUM;
	}
	if (nn === 0) {
		return kk === 0 ? 1 : 0;
	}
	return binomial(nn + kk - 1, kk);
}

/**
 * Collect the integer arguments of GCD / LCM / MULTINOMIAL. Numbers in
 * ranges count, blank cells are skipped, text or booleans inside ranges are
 * #VALUE!; direct arguments are coerced. Values are truncated and must be
 * non-negative.
 */
function collectIntegers(args: readonly FunctionArg[], ctx: FunctionContext): number[] | CellError {
	const out: number[] = [];
	let error: CellError | undefined;
	const accept = (n: number) => {
		if (n < 0 || n >= MAX_INTEGER) {
			error = ERRORS.NUM;
		} else {
			out.push(Math.trunc(n));
		}
	};
	for (const arg of args) {
		if (isMissingArg(arg)) {
			accept(0);
			continue;
		}
		forEachScalar(arg, ctx, (v, fromRef) => {
			if (error) {
				return;
			}
			if (isCellError(v)) {
				error = v;
			} else if (typeof v === 'number') {
				accept(v);
			} else if (v === null) {
				// Blank cells are ignored.
			} else if (fromRef) {
				error = ERRORS.VALUE;
			} else {
				const n = toNumber(v, ctx.date1904);
				if (typeof n === 'number') {
					accept(n);
				} else {
					error = n;
				}
			}
		});
		if (error) {
			return error;
		}
	}
	return out;
}

function gcd2(a: number, b: number): number {
	let x = a;
	let y = b;
	while (y !== 0) {
		const t = x % y;
		x = y;
		y = t;
	}
	return x;
}

function variadicIntegers(
	name: string,
	reduce: (values: number[]) => number | CellError,
): FunctionDefinition {
	return {
		name,
		minArgs: 1,
		maxArgs: 255,
		params: ['any'],
		impl: (args, ctx) => {
			const values = collectIntegers(args, ctx);
			return isCellError(values) ? values : reduce(values);
		},
	};
}

function lcm(values: number[]): number | CellError {
	let r = 1;
	for (const v of values) {
		if (v === 0) {
			return 0;
		}
		r = (r / gcd2(r, v)) * v;
		if (r >= MAX_INTEGER) {
			return ERRORS.NUM;
		}
	}
	return r;
}

function multinomial(values: number[]): number | CellError {
	let total = 0;
	let r = 1;
	for (const v of values) {
		total += v;
		r *= binomial(total, v);
		if (!Number.isFinite(r)) {
			return ERRORS.NUM;
		}
	}
	return r;
}

export const COMBINATORICS_FUNCTIONS: FunctionDefinition[] = [
	unaryMath('FACT', fact),
	unaryMath('FACTDOUBLE', factDouble),
	binaryMath('COMBIN', combin),
	binaryMath('COMBINA', combinA, XLFN),
	variadicIntegers('GCD', (values) => values.reduce(gcd2, 0)),
	variadicIntegers('LCM', lcm),
	variadicIntegers('MULTINOMIAL', multinomial),
];
