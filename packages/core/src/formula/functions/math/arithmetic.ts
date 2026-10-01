import type { CellError } from '../../../types/cell';
import { ERRORS } from '../../../values/errors';
import type { FunctionDefinition } from '../../types';
import { finite, numberArg, optionalNumber } from './args';
import { binaryMath, unaryMath } from './define';
import { cleanNoise, NOISE, roundDirected, snapToInteger } from './rounding';

/**
 * `base ^ exponent` with Excel's rules: `0^0` is #NUM!, `0^-n` is #DIV/0!,
 * odd roots of negative numbers are real (`(-8)^(1/3)` is -2) and other
 * fractional powers of negative numbers are #NUM!.
 */
export function power(base: number, exponent: number): number | CellError {
	if (base === 0) {
		if (exponent === 0) {
			return ERRORS.NUM;
		}
		return exponent < 0 ? ERRORS.DIV0 : 0;
	}
	if (base < 0 && !Number.isInteger(exponent)) {
		const inverse = 1 / exponent;
		const root = Math.round(inverse);
		if (Math.abs(inverse - root) <= Math.abs(root) * NOISE * 4 && Math.abs(root % 2) === 1) {
			return finite(-Math.pow(-base, exponent));
		}
		return ERRORS.NUM;
	}
	return finite(Math.pow(base, exponent));
}

/** Quotients below this size get binary-noise cleanup in MOD (beyond it the exact remainder is kept). */
const MOD_NOISE_LIMIT = 2 ** 32;

/**
 * Excel `MOD`: the remainder has the sign of the divisor; `MOD(n, 0)` is
 * #DIV/0!. The remainder is computed exactly (`%` is IEEE fmod); results
 * that are only binary noise away from 0 or from the divisor become 0 and
 * the rest is cleaned to 15 digits (`MOD(1.1, 1)` is 0.1, `MOD(0.3, 0.1)` is 0).
 */
export function excelMod(n: number, d: number): number | CellError {
	if (d === 0) {
		return ERRORS.DIV0;
	}
	let r = n % d;
	if (r !== 0 && r < 0 !== d < 0) {
		r += d;
	}
	if (r === d) {
		return 0;
	}
	if (Math.abs(n / d) < MOD_NOISE_LIMIT) {
		const scale = Math.max(Math.abs(n), Math.abs(d));
		if (Math.abs(r) <= scale * NOISE || Math.abs(d - r) <= scale * NOISE) {
			return 0;
		}
		r = cleanNoise(r, scale);
	}
	return finite(r);
}

function logarithm(x: number, base: number): number | CellError {
	if (x <= 0 || base <= 0) {
		return ERRORS.NUM;
	}
	if (base === 1) {
		return ERRORS.DIV0;
	}
	if (base === 10) {
		return Math.log10(x);
	}
	if (base === 2) {
		return Math.log2(x);
	}
	return snapToInteger(Math.log(x) / Math.log(base));
}

const LOG: FunctionDefinition = {
	name: 'LOG',
	minArgs: 1,
	maxArgs: 2,
	params: ['value', 'value'],
	impl: (args, ctx) => {
		const x = numberArg(args[0], ctx);
		if (typeof x !== 'number') {
			return x;
		}
		const base = optionalNumber(args, 1, ctx, 10);
		if (typeof base !== 'number') {
			return base;
		}
		return finite(logarithm(x, base));
	},
};

const PI: FunctionDefinition = {
	name: 'PI',
	minArgs: 0,
	maxArgs: 0,
	params: [],
	impl: () => Math.PI,
};

export const ARITHMETIC_FUNCTIONS: FunctionDefinition[] = [
	unaryMath('ABS', Math.abs),
	unaryMath('SIGN', (x) => Math.sign(x)),
	unaryMath('SQRT', (x) => (x < 0 ? ERRORS.NUM : Math.sqrt(x))),
	unaryMath('SQRTPI', (x) => (x < 0 ? ERRORS.NUM : Math.sqrt(x * Math.PI))),
	unaryMath('EXP', Math.exp),
	unaryMath('LN', (x) => (x <= 0 ? ERRORS.NUM : Math.log(x))),
	unaryMath('LOG10', (x) => (x <= 0 ? ERRORS.NUM : Math.log10(x))),
	LOG,
	binaryMath('POWER', power),
	binaryMath('MOD', excelMod),
	binaryMath('QUOTIENT', (n, d) => (d === 0 ? ERRORS.DIV0 : roundDirected(n / d, 0, false))),
	unaryMath('DEGREES', (x) => (x * 180) / Math.PI),
	unaryMath('RADIANS', (x) => (x * Math.PI) / 180),
	PI,
];
