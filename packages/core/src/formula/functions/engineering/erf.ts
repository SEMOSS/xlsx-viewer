/**
 * Error function: ERF, ERF.PRECISE, ERFC and ERFC.PRECISE.
 *
 * erf uses the all-positive series
 * `erf(x) = 2/sqrt(pi) * exp(-x^2) * sum(2^n x^(2n+1) / (2n+1)!!)` (no
 * cancellation, ~1e-16 relative); erfc uses the continued fraction
 * `sqrt(pi) e^(x^2) erfc(x) = 1 / (x + (1/2) / (x + 1 / (x + (3/2) / (x + ...))))`
 * evaluated with Lentz's method, accurate to ~1e-15 for x >= 0.5.
 */
import { isCellError } from '../../../values/errors';
import type { FunctionArg, FunctionContext, FunctionDefinition } from '../../types';
import { isOmitted } from '../datetime/args';
import { engNum } from './args';

const TWO_OVER_SQRT_PI = 2 / Math.sqrt(Math.PI);
const SQRT_PI = Math.sqrt(Math.PI);

function erfSeries(x: number): number {
	const x2 = 2 * x * x;
	let term = x;
	let sum = x;
	for (let n = 1; n < 500; n++) {
		term *= x2 / (2 * n + 1);
		sum += term;
		if (Math.abs(term) < 1e-17 * Math.abs(sum)) {
			break;
		}
	}
	return TWO_OVER_SQRT_PI * Math.exp(-x * x) * sum;
}

/** erfc for x >= 0.5 by continued fraction (modified Lentz). */
function erfcFraction(x: number): number {
	if (x > 27) {
		return 0;
	}
	const tiny = 1e-300;
	let f = x;
	let c = x;
	let d = 0;
	for (let k = 1; k < 5000; k++) {
		const a = k / 2;
		d = x + a * d;
		c = x + a / c;
		d = d === 0 ? 1 / tiny : 1 / d;
		if (c === 0) {
			c = tiny;
		}
		const delta = c * d;
		f *= delta;
		if (Math.abs(delta - 1) < 1e-16) {
			break;
		}
	}
	return Math.exp(-x * x) / (SQRT_PI * f);
}

/** The error function. */
export function erf(x: number): number {
	if (x < 0) {
		return -erf(-x);
	}
	return x < 2 ? erfSeries(x) : 1 - erfcFraction(x);
}

/** The complementary error function 1 - erf(x), accurate in the tails. */
export function erfc(x: number): number {
	if (x < 0) {
		return 2 - erfc(-x);
	}
	return x < 0.5 ? 1 - erfSeries(x) : erfcFraction(x);
}

/** erf(b) - erf(a), computed from erfc when both limits share a sign to avoid cancellation. */
function erfBetween(a: number, b: number): number {
	if (a >= 0 && b >= 0) {
		return erfc(a) - erfc(b);
	}
	if (a <= 0 && b <= 0) {
		return erfc(-b) - erfc(-a);
	}
	return erf(b) - erf(a);
}

function erfImpl(args: FunctionArg[], ctx: FunctionContext) {
	const lower = engNum(args[0], ctx);
	if (isCellError(lower)) {
		return lower;
	}
	if (isOmitted(args[1])) {
		return erf(lower);
	}
	const upper = engNum(args[1], ctx);
	return isCellError(upper) ? upper : erfBetween(lower, upper);
}

function unary(f: (x: number) => number) {
	return (args: FunctionArg[], ctx: FunctionContext) => {
		const x = engNum(args[0], ctx);
		return isCellError(x) ? x : f(x);
	};
}

export const ERF_FUNCTIONS: FunctionDefinition[] = [
	{ name: 'ERF', minArgs: 1, maxArgs: 2, params: ['value'], impl: erfImpl },
	{
		name: 'ERF.PRECISE',
		minArgs: 1,
		maxArgs: 1,
		params: ['value'],
		filePrefix: '_xlfn.',
		impl: unary(erf),
	},
	{ name: 'ERFC', minArgs: 1, maxArgs: 1, params: ['value'], impl: unary(erfc) },
	{
		name: 'ERFC.PRECISE',
		minArgs: 1,
		maxArgs: 1,
		params: ['value'],
		filePrefix: '_xlfn.',
		impl: unary(erfc),
	},
];
