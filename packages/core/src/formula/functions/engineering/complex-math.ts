/**
 * Transcendental functions of complex numbers: IMEXP, IMLN, IMLOG10,
 * IMLOG2 and the trigonometric / hyperbolic family (IMSIN ... IMCSCH).
 */
import type { CellError } from '../../../types/cell';
import { ERRORS } from '../../../values/errors';
import type { FunctionDefinition } from '../../types';
import { complexResult } from './complex-functions';
import { cos, cosh, cot, csc, cx, exp, inverse, ln, sec, sin, sinh, tan, type Cx } from './complex';

/** `log_base(z)`; the logarithm of zero is #NUM!. */
function logBase(z: Cx, base: number): Cx | CellError {
	if (z.re === 0 && z.im === 0) {
		return ERRORS.NUM;
	}
	const l = ln(z);
	return base === Math.E ? l : cx(l.re / Math.log(base), l.im / Math.log(base));
}

const unary = (name: string, f: (z: Cx) => Cx | CellError, xlfn = false): FunctionDefinition => ({
	name,
	minArgs: 1,
	maxArgs: 1,
	params: ['value'],
	...(xlfn ? { filePrefix: '_xlfn.' as const } : {}),
	impl: complexResult(1, ([z]) => f(z)),
});

export const COMPLEX_MATH_FUNCTIONS: FunctionDefinition[] = [
	unary('IMEXP', exp),
	unary('IMLN', (z) => logBase(z, Math.E)),
	unary('IMLOG10', (z) => logBase(z, 10)),
	unary('IMLOG2', (z) => logBase(z, 2)),
	unary('IMSIN', sin),
	unary('IMCOS', cos),
	unary('IMTAN', tan, true),
	unary('IMCOT', cot, true),
	unary('IMSEC', sec, true),
	unary('IMCSC', csc, true),
	unary('IMSINH', sinh, true),
	unary('IMCOSH', cosh, true),
	unary('IMSECH', (z) => inverse(cosh(z)), true),
	unary('IMCSCH', (z) => inverse(sinh(z)), true),
];
