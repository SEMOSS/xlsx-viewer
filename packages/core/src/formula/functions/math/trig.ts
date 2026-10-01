import { ERRORS } from '../../../values/errors';
import type { FunctionDefinition } from '../../types';
import { binaryMath, unaryMath, XLFN } from './define';

/** COT, CSC, SEC and their hyperbolic forms reject |x| >= 2^27 with #NUM!. */
const RECIPROCAL_LIMIT = 2 ** 27;

function reciprocal(f: (x: number) => number, zeroIsDiv0: boolean) {
	return (x: number) => {
		if (Math.abs(x) >= RECIPROCAL_LIMIT) {
			return ERRORS.NUM;
		}
		if (zeroIsDiv0 && x === 0) {
			return ERRORS.DIV0;
		}
		return 1 / f(x);
	};
}

export const TRIG_FUNCTIONS: FunctionDefinition[] = [
	unaryMath('SIN', Math.sin),
	unaryMath('COS', Math.cos),
	unaryMath('TAN', Math.tan),
	unaryMath('ASIN', (x) => (Math.abs(x) > 1 ? ERRORS.NUM : Math.asin(x))),
	unaryMath('ACOS', (x) => (Math.abs(x) > 1 ? ERRORS.NUM : Math.acos(x))),
	unaryMath('ATAN', Math.atan),
	// Excel's argument order is ATAN2(x_num, y_num).
	binaryMath('ATAN2', (x, y) => (x === 0 && y === 0 ? ERRORS.DIV0 : Math.atan2(y, x))),
	unaryMath('SINH', Math.sinh),
	unaryMath('COSH', Math.cosh),
	unaryMath('TANH', Math.tanh),
	unaryMath('ASINH', Math.asinh),
	unaryMath('ACOSH', (x) => (x < 1 ? ERRORS.NUM : Math.acosh(x))),
	unaryMath('ATANH', (x) => (Math.abs(x) >= 1 ? ERRORS.NUM : Math.atanh(x))),
	unaryMath('COT', reciprocal(Math.tan, true), XLFN),
	unaryMath('CSC', reciprocal(Math.sin, true), XLFN),
	unaryMath('SEC', reciprocal(Math.cos, false), XLFN),
	unaryMath('COTH', reciprocal(Math.tanh, true), XLFN),
	unaryMath('CSCH', reciprocal(Math.sinh, true), XLFN),
	unaryMath('SECH', reciprocal(Math.cosh, false), XLFN),
	// The arccotangent is in (0, pi): ACOT(-1) is 3*pi/4.
	unaryMath('ACOT', (x) => Math.PI / 2 - Math.atan(x), XLFN),
	unaryMath(
		'ACOTH',
		(x) => (Math.abs(x) <= 1 ? ERRORS.NUM : 0.5 * Math.log((x + 1) / (x - 1))),
		XLFN,
	),
];
