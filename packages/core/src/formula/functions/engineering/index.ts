import type { FunctionDefinition } from '../../types';
import { BASE_FUNCTIONS } from './base-convert';
import { BESSEL_FUNCTIONS } from './bessel';
import { BITWISE_FUNCTIONS } from './bitwise';
import { COMPLEX_FUNCTIONS } from './complex-functions';
import { COMPLEX_MATH_FUNCTIONS } from './complex-math';
import { CONVERT_FUNCTIONS } from './convert';
import { ERF_FUNCTIONS } from './erf';
import { STEP_FUNCTIONS } from './step';

/** Engineering functions (BESSELJ, DEC2HEX, BITAND, CONVERT, ERF, IMSUM, ...). */
export const ENGINEERING_FUNCTIONS: FunctionDefinition[] = [
	...BESSEL_FUNCTIONS,
	...BASE_FUNCTIONS,
	...BITWISE_FUNCTIONS,
	...CONVERT_FUNCTIONS,
	...ERF_FUNCTIONS,
	...STEP_FUNCTIONS,
	...COMPLEX_FUNCTIONS,
	...COMPLEX_MATH_FUNCTIONS,
];

export { erf, erfc } from './erf';
