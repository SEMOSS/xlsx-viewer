import type { FunctionDefinition } from '../../types';
import { ARITHMETIC_FUNCTIONS } from './arithmetic';
import { COMBINATORICS_FUNCTIONS } from './combinatorics';
import { MATRIX_FUNCTIONS } from './matrix';
import { RADIX_FUNCTIONS } from './radix';
import { RANDOM_FUNCTIONS } from './random';
import { ROUND_FUNCTIONS } from './round';
import { SUBTOTAL_FUNCTIONS } from './subtotal';
import { SUM_FUNCTIONS } from './sums';
import { SUMIF_FUNCTIONS } from './sumif';
import { TRIG_FUNCTIONS } from './trig';

/** Math and trigonometry functions. */
export const MATH_FUNCTIONS: FunctionDefinition[] = [
	...ARITHMETIC_FUNCTIONS,
	...TRIG_FUNCTIONS,
	...ROUND_FUNCTIONS,
	...COMBINATORICS_FUNCTIONS,
	...RADIX_FUNCTIONS,
	...MATRIX_FUNCTIONS,
	...SUM_FUNCTIONS,
	...SUMIF_FUNCTIONS,
	...SUBTOTAL_FUNCTIONS,
	...RANDOM_FUNCTIONS,
];

export { power, excelMod } from './arithmetic';
export {
	approxCeil,
	approxFloor,
	cleanNoise,
	roundDirected,
	roundHalfAway,
	shiftDecimal,
	snapToInteger,
} from './rounding';
export { CompensatedSum, sumOf } from './sums';
