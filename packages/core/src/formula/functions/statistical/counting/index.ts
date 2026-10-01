import type { FunctionDefinition } from '../../../types';
import { CONDITIONAL_FUNCTIONS } from './conditional';
import { COUNT_FUNCTIONS } from './count';

/** COUNT, COUNTA, COUNTBLANK and the criteria-based aggregates. */
export const COUNTING_FUNCTIONS: FunctionDefinition[] = [
	...COUNT_FUNCTIONS,
	...CONDITIONAL_FUNCTIONS,
];

export * from './grid';
export { cellCount } from './count';
