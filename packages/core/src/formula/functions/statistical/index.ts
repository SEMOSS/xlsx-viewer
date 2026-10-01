import type { FunctionDefinition } from '../../types';
import { COUNTING_FUNCTIONS } from './counting';
import { DESCRIPTIVE_FUNCTIONS } from './descriptive';
import { DISTRIBUTION_FUNCTIONS } from './distributions';
import { REGRESSION_FUNCTIONS } from './regression';

/** Every statistical worksheet function. */
export const STATISTICAL_FUNCTIONS: FunctionDefinition[] = [
	...COUNTING_FUNCTIONS,
	...DESCRIPTIVE_FUNCTIONS,
	...DISTRIBUTION_FUNCTIONS,
	...REGRESSION_FUNCTIONS,
];
