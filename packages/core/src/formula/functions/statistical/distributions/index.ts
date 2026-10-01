import type { FunctionDefinition } from '../../../types';
import { BETA_FUNCTIONS } from './beta';
import { CHISQ_FUNCTIONS } from './chisq';
import { DISCRETE_FUNCTIONS } from './discrete';
import { F_FUNCTIONS } from './f';
import { GAMMA_FUNCTIONS } from './gamma';
import { HYPOTHESIS_FUNCTIONS } from './hypothesis';
import { NORMAL_FUNCTIONS } from './normal';
import { T_FUNCTIONS } from './t';

/** Probability distributions, their inverses and the hypothesis tests. */
export const DISTRIBUTION_FUNCTIONS: FunctionDefinition[] = [
	...NORMAL_FUNCTIONS,
	...GAMMA_FUNCTIONS,
	...CHISQ_FUNCTIONS,
	...BETA_FUNCTIONS,
	...T_FUNCTIONS,
	...F_FUNCTIONS,
	...DISCRETE_FUNCTIONS,
	...HYPOTHESIS_FUNCTIONS,
];
