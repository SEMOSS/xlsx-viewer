import type { FunctionDefinition } from '../../../types';
import { AVERAGE_FUNCTIONS } from './averages';
import { DISPERSION_FUNCTIONS } from './dispersion';
import { EXTREME_FUNCTIONS } from './extremes';
import { FREQUENCY_FUNCTIONS } from './frequency';
import { MODE_FUNCTIONS } from './mode';
import { PERCENTILE_FUNCTIONS } from './percentile';
import { RANK_FUNCTIONS } from './rank';

/** Averages, dispersion, extremes, modes, percentiles, ranks and frequency tables. */
export const DESCRIPTIVE_FUNCTIONS: FunctionDefinition[] = [
	...AVERAGE_FUNCTIONS,
	...DISPERSION_FUNCTIONS,
	...EXTREME_FUNCTIONS,
	...MODE_FUNCTIONS,
	...PERCENTILE_FUNCTIONS,
	...RANK_FUNCTIONS,
	...FREQUENCY_FUNCTIONS,
];

export * from './core';
export { median } from './averages';
export { modes } from './mode';
export { percentileExc, percentileInc, percentRank } from './percentile';
