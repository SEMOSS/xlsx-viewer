import type { FunctionDefinition } from '../../../types';
import { CORRELATION_FUNCTIONS } from './correlation';
import { LINEST_FUNCTIONS } from './linest';
import { TREND_FUNCTIONS } from './trend';

/** Correlation, covariance and least-squares regression. */
export const REGRESSION_FUNCTIONS: FunctionDefinition[] = [
	...CORRELATION_FUNCTIONS,
	...LINEST_FUNCTIONS,
	...TREND_FUNCTIONS,
];

export { type LeastSquaresFit, leastSquares } from './least-squares';
