import type { FunctionDefinition } from '../../types';
import { ITERATE_FUNCTIONS } from './iterate';
import { LET_LAMBDA_FUNCTIONS } from './let-lambda';

/** LET, LAMBDA, ISOMITTED and the LAMBDA helpers MAP, REDUCE, SCAN, MAKEARRAY, BYROW, BYCOL. */
export const LAMBDA_FUNCTIONS: FunctionDefinition[] = [
	...LET_LAMBDA_FUNCTIONS,
	...ITERATE_FUNCTIONS,
];
