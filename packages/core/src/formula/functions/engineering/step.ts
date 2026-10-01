/**
 * Comparison helpers: DELTA (equality test) and GESTEP (threshold test).
 */
import { isCellError } from '../../../values/errors';
import type { FunctionArg, FunctionContext, FunctionDefinition } from '../../types';
import { engNumbers } from './args';

function compare(test: (a: number, b: number) => boolean) {
	return (args: FunctionArg[], ctx: FunctionContext) => {
		const n = engNumbers(args, ctx, [null, 0]);
		if (isCellError(n)) {
			return n;
		}
		return test(n[0], n[1]) ? 1 : 0;
	};
}

export const STEP_FUNCTIONS: FunctionDefinition[] = [
	{ name: 'DELTA', minArgs: 1, maxArgs: 2, params: ['value'], impl: compare((a, b) => a === b) },
	{ name: 'GESTEP', minArgs: 1, maxArgs: 2, params: ['value'], impl: compare((a, b) => a >= b) },
];
