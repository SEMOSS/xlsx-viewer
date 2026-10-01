import { ERRORS } from '../../../values/errors';
import type { FunctionDefinition } from '../../types';
import { numberArg } from './args';

const RAND: FunctionDefinition = {
	name: 'RAND',
	minArgs: 0,
	maxArgs: 0,
	params: [],
	volatile: true,
	impl: (_args, ctx) => ctx.random(),
};

/** RANDBETWEEN(bottom, top): a uniform integer in [ceil(bottom), floor(top)]. */
const RANDBETWEEN: FunctionDefinition = {
	name: 'RANDBETWEEN',
	minArgs: 2,
	maxArgs: 2,
	params: ['value', 'value'],
	volatile: true,
	impl: (args, ctx) => {
		const bottom = numberArg(args[0], ctx);
		if (typeof bottom !== 'number') {
			return bottom;
		}
		const top = numberArg(args[1], ctx);
		if (typeof top !== 'number') {
			return top;
		}
		const lo = Math.ceil(bottom);
		const hi = Math.floor(top);
		if (lo > hi) {
			return ERRORS.NUM;
		}
		const r = lo + Math.floor(ctx.random() * (hi - lo + 1));
		return Math.min(r, hi);
	},
};

export const RANDOM_FUNCTIONS: FunctionDefinition[] = [RAND, RANDBETWEEN];
