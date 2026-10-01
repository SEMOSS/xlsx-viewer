import type { CellError } from '../../../../types/cell';
import { ERRORS, isCellError } from '../../../../values/errors';
import { forEachScalar } from '../../../coerce';
import type { FunctionArg, FunctionContext, FunctionDefinition } from '../../../types';
import { coerceArgs, define } from '../args';

/** Numbers inside a reference or array (other values ignored, errors returned). */
function refNumbers(arg: FunctionArg, ctx: FunctionContext): number[] | CellError {
	const out: number[] = [];
	let error: CellError | undefined;
	forEachScalar(arg, ctx, (v) => {
		if (error) {
			return;
		}
		if (typeof v === 'number') {
			out.push(v);
		} else if (isCellError(v)) {
			error = v;
		}
	});
	return error ?? out;
}

/** RANK / RANK.EQ (ties share the best rank) and RANK.AVG (ties share the mean rank). */
function rank(name: string, averageTies: boolean): FunctionDefinition {
	return define({
		name,
		minArgs: 2,
		maxArgs: 3,
		params: ['value', 'ref', 'value'],
		impl: (args, ctx) => {
			const v = coerceArgs(['n', { n: 0 }], [args[0], args[2]], ctx);
			if (isCellError(v)) {
				return v;
			}
			const [x, order] = v;
			const xs = refNumbers(args[1], ctx);
			if (isCellError(xs)) {
				return xs;
			}
			let better = 0;
			let equal = 0;
			for (const y of xs) {
				if (y === x) {
					equal++;
				} else if (order === 0 ? y > x : y < x) {
					better++;
				}
			}
			if (equal === 0) {
				return ERRORS.NA;
			}
			return averageTies ? better + (equal + 1) / 2 : better + 1;
		},
	});
}

export const RANK_FUNCTIONS: FunctionDefinition[] = [
	rank('RANK', false),
	rank('RANK.EQ', false),
	rank('RANK.AVG', true),
];
