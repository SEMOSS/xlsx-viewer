import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import { toBoolean } from '../../coerce';
import type { FunctionArg, FunctionContext, FunctionDefinition, Scalar } from '../../types';
import { isArrayValue, isLambdaValue, isMultiRefValue, isRefValue } from '../../types';
import { isMissingNode } from './lift';
import { withFilePrefix } from './prefix';

/**
 * AND / OR / XOR / NOT and the TRUE / FALSE functions.
 *
 * AND, OR and XOR accept ranges and arrays: logical values and numbers in
 * them count, text and blanks are ignored. Arguments typed directly are
 * coerced (the text "TRUE" / "FALSE" counts, other text is `#VALUE!`). With
 * no logical value at all the result is `#VALUE!`.
 */

interface Tally {
	count: number;
	trues: number;
}

function tallyValue(v: Scalar, fromRef: boolean, t: Tally): CellError | undefined {
	if (isCellError(v)) {
		return v;
	}
	if (typeof v === 'boolean') {
		t.count++;
		if (v) {
			t.trues++;
		}
		return undefined;
	}
	if (typeof v === 'number') {
		t.count++;
		if (v !== 0) {
			t.trues++;
		}
		return undefined;
	}
	if (v === null || fromRef) {
		return undefined;
	}
	const b = toBoolean(v);
	if (isCellError(b)) {
		return b;
	}
	t.count++;
	if (b) {
		t.trues++;
	}
	return undefined;
}

function tally(args: readonly FunctionArg[], ctx: FunctionContext): Tally | CellError {
	const t: Tally = { count: 0, trues: 0 };
	for (const arg of args) {
		let error: CellError | undefined;
		if (isRefValue(arg) || isMultiRefValue(arg)) {
			const refs = isRefValue(arg) ? [arg] : arg.refs;
			for (const ref of refs) {
				ctx.forEachCell(ref, (v) => {
					error ??= tallyValue(v, true, t);
				});
			}
		} else if (isArrayValue(arg)) {
			for (const row of arg.values) {
				for (const v of row) {
					error ??= tallyValue(v, true, t);
				}
			}
		} else if (isMissingNode(arg)) {
			// An empty argument counts as FALSE.
			t.count++;
		} else if (isLambdaValue(arg)) {
			error = ERRORS.CALC;
		} else {
			error = tallyValue(arg as Scalar, false, t);
		}
		if (error) {
			return error;
		}
	}
	return t.count === 0 ? ERRORS.VALUE : t;
}

function aggregate(name: string, decide: (t: Tally) => boolean): FunctionDefinition {
	return withFilePrefix({
		name,
		minArgs: 1,
		maxArgs: 255,
		params: ['any'],
		impl: (args, ctx) => {
			const t = tally(args, ctx);
			return isCellError(t) ? t : decide(t);
		},
	});
}

const NOT: FunctionDefinition = {
	name: 'NOT',
	minArgs: 1,
	maxArgs: 1,
	params: ['value'],
	impl: (args) => {
		const v = isMissingNode(args[0]) ? null : (args[0] as Scalar);
		const b = toBoolean(v);
		return isCellError(b) ? b : !b;
	},
};

function constant(name: string, value: boolean): FunctionDefinition {
	return { name, minArgs: 0, maxArgs: 0, params: [], impl: () => value };
}

export const BOOLEAN_FUNCTIONS: FunctionDefinition[] = [
	aggregate('AND', (t) => t.trues === t.count),
	aggregate('OR', (t) => t.trues > 0),
	aggregate('XOR', (t) => t.trues % 2 === 1),
	NOT,
	constant('TRUE', true),
	constant('FALSE', false),
];
