import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import { compareScalars, toBoolean, toNumber } from '../../coerce';
import { compareNumbers } from '../../evaluator/operators';
import type {
	ArrayValue,
	FormulaValue,
	FunctionArg,
	FunctionContext,
	FunctionDefinition,
	Scalar,
} from '../../types';
import { isArrayValue } from '../../types';
import { branchValue, buildArray, elementAt, evalLifted, liftedShape, liftValue } from './lift';
import { withFilePrefix } from './prefix';

/**
 * Conditional functions. All parameters are lazy so only the branch that is
 * taken gets evaluated; array conditions (dynamic arrays) are lifted
 * element-wise and then both branches are evaluated once.
 */

function hasTrue(v: ArrayValue, want: boolean): boolean {
	for (const row of v.values) {
		for (const x of row) {
			const b = toBoolean(x);
			if (b === want) {
				return true;
			}
		}
	}
	return false;
}

const IF: FunctionDefinition = {
	name: 'IF',
	minArgs: 2,
	maxArgs: 3,
	params: ['lazy'],
	impl: (args, ctx) => {
		const cond = evalLifted(args[0], ctx);
		const otherwise = (): FormulaValue => (args.length > 2 ? branchValue(args[2], ctx) : false);
		if (!isArrayValue(cond)) {
			const b = toBoolean(cond);
			if (isCellError(b)) {
				return b;
			}
			return b ? branchValue(args[1], ctx) : otherwise();
		}
		const yes = hasTrue(cond, true) ? liftValue(branchValue(args[1], ctx), ctx) : null;
		const no = hasTrue(cond, false) ? liftValue(otherwise(), ctx) : null;
		const shape = liftedShape([cond, yes ?? null, no ?? null]);
		return buildArray(shape.rows, shape.cols, (r, c) => {
			const b = toBoolean(elementAt(cond, r, c));
			if (isCellError(b)) {
				return b;
			}
			const pick = b ? yes : no;
			return pick === null ? ERRORS.NA : elementAt(pick, r, c);
		});
	},
};

function errorTrap(name: string, matches: (e: CellError) => boolean): FunctionDefinition {
	return withFilePrefix({
		name,
		minArgs: 2,
		maxArgs: 2,
		params: ['lazy'],
		impl: (args, ctx) => {
			const v = evalLifted(args[0], ctx);
			if (!isArrayValue(v)) {
				return isCellError(v) && matches(v) ? branchValue(args[1], ctx) : v;
			}
			let trapped = false;
			for (const row of v.values) {
				if (row.some((x) => isCellError(x) && matches(x))) {
					trapped = true;
					break;
				}
			}
			if (!trapped) {
				return v;
			}
			const fallback = liftValue(branchValue(args[1], ctx), ctx);
			const shape = liftedShape([v, fallback]);
			return buildArray(shape.rows, shape.cols, (r, c) => {
				const x = elementAt(v, r, c);
				return isCellError(x) && matches(x) ? elementAt(fallback, r, c) : x;
			});
		},
	});
}

const IFS: FunctionDefinition = withFilePrefix({
	name: 'IFS',
	minArgs: 2,
	maxArgs: 254,
	params: ['lazy'],
	impl: (args, ctx) => {
		if (args.length % 2 !== 0) {
			return ERRORS.VALUE;
		}
		for (let i = 0; i < args.length; i += 2) {
			const c = evalLifted(args[i], ctx);
			if (isArrayValue(c)) {
				return ifsArray(args, i, c, ctx);
			}
			const b = toBoolean(c);
			if (isCellError(b)) {
				return b;
			}
			if (b) {
				return branchValue(args[i + 1], ctx);
			}
		}
		return ERRORS.NA;
	},
});

function ifsArray(
	args: readonly FunctionArg[],
	start: number,
	first: ArrayValue,
	ctx: FunctionContext,
): FormulaValue {
	const conds: Array<Scalar | ArrayValue> = [first];
	const values: Array<Scalar | ArrayValue> = [];
	for (let i = start; i < args.length; i += 2) {
		if (i > start) {
			conds.push(evalLifted(args[i], ctx));
		}
		values.push(liftValue(branchValue(args[i + 1], ctx), ctx));
	}
	const shape = liftedShape([...conds, ...values]);
	return buildArray(shape.rows, shape.cols, (r, c) => {
		for (let k = 0; k < conds.length; k++) {
			const b = toBoolean(elementAt(conds[k], r, c));
			if (isCellError(b)) {
				return b;
			}
			if (b) {
				return elementAt(values[k], r, c);
			}
		}
		return ERRORS.NA;
	});
}

/** SWITCH equality: same type (blank matches 0, "" and FALSE), text case-insensitive. */
function switchMatches(a: Scalar, b: Scalar): boolean {
	if (typeof a === 'number' && typeof b === 'number') {
		return compareNumbers(a, b) === 0;
	}
	return compareScalars(a, b) === 0;
}

function firstScalar(v: Scalar | ArrayValue): Scalar {
	return isArrayValue(v) ? v.values[0][0] : v;
}

const SWITCH: FunctionDefinition = withFilePrefix({
	name: 'SWITCH',
	minArgs: 3,
	maxArgs: 254,
	params: ['lazy'],
	impl: (args, ctx) => {
		const pairs = Math.floor((args.length - 1) / 2);
		const hasDefault = (args.length - 1) % 2 === 1;
		const subject = evalLifted(args[0], ctx);
		if (isArrayValue(subject)) {
			const keys: Scalar[] = [];
			const results: Array<Scalar | ArrayValue> = [];
			for (let k = 0; k < pairs; k++) {
				keys.push(firstScalar(evalLifted(args[1 + 2 * k], ctx)));
				results.push(liftValue(branchValue(args[2 + 2 * k], ctx), ctx));
			}
			const fallback = hasDefault
				? liftValue(branchValue(args[args.length - 1], ctx), ctx)
				: ERRORS.NA;
			const shape = liftedShape([subject, ...results, fallback]);
			return buildArray(shape.rows, shape.cols, (r, c) => {
				const x = elementAt(subject, r, c);
				if (isCellError(x)) {
					return x;
				}
				for (let k = 0; k < keys.length; k++) {
					if (isCellError(keys[k])) {
						return keys[k];
					}
					if (switchMatches(x, keys[k])) {
						return elementAt(results[k], r, c);
					}
				}
				return elementAt(fallback, r, c);
			});
		}
		if (isCellError(subject)) {
			return subject;
		}
		for (let k = 0; k < pairs; k++) {
			const key = firstScalar(evalLifted(args[1 + 2 * k], ctx));
			if (isCellError(key)) {
				return key;
			}
			if (switchMatches(subject, key)) {
				return branchValue(args[2 + 2 * k], ctx);
			}
		}
		return hasDefault ? branchValue(args[args.length - 1], ctx) : ERRORS.NA;
	},
});

function chooseIndex(v: Scalar, count: number, ctx: FunctionContext): number | CellError {
	const n = toNumber(v, ctx.date1904);
	if (typeof n !== 'number') {
		return n;
	}
	const k = Math.trunc(n);
	return k < 1 || k > count ? ERRORS.VALUE : k;
}

const CHOOSE: FunctionDefinition = {
	name: 'CHOOSE',
	minArgs: 2,
	maxArgs: 255,
	params: ['lazy'],
	impl: (args, ctx) => {
		const count = args.length - 1;
		const index = evalLifted(args[0], ctx);
		if (!isArrayValue(index)) {
			const k = chooseIndex(index, count, ctx);
			return isCellError(k) ? k : branchValue(args[k], ctx);
		}
		const chosen = new Map<number, Scalar | ArrayValue>();
		for (const row of index.values) {
			for (const x of row) {
				const k = chooseIndex(x, count, ctx);
				if (typeof k === 'number' && !chosen.has(k)) {
					chosen.set(k, liftValue(branchValue(args[k], ctx), ctx));
				}
			}
		}
		const shape = liftedShape([index, ...chosen.values()]);
		return buildArray(shape.rows, shape.cols, (r, c) => {
			const k = chooseIndex(elementAt(index, r, c), count, ctx);
			return typeof k === 'number' ? elementAt(chosen.get(k)!, r, c) : k;
		});
	},
};

export const CONDITION_FUNCTIONS: FunctionDefinition[] = [
	IF,
	IFS,
	errorTrap('IFERROR', () => true),
	errorTrap('IFNA', (e) => e.code === '#N/A'),
	SWITCH,
	CHOOSE,
];
