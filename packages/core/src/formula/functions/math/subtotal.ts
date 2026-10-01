import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import { toNumber } from '../../coerce';
import type {
	FormulaValue,
	FunctionArg,
	FunctionContext,
	FunctionDefinition,
	RefValue,
	Scalar,
} from '../../types';
import { isArrayValue, isMultiRefValue, isRefValue, makeArray } from '../../types';
import { finite, integerArg, isMissingArg } from './args';
import * as stats from './aggregate-stats';
import { XLFN } from './define';
import { sumOf } from './sums';

/** Which cells SUBTOTAL / AGGREGATE leave out. */
interface SkipOptions {
	hidden: boolean;
	nested: boolean;
	errors: boolean;
}

interface Collected {
	numbers: number[];
	/** Non-blank values (COUNTA). */
	nonBlank: number;
	/** First error that was not skipped. */
	error?: CellError;
}

/**
 * Gather the values of SUBTOTAL / AGGREGATE arguments. References honour
 * the skip options (hidden rows, nested SUBTOTAL / AGGREGATE cells,
 * errors); arrays (AGGREGATE's array form only) honour the error option.
 */
function collect(
	args: readonly FunctionArg[],
	skip: SkipOptions,
	ctx: FunctionContext,
	allowArrays: boolean,
): Collected | CellError {
	const out: Collected = { numbers: [], nonBlank: 0 };
	const visit = (v: Scalar) => {
		if (v === null) {
			return;
		}
		if (isCellError(v)) {
			if (!skip.errors) {
				out.nonBlank++;
				out.error ??= v;
			}
			return;
		}
		out.nonBlank++;
		if (typeof v === 'number') {
			out.numbers.push(v);
		}
	};
	const visitRef = (ref: RefValue): CellError | undefined => {
		if (ref.sheetEnd !== undefined && ref.sheetEnd !== ref.sheet) {
			return ERRORS.VALUE;
		}
		ctx.forEachCell(ref, (v, row, col, sheet) => {
			if (skip.hidden && ctx.isRowHidden(sheet, row)) {
				return;
			}
			if (skip.nested && ctx.isSubtotalCell(sheet, row, col)) {
				return;
			}
			visit(v);
		});
		return undefined;
	};
	for (const arg of args) {
		if (isMissingArg(arg)) {
			continue;
		}
		let error: CellError | undefined;
		if (isRefValue(arg)) {
			error = visitRef(arg);
		} else if (isMultiRefValue(arg)) {
			for (const ref of arg.refs) {
				error ??= visitRef(ref);
			}
		} else if (isArrayValue(arg) && allowArrays) {
			for (const row of arg.values) {
				row.forEach(visit);
			}
		} else {
			error = isCellError(arg) ? arg : ERRORS.VALUE;
		}
		if (error) {
			return error;
		}
	}
	return out;
}

/** Functions 1-13 shared by SUBTOTAL (1-11) and AGGREGATE (1-13). */
function reduce(fn: number, data: Collected): FormulaValue {
	if (fn === 2) {
		return data.numbers.length;
	}
	if (fn === 3) {
		return data.nonBlank;
	}
	if (data.error) {
		return data.error;
	}
	const v = data.numbers;
	switch (fn) {
		case 1:
			return finite(stats.average(v));
		case 4:
			return stats.max(v);
		case 5:
			return stats.min(v);
		case 6:
			return finite(stats.product(v));
		case 7:
			return finite(stats.stdev(v, true));
		case 8:
			return finite(stats.stdev(v, false));
		case 9:
			return finite(sumOf(v));
		case 10:
			return finite(stats.variance(v, true));
		case 11:
			return finite(stats.variance(v, false));
		case 12:
			return stats.median(v);
		default:
			return stats.modeSingle(v);
	}
}

const SUBTOTAL: FunctionDefinition = {
	name: 'SUBTOTAL',
	minArgs: 2,
	maxArgs: 255,
	params: ['value', 'ref'],
	impl: (args, ctx) => {
		const code = integerArg(args[0], ctx);
		if (typeof code !== 'number') {
			return code;
		}
		const ignoreHidden = code >= 101 && code <= 111;
		const fn = ignoreHidden ? code - 100 : code;
		if (fn < 1 || fn > 11) {
			return ERRORS.VALUE;
		}
		const data = collect(
			args.slice(1),
			{ hidden: ignoreHidden, nested: true, errors: false },
			ctx,
			false,
		);
		return isCellError(data) ? data : reduce(fn, data);
	},
};

/** AGGREGATE functions 14-19 (array form) applied to one `k`. */
function arrayForm(
	fn: number,
	values: readonly number[],
	kValue: Scalar,
	ctx: FunctionContext,
): Scalar {
	const k = toNumber(kValue, ctx.date1904);
	if (typeof k !== 'number') {
		return k;
	}
	let r: number | CellError;
	switch (fn) {
		case 14:
			r = stats.kth(values, k, true);
			break;
		case 15:
			r = stats.kth(values, k, false);
			break;
		case 16:
			r = stats.percentileInc(values, k);
			break;
		case 17:
			r = stats.quartileInc(values, k);
			break;
		case 18:
			r = stats.percentileExc(values, k);
			break;
		default:
			r = stats.quartileExc(values, k);
	}
	return finite(r);
}

/** The `k` argument of the array form; it may itself be an array (one result per element). */
function kValues(arg: FunctionArg, ctx: FunctionContext): Scalar | Scalar[][] {
	if (isMissingArg(arg)) {
		return null;
	}
	if (isRefValue(arg)) {
		const r = arg.range;
		if (r.top === r.bottom && r.left === r.right) {
			return ctx.getCellValue(arg.sheet, r.top, r.left);
		}
		return ctx.resolveRef(arg).values;
	}
	if (isMultiRefValue(arg)) {
		return ERRORS.VALUE;
	}
	if (isArrayValue(arg)) {
		return arg.rows === 1 && arg.cols === 1 ? arg.values[0][0] : arg.values;
	}
	return typeof arg === 'object' && arg !== null && 'kind' in arg ? ERRORS.VALUE : (arg as Scalar);
}

const AGGREGATE: FunctionDefinition = {
	name: 'AGGREGATE',
	minArgs: 3,
	maxArgs: 255,
	params: ['value', 'value', 'any', 'any'],
	filePrefix: XLFN,
	impl: (args, ctx) => {
		const fn = integerArg(args[0], ctx);
		if (typeof fn !== 'number') {
			return fn;
		}
		const option = integerArg(args[1], ctx);
		if (typeof option !== 'number') {
			return option;
		}
		if (fn < 1 || fn > 19 || option < 0 || option > 7) {
			return ERRORS.VALUE;
		}
		const skip: SkipOptions = {
			nested: option <= 3,
			hidden: option % 2 === 1,
			errors: option === 2 || option === 3 || option === 6 || option === 7,
		};
		if (fn <= 13) {
			const data = collect(args.slice(2), skip, ctx, false);
			return isCellError(data) ? data : reduce(fn, data);
		}
		if (args.length !== 4) {
			return ERRORS.VALUE;
		}
		const data = collect([args[2]], skip, ctx, true);
		if (isCellError(data)) {
			return data;
		}
		if (data.error) {
			return data.error;
		}
		const k = kValues(args[3], ctx);
		if (Array.isArray(k)) {
			return makeArray(k.map((row) => row.map((kv) => arrayForm(fn, data.numbers, kv, ctx))));
		}
		return arrayForm(fn, data.numbers, k, ctx);
	},
};

export const SUBTOTAL_FUNCTIONS: FunctionDefinition[] = [SUBTOTAL, AGGREGATE];
