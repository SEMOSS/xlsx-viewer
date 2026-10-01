import type { CellError } from '../types/cell';
import { ERRORS, isCellError } from '../values/errors';
import { parseNumericText } from '../values/parse-number';
import type { ArrayValue, FunctionArg, FunctionContext, Scalar } from './types';
import { isArrayValue, isMultiRefValue, isRefValue } from './types';

/**
 * Excel's implicit number-to-text conversion (used by `&`, text functions
 * and comparisons): up to 15 significant digits, no grouping, scientific
 * notation for very large or very small magnitudes.
 */
export function numberToText(n: number): string {
	if (Number.isNaN(n) || !Number.isFinite(n)) {
		return '#NUM!';
	}
	if (n === 0) {
		return '0';
	}
	const rounded = Number(n.toPrecision(15));
	const abs = Math.abs(rounded);
	const exp = Math.floor(Math.log10(abs));
	if (exp >= 15 || exp < -6) {
		const [mant, e] = rounded.toExponential(14).split('e');
		const mantissa = mant.includes('.') ? mant.replace(/0+$/, '').replace(/\.$/, '') : mant;
		const expNum = Number(e);
		const expText = Math.abs(expNum).toString().padStart(2, '0');
		return `${mantissa}E${expNum < 0 ? '-' : '+'}${expText}`;
	}
	const decimals = Math.max(0, Math.min(20, 14 - exp));
	let s = rounded.toFixed(decimals);
	if (s.includes('.')) {
		s = s.replace(/0+$/, '').replace(/\.$/, '');
	}
	return s === '-0' ? '0' : s;
}

/** Coerce to a number following formula rules (blank = 0, TRUE = 1, numeric text parsed). */
export function toNumber(v: Scalar, date1904 = false): number | CellError {
	if (typeof v === 'number') {
		return v;
	}
	if (v === null) {
		return 0;
	}
	if (typeof v === 'boolean') {
		return v ? 1 : 0;
	}
	if (typeof v === 'string') {
		if (v.trim() === '') {
			return ERRORS.VALUE;
		}
		const parsed = parseNumericText(v, { date1904 });
		return parsed ? parsed.value : ERRORS.VALUE;
	}
	return v;
}

/** Coerce to text (numbers use {@link numberToText}, booleans become TRUE/FALSE, blank is ""). */
export function toText(v: Scalar): string | CellError {
	if (typeof v === 'string') {
		return v;
	}
	if (typeof v === 'number') {
		return numberToText(v);
	}
	if (typeof v === 'boolean') {
		return v ? 'TRUE' : 'FALSE';
	}
	if (v === null) {
		return '';
	}
	return v;
}

/** Coerce to boolean (numbers: non-zero is TRUE; text must be TRUE/FALSE; blank is FALSE). */
export function toBoolean(v: Scalar): boolean | CellError {
	if (typeof v === 'boolean') {
		return v;
	}
	if (typeof v === 'number') {
		return v !== 0;
	}
	if (v === null) {
		return false;
	}
	if (typeof v === 'string') {
		const u = v.trim().toUpperCase();
		if (u === 'TRUE') {
			return true;
		}
		if (u === 'FALSE') {
			return false;
		}
		return ERRORS.VALUE;
	}
	return v;
}

/** Truncate toward zero after numeric coercion (argument conventions such as `num_chars`). */
export function toInteger(v: Scalar, date1904 = false): number | CellError {
	const n = toNumber(v, date1904);
	return typeof n === 'number' ? Math.trunc(n) : n;
}

/** A valid finite result, or `#NUM!`. */
export function checkNumber(n: number): number | CellError {
	return Number.isFinite(n) ? n : ERRORS.NUM;
}

let collator: Intl.Collator | undefined;

/** Case-insensitive but accent-sensitive text ordering (Excel: "a"="A" but "a"<>"á"), used for comparisons and sorting. */
export function compareText(a: string, b: string): number {
	collator ??= new Intl.Collator('en', { sensitivity: 'accent', numeric: false });
	return collator.compare(a, b);
}

function typeRank(v: Scalar): number {
	if (typeof v === 'number') {
		return 0;
	}
	if (typeof v === 'string') {
		return 1;
	}
	if (typeof v === 'boolean') {
		return 2;
	}
	return 3;
}

/**
 * Compare two non-error scalars with Excel semantics: numbers < text <
 * booleans, text compared case-insensitively, blank behaves as 0, "" or
 * FALSE depending on the other operand. Returns <0, 0 or >0.
 */
export function compareScalars(a: Scalar, b: Scalar): number {
	let x = a;
	let y = b;
	if (x === null && y === null) {
		return 0;
	}
	if (x === null) {
		x = typeof y === 'string' ? '' : typeof y === 'boolean' ? false : 0;
	}
	if (y === null) {
		y = typeof x === 'string' ? '' : typeof x === 'boolean' ? false : 0;
	}
	const rx = typeRank(x);
	const ry = typeRank(y);
	if (rx !== ry) {
		return rx - ry;
	}
	if (typeof x === 'number' && typeof y === 'number') {
		return x < y ? -1 : x > y ? 1 : 0;
	}
	if (typeof x === 'string' && typeof y === 'string') {
		return compareText(x, y);
	}
	if (typeof x === 'boolean' && typeof y === 'boolean') {
		return x === y ? 0 : x ? 1 : -1;
	}
	return 0;
}

/** Equality used by `=` and lookups (case-insensitive text). Errors are never equal. */
export function scalarEquals(a: Scalar, b: Scalar): boolean {
	if (isCellError(a) || isCellError(b)) {
		return false;
	}
	return compareScalars(a, b) === 0;
}

/** First error found in a list of scalars, if any. */
export function firstError(values: readonly Scalar[]): CellError | undefined {
	for (const v of values) {
		if (isCellError(v)) {
			return v;
		}
	}
	return undefined;
}

/**
 * Visit every scalar of an argument as delivered for an `array` or `any`
 * parameter. `fromRef` is true for values that came from cells or arrays
 * (as opposed to a directly typed scalar argument), which matters for the
 * SUM-family rule that text and booleans inside ranges are ignored.
 */
export function forEachScalar(
	arg: FunctionArg,
	ctx: FunctionContext,
	visit: (value: Scalar, fromRef: boolean) => void,
): void {
	if (isRefValue(arg)) {
		ctx.forEachCell(arg, (v) => visit(v, true));
		return;
	}
	if (isMultiRefValue(arg)) {
		for (const r of arg.refs) {
			ctx.forEachCell(r, (v) => visit(v, true));
		}
		return;
	}
	if (isArrayValue(arg)) {
		for (const row of arg.values) {
			for (const v of row) {
				visit(v, true);
			}
		}
		return;
	}
	if (typeof arg === 'object' && arg !== null && 'kind' in arg && arg.kind === 'lambda') {
		visit(ERRORS.CALC, false);
		return;
	}
	if (typeof arg === 'object' && arg !== null && 'type' in arg && arg.type !== 'error') {
		// An AST node (lazy argument) should never reach here.
		visit(ERRORS.VALUE, false);
		return;
	}
	visit(arg as Scalar, false);
}

export interface CollectNumbersOptions {
	/** Count TRUE/FALSE found in references and arrays (the `*A` functions). Default false. */
	boolsInRefs?: boolean;
	/** Count text in references as 0 (the `*A` functions). Default false. */
	textInRefsAsZero?: boolean;
	/** Ignore errors instead of returning them (AGGREGATE options). Default false. */
	ignoreErrors?: boolean;
}

/**
 * Collect numbers from SUM-style arguments. Direct scalars are coerced (text
 * that is not numeric is `#VALUE!`); values inside references and arrays
 * count only when numeric, unless the options widen that. The first error
 * encountered is returned instead of the list.
 */
export function collectNumbers(
	args: readonly FunctionArg[],
	ctx: FunctionContext,
	opts: CollectNumbersOptions = {},
): number[] | CellError {
	const out: number[] = [];
	let error: CellError | undefined;
	for (const arg of args) {
		if (error) {
			break;
		}
		if (typeof arg === 'object' && arg !== null && 'type' in arg && arg.type === 'missing') {
			// An omitted argument (`SUM(1,)`, `AVERAGE(1,)`) is the number 0 in Excel.
			out.push(0);
			continue;
		}
		forEachScalar(arg, ctx, (v, fromRef) => {
			if (error) {
				return;
			}
			if (typeof v === 'number') {
				out.push(v);
				return;
			}
			if (isCellError(v)) {
				if (!opts.ignoreErrors) {
					error = v;
				}
				return;
			}
			if (fromRef) {
				if (typeof v === 'boolean' && opts.boolsInRefs) {
					out.push(v ? 1 : 0);
				} else if (typeof v === 'string' && opts.textInRefsAsZero) {
					out.push(0);
				}
				return;
			}
			if (v === null) {
				return;
			}
			const n = toNumber(v, ctx.date1904);
			if (typeof n === 'number') {
				out.push(n);
			} else if (!opts.ignoreErrors) {
				error = n;
			}
		});
	}
	return error ?? out;
}

/** Turn an `array` parameter value into an ArrayValue (it already is one; this narrows the type). */
export function asArray(arg: FunctionArg): ArrayValue {
	if (isArrayValue(arg)) {
		return arg;
	}
	return { kind: 'array', rows: 1, cols: 1, values: [[(arg as Scalar) ?? null]] };
}

/** Flatten an array row by row. */
export function flattenArray(arr: ArrayValue): Scalar[] {
	const out: Scalar[] = [];
	for (const row of arr.values) {
		for (const v of row) {
			out.push(v);
		}
	}
	return out;
}
