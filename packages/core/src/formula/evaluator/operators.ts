import { ERRORS, isCellError } from '../../values/errors';
import { compareScalars, toNumber, toText } from '../coerce';
import { broadcastElement } from '../invoke';
import type { ArrayValue, Scalar } from '../types';

/**
 * Scalar semantics of the arithmetic, text and comparison operators, plus
 * element-wise broadcasting over arrays. Operands arrive already reduced to
 * a scalar or an array (references are dereferenced by the caller).
 */

export type ValueOperator =
	| '+'
	| '-'
	| '*'
	| '/'
	| '^'
	| '&'
	| '='
	| '<>'
	| '<'
	| '>'
	| '<='
	| '>=';

/** Excel's maximum text length. */
const MAX_TEXT_LENGTH = 32767;

/** A finite number result (`-0` normalised), `#NUM!` otherwise. */
export function numberResult(n: number): Scalar {
	if (!Number.isFinite(n)) {
		return ERRORS.NUM;
	}
	return n === 0 ? 0 : n;
}

function power(x: number, y: number): Scalar {
	if (x === 0) {
		if (y === 0) {
			return ERRORS.NUM;
		}
		if (y < 0) {
			return ERRORS.DIV0;
		}
	}
	return numberResult(Math.pow(x, y));
}

function arithmetic(
	op: '+' | '-' | '*' | '/' | '^',
	a: Scalar,
	b: Scalar,
	date1904: boolean,
): Scalar {
	const x = typeof a === 'number' ? a : toNumber(a, date1904);
	if (typeof x !== 'number') {
		return x;
	}
	const y = typeof b === 'number' ? b : toNumber(b, date1904);
	if (typeof y !== 'number') {
		return y;
	}
	switch (op) {
		case '+':
			return numberResult(x + y);
		case '-':
			return numberResult(x - y);
		case '*':
			return numberResult(x * y);
		case '/':
			return y === 0 ? ERRORS.DIV0 : numberResult(x / y);
		case '^':
			return power(x, y);
	}
}

/**
 * Numeric comparison the way Excel does it: values that agree to 15
 * significant digits are equal (`=0.1+0.2=0.3` is TRUE).
 */
export function compareNumbers(a: number, b: number): number {
	if (a === b) {
		return 0;
	}
	const scale = Math.max(Math.abs(a), Math.abs(b));
	if (Math.abs(a - b) > scale * 1e-14) {
		return a < b ? -1 : 1;
	}
	const ra = Number(a.toPrecision(15));
	const rb = Number(b.toPrecision(15));
	return ra === rb ? 0 : a < b ? -1 : 1;
}

function comparison(op: '=' | '<>' | '<' | '>' | '<=' | '>=', a: Scalar, b: Scalar): Scalar {
	if (isCellError(a)) {
		return a;
	}
	if (isCellError(b)) {
		return b;
	}
	const c =
		typeof a === 'number' && typeof b === 'number' ? compareNumbers(a, b) : compareScalars(a, b);
	switch (op) {
		case '=':
			return c === 0;
		case '<>':
			return c !== 0;
		case '<':
			return c < 0;
		case '>':
			return c > 0;
		case '<=':
			return c <= 0;
		case '>=':
			return c >= 0;
	}
}

function concat(a: Scalar, b: Scalar): Scalar {
	const x = toText(a);
	if (typeof x !== 'string') {
		return x;
	}
	const y = toText(b);
	if (typeof y !== 'string') {
		return y;
	}
	return x.length + y.length > MAX_TEXT_LENGTH ? ERRORS.VALUE : x + y;
}

/** Apply a binary operator to two scalars. */
export function applyOperator(op: ValueOperator, a: Scalar, b: Scalar, date1904 = false): Scalar {
	switch (op) {
		case '&':
			return concat(a, b);
		case '=':
		case '<>':
		case '<':
		case '>':
		case '<=':
		case '>=':
			return comparison(op, a, b);
		default:
			return arithmetic(op, a, b, date1904);
	}
}

/** Unary minus with numeric coercion. */
export function negate(v: Scalar, date1904 = false): Scalar {
	const n = typeof v === 'number' ? v : toNumber(v, date1904);
	return typeof n === 'number' ? numberResult(-n) : n;
}

/** Postfix `%`. */
export function percent(v: Scalar, date1904 = false): Scalar {
	const n = typeof v === 'number' ? v : toNumber(v, date1904);
	return typeof n === 'number' ? numberResult(n / 100) : n;
}

function isArray(v: Scalar | ArrayValue): v is ArrayValue {
	return typeof v === 'object' && v !== null && (v as { kind?: unknown }).kind === 'array';
}

/**
 * Combine two operands element-wise. Single rows / columns (and scalars)
 * repeat to fill the larger shape; positions missing from a smaller array
 * are `#N/A`.
 */
export function broadcastBinary(
	a: Scalar | ArrayValue,
	b: Scalar | ArrayValue,
	f: (x: Scalar, y: Scalar) => Scalar,
): ArrayValue {
	const A = isArray(a) ? a : null;
	const B = isArray(b) ? b : null;
	const rows = Math.max(A ? A.rows : 1, B ? B.rows : 1);
	const cols = Math.max(A ? A.cols : 1, B ? B.cols : 1);
	const values: Scalar[][] = new Array(rows);
	for (let r = 0; r < rows; r++) {
		const line: Scalar[] = new Array(cols);
		for (let c = 0; c < cols; c++) {
			const x = A ? broadcastElement(A, r, c) : (a as Scalar);
			const y = B ? broadcastElement(B, r, c) : (b as Scalar);
			line[c] = f(x, y);
		}
		values[r] = line;
	}
	return { kind: 'array', rows, cols, values };
}

/** Apply a scalar function to every element of an array. */
export function mapArray(a: ArrayValue, f: (x: Scalar) => Scalar): ArrayValue {
	const values: Scalar[][] = new Array(a.rows);
	for (let r = 0; r < a.rows; r++) {
		const src = a.values[r];
		const line: Scalar[] = new Array(a.cols);
		for (let c = 0; c < a.cols; c++) {
			line[c] = f(src[c]);
		}
		values[r] = line;
	}
	return { kind: 'array', rows: a.rows, cols: a.cols, values };
}
