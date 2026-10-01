import { ERRORS, isCellError } from '../../values/errors';
import type { ArrayValue, FormulaValue, RefValue, Scalar } from '../types';
import { isArrayValue, isLambdaValue, isMultiRefValue, isRefValue } from '../types';

/**
 * Conversion of raw evaluation results into what a cell can hold. A cell
 * never shows "blank" for a formula: a reference to an empty cell (or an
 * empty argument) yields 0, like Excel.
 */

/** Minimal cell reader needed to dereference results. */
export interface ResultReader {
	getCellValue(sheet: number, row: number, col: number): Scalar;
	resolveRef(ref: RefValue): ArrayValue;
}

/** A scalar as stored in a cell: blank becomes 0, non-finite numbers `#NUM!`, `-0` is 0. */
export function cellScalar(v: unknown): Scalar {
	if (v === null || v === undefined) {
		return 0;
	}
	if (typeof v === 'number') {
		return Number.isFinite(v) ? (v === 0 ? 0 : v) : ERRORS.NUM;
	}
	if (typeof v === 'string' || typeof v === 'boolean') {
		return v;
	}
	if (isCellError(v)) {
		return v;
	}
	// A missing-argument node or a nested non-scalar.
	return typeof v === 'object' && (v as { type?: unknown }).type === 'missing' ? 0 : ERRORS.VALUE;
}

function is3d(ref: RefValue): boolean {
	return ref.sheetEnd !== undefined && ref.sheetEnd !== ref.sheet;
}

/**
 * Legacy (pre-dynamic-array) cell result: a reference is implicitly
 * intersected with the formula cell, an array contributes its top-left value.
 */
export function legacyCellResult(
	v: FormulaValue,
	reader: ResultReader,
	row: number,
	col: number,
): Scalar {
	if (isRefValue(v)) {
		if (is3d(v)) {
			return ERRORS.REF;
		}
		const r = v.range;
		if (r.top === r.bottom && r.left === r.right) {
			return cellScalar(reader.getCellValue(v.sheet, r.top, r.left));
		}
		if (r.left === r.right && row >= r.top && row <= r.bottom) {
			return cellScalar(reader.getCellValue(v.sheet, row, r.left));
		}
		if (r.top === r.bottom && col >= r.left && col <= r.right) {
			return cellScalar(reader.getCellValue(v.sheet, r.top, col));
		}
		return ERRORS.VALUE;
	}
	if (isMultiRefValue(v)) {
		return v.refs.length === 1 ? legacyCellResult(v.refs[0], reader, row, col) : ERRORS.VALUE;
	}
	if (isArrayValue(v)) {
		return v.rows > 0 && v.cols > 0 ? cellScalar(v.values[0][0]) : ERRORS.CALC;
	}
	if (isLambdaValue(v)) {
		return ERRORS.CALC;
	}
	return cellScalar(v);
}

/** Copy of an array with every element converted by {@link cellScalar}. */
export function cellArray(a: ArrayValue): ArrayValue {
	const values: Scalar[][] = new Array(a.rows);
	for (let r = 0; r < a.rows; r++) {
		const src = a.values[r];
		const line: Scalar[] = new Array(a.cols);
		for (let c = 0; c < a.cols; c++) {
			line[c] = cellScalar(src[c]);
		}
		values[r] = line;
	}
	return { kind: 'array', rows: a.rows, cols: a.cols, values };
}

/**
 * Array-aware cell result (dynamic arrays and CSE formulas): a single value,
 * or an array of cell values for results larger than 1x1. References are
 * materialised; an empty array is `#CALC!`.
 */
export function arrayCellResult(v: FormulaValue, reader: ResultReader): Scalar | ArrayValue {
	if (isRefValue(v)) {
		if (is3d(v)) {
			return ERRORS.REF;
		}
		const r = v.range;
		if (r.top === r.bottom && r.left === r.right) {
			return cellScalar(reader.getCellValue(v.sheet, r.top, r.left));
		}
		return arrayCellResult(reader.resolveRef(v), reader);
	}
	if (isMultiRefValue(v)) {
		return v.refs.length === 1 ? arrayCellResult(v.refs[0], reader) : ERRORS.VALUE;
	}
	if (isArrayValue(v)) {
		if (v.rows === 0 || v.cols === 0) {
			return ERRORS.CALC;
		}
		if (v.rows === 1 && v.cols === 1) {
			return cellScalar(v.values[0][0]);
		}
		return cellArray(v);
	}
	if (isLambdaValue(v)) {
		return ERRORS.CALC;
	}
	return cellScalar(v);
}
