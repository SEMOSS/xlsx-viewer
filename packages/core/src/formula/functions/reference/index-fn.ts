import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import type { ArrayValue, FormulaValue, FunctionDefinition, RefValue, Scalar } from '../../types';
import { isArrayValue, isLambdaValue, isMultiRefValue, isRefValue } from '../../types';
import { isMissingNode } from '../logical/lift';
import { integerOf, isOmittedArg } from './shared';

/**
 * INDEX(array, row_num, [column_num], [area_num]).
 *
 * Reference form: picks an area of a (multi-area) reference and returns a
 * reference to a cell, a whole row (column_num 0) or column (row_num 0), so
 * `A1:INDEX(...)` and `SUM(INDEX(...))` work. Array form: an element, row
 * or column of the array. For a single row / column the one index given
 * addresses along it. Indices beyond the size are `#REF!`.
 */

interface Pick {
	row: number;
	col: number;
}

function pick(
	height: number,
	width: number,
	rowNum: number,
	colNum: number | undefined,
): Pick | CellError {
	let row = rowNum;
	let col = colNum ?? 0;
	if (colNum === undefined && height === 1 && width > 1) {
		// A single row: the only index is the column.
		col = rowNum;
		row = 0;
	}
	if (row < 0 || col < 0) {
		return ERRORS.VALUE;
	}
	if (row > height || col > width) {
		return ERRORS.REF;
	}
	return { row, col };
}

function indexRef(ref: RefValue, p: Pick): RefValue {
	const r = ref.range;
	const top = p.row === 0 ? r.top : r.top + p.row - 1;
	const bottom = p.row === 0 ? r.bottom : top;
	const left = p.col === 0 ? r.left : r.left + p.col - 1;
	const right = p.col === 0 ? r.right : left;
	return { kind: 'ref', sheet: ref.sheet, range: { top, left, bottom, right } };
}

function indexArray(a: ArrayValue, p: Pick): FormulaValue {
	if (p.row > 0 && p.col > 0) {
		return a.values[p.row - 1][p.col - 1];
	}
	if (p.row > 0) {
		const line = a.values[p.row - 1];
		return line.length === 1
			? line[0]
			: { kind: 'array', rows: 1, cols: line.length, values: [line] };
	}
	if (p.col > 0) {
		const column: Scalar[][] = a.values.map((line) => [line[p.col - 1]]);
		return column.length === 1
			? column[0][0]
			: { kind: 'array', rows: column.length, cols: 1, values: column };
	}
	return a.rows === 1 && a.cols === 1 ? a.values[0][0] : a;
}

const INDEX: FunctionDefinition = {
	name: 'INDEX',
	minArgs: 1,
	maxArgs: 4,
	params: ['any', 'value', 'value', 'value'],
	impl: (args, ctx) => {
		const src = args[0];
		const rowNum = integerOf(args[1], ctx);
		if (isCellError(rowNum)) {
			return rowNum;
		}
		const hasCol = args.length > 2 && !isMissingNode(args[2]);
		const colNum = hasCol ? integerOf(args[2], ctx) : undefined;
		if (isCellError(colNum)) {
			return colNum;
		}
		const area = isOmittedArg(args, 3) ? 1 : integerOf(args[3], ctx);
		if (isCellError(area)) {
			return area;
		}
		if (isRefValue(src) || isMultiRefValue(src)) {
			const areas = isRefValue(src) ? [src] : src.refs;
			if (area < 1 || area > areas.length) {
				return ERRORS.REF;
			}
			const ref = areas[area - 1];
			if (ref.sheetEnd !== undefined && ref.sheetEnd !== ref.sheet) {
				return ERRORS.VALUE;
			}
			const r = ref.range;
			const p = pick(r.bottom - r.top + 1, r.right - r.left + 1, rowNum, colNum);
			return isCellError(p) ? p : indexRef(ref, p);
		}
		if (area !== 1) {
			return ERRORS.REF;
		}
		if (isCellError(src)) {
			return src;
		}
		if (isLambdaValue(src) || isMissingNode(src)) {
			return ERRORS.VALUE;
		}
		const a: ArrayValue = isArrayValue(src)
			? src
			: { kind: 'array', rows: 1, cols: 1, values: [[src as Scalar]] };
		if (a.rows === 0 || a.cols === 0) {
			return ERRORS.REF;
		}
		const p = pick(a.rows, a.cols, rowNum, colNum);
		return isCellError(p) ? p : indexArray(a, p);
	},
};

export const INDEX_FUNCTIONS: FunctionDefinition[] = [INDEX];
