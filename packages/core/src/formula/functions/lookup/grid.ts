import { isWholeColumnRange, isWholeRowRange } from '../../../address/range';
import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import type {
	ArrayValue,
	FormulaValue,
	FunctionArg,
	FunctionContext,
	RefValue,
	Scalar,
} from '../../types';
import { isArrayValue, isLambdaValue, isMultiRefValue, isRefValue } from '../../types';
import { isMissingNode } from './args';
import { buildExactIndex, type ExactIndex } from './exact-index';

/** A one-dimensional run of cells or array elements. */
export interface Vec {
	/** Declared length (a whole-column reference has 1,048,576 elements). */
	readonly length: number;
	/** Elements at this index and beyond are known to be blank (outside the used range). */
	readonly dataLength: number;
	get(i: number): Scalar;
	/** Memoised exact-match index of the data, when the context can cache it (references in a recalculation). */
	exactIndex?: () => ExactIndex;
}

/**
 * Read-only 2-D view of a lookup source. References are read cell by cell on
 * demand instead of being materialised, so a search over a large or
 * whole-column range only touches the cells it actually visits.
 */
export interface Grid {
	/** Declared size. */
	readonly rows: number;
	readonly cols: number;
	/** Rows / columns at or beyond these indices are blank (references outside the used range). */
	readonly dataRows: number;
	readonly dataCols: number;
	/** The reference behind the grid, or null for an in-memory array. */
	readonly ref: RefValue | null;
	get(r: number, c: number): Scalar;
	column(c: number): Vec;
	row(r: number): Vec;
}

/** Grid over an in-memory array. */
export function arrayGrid(a: ArrayValue): Grid {
	const { rows, cols, values } = a;
	return {
		rows,
		cols,
		dataRows: rows,
		dataCols: cols,
		ref: null,
		get: (r, c) => values[r][c],
		column: (c) => ({ length: rows, dataLength: rows, get: (i) => values[i][c] }),
		row: (r) => {
			const line = values[r];
			return { length: cols, dataLength: cols, get: (i) => line[i] };
		},
	};
}

function clamp(n: number, lo: number, hi: number): number {
	return n < lo ? lo : n > hi ? hi : n;
}

type MemoContext = { memo?: <T>(ref: RefValue, key: string, build: () => T) => T };

/** Attach a lazily built, engine-memoised exact-match index to a reference vector. */
function withIndex(vec: Vec, ctx: FunctionContext, ref: RefValue): Vec {
	const host = ctx as unknown as MemoContext;
	if (host.memo === undefined) {
		return vec;
	}
	let local: ExactIndex | undefined;
	vec.exactIndex = (): ExactIndex => {
		local ??= host.memo!<ExactIndex>(ref, 'exact-index', () => buildExactIndex(vec));
		return local;
	};
	return vec;
}

/** Grid over a single-sheet reference; cells are read lazily through the context. */
export function refGrid(ref: RefValue, ctx: FunctionContext): Grid {
	const { sheet } = ref;
	const { top, left, bottom, right } = ref.range;
	const rows = bottom - top + 1;
	const cols = right - left + 1;
	const used = ctx.usedRange(sheet);
	const dataRows = used ? clamp(used.bottom - top + 1, 0, rows) : 0;
	const dataCols = used ? clamp(used.right - left + 1, 0, cols) : 0;
	return {
		rows,
		cols,
		dataRows,
		dataCols,
		ref,
		get: (r, c) => ctx.getCellValue(sheet, top + r, left + c),
		column: (c) => {
			const col = left + c;
			const vec: Vec = {
				length: rows,
				dataLength: c < dataCols ? dataRows : 0,
				get: (i) => ctx.getCellValue(sheet, top + i, col),
			};
			return withIndex(vec, ctx, { ...ref, range: { top, bottom, left: col, right: col } });
		},
		row: (r) => {
			const row = top + r;
			const vec: Vec = {
				length: cols,
				dataLength: r < dataRows ? dataCols : 0,
				get: (i) => ctx.getCellValue(sheet, row, left + i),
			};
			return withIndex(vec, ctx, { ...ref, range: { top: row, bottom: row, left, right } });
		},
	};
}

/**
 * View any lookup argument (`any` parameter) as a grid: references stay
 * lazy, arrays are used directly and scalars become 1x1. Errors pass
 * through; unions of several areas are `#VALUE!` and 3-D references `#REF!`.
 */
export function toGrid(arg: FunctionArg | undefined, ctx: FunctionContext): Grid | CellError {
	if (isRefValue(arg)) {
		if (arg.sheetEnd !== undefined && arg.sheetEnd !== arg.sheet) {
			return ERRORS.REF;
		}
		return refGrid(arg, ctx);
	}
	if (isMultiRefValue(arg)) {
		return arg.refs.length === 1 ? toGrid(arg.refs[0], ctx) : ERRORS.VALUE;
	}
	if (isArrayValue(arg)) {
		return arrayGrid(arg);
	}
	if (isCellError(arg)) {
		return arg;
	}
	if (isLambdaValue(arg)) {
		return ERRORS.CALC;
	}
	if (arg === undefined || isMissingNode(arg)) {
		return ERRORS.VALUE;
	}
	if (arg !== null && typeof arg === 'object') {
		// Any other AST node: never expected for an evaluated argument.
		return ERRORS.VALUE;
	}
	return arrayGrid({ kind: 'array', rows: 1, cols: 1, values: [[arg]] });
}

/** The grid is one row or one column. */
export function isVectorGrid(g: Grid): boolean {
	return g.rows === 1 || g.cols === 1;
}

/** A reference spanning every row of the sheet (`A:A`). */
export function isFullHeight(g: Grid): boolean {
	return g.ref !== null && isWholeColumnRange(g.ref.range);
}

/** A reference spanning every column of the sheet (`1:1`). */
export function isFullWidth(g: Grid): boolean {
	return g.ref !== null && isWholeRowRange(g.ref.range);
}

/**
 * Two extents describe the same number of elements. Arrays computed from a
 * whole-column (or whole-row) reference are clamped to the used range by the
 * evaluator, so such an array is accepted against the full reference when
 * it is no longer than it.
 */
export function sameExtent(
	a: Grid,
	aLength: number,
	b: Grid,
	bLength: number,
	full: (g: Grid) => boolean,
): boolean {
	if (aLength === bLength) {
		return true;
	}
	if (full(a) && b.ref === null) {
		return bLength <= aLength;
	}
	if (full(b) && a.ref === null) {
		return aLength <= bLength;
	}
	return false;
}

/** Row `r` of a grid as a result: a reference into the sheet, or the array row (scalar when 1x1). */
export function gridRowResult(g: Grid, r: number): FormulaValue {
	if (g.ref) {
		const { top, left, right } = g.ref.range;
		return {
			kind: 'ref',
			sheet: g.ref.sheet,
			range: { top: top + r, left, bottom: top + r, right },
		};
	}
	if (r >= g.rows) {
		return null;
	}
	if (g.cols === 1) {
		return g.get(r, 0);
	}
	const line: Scalar[] = new Array(g.cols);
	for (let c = 0; c < g.cols; c++) {
		line[c] = g.get(r, c);
	}
	return { kind: 'array', rows: 1, cols: g.cols, values: [line] };
}

/** Column `c` of a grid as a result: a reference into the sheet, or the array column. */
export function gridColumnResult(g: Grid, c: number): FormulaValue {
	if (g.ref) {
		const { top, left, bottom } = g.ref.range;
		return {
			kind: 'ref',
			sheet: g.ref.sheet,
			range: { top, left: left + c, bottom, right: left + c },
		};
	}
	if (c >= g.cols) {
		return null;
	}
	if (g.rows === 1) {
		return g.get(0, c);
	}
	const values: Scalar[][] = new Array(g.rows);
	for (let r = 0; r < g.rows; r++) {
		values[r] = [g.get(r, c)];
	}
	return { kind: 'array', rows: g.rows, cols: 1, values };
}
