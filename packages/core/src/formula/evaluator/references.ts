import { MAX_COL_INDEX, MAX_COLS, MAX_ROW_INDEX, MAX_ROWS } from '../../address/constants';
import type { RangeAddress } from '../../types/address';
import type { CellError } from '../../types/cell';
import { ERRORS, isCellError } from '../../values/errors';
import type { CellRefNode, RangeRefNode, RefCorner, SheetQualifier } from '../ast';
import type { FormulaValue, RefValue } from '../types';
import { isMultiRefValue, isRefValue } from '../types';
import type { EvaluationHost } from './host';

/** First and last sheet index of a reference (equal unless 3-D). */
export interface SheetSpan {
	readonly start: number;
	readonly end: number;
}

function wrap(n: number, size: number): number {
	const m = n % size;
	return m < 0 ? m + size : m;
}

/**
 * Row of a corner. Inside a defined name, relative parts are relative to
 * A1 and shift by the formula cell (wrapping around the grid like Excel).
 */
function cornerRow(c: RefCorner, inName: boolean, baseRow: number): number {
	return inName && !c.rowAbs ? wrap(c.row + baseRow, MAX_ROWS) : c.row;
}

function cornerCol(c: RefCorner, inName: boolean, baseCol: number): number {
	return inName && !c.colAbs ? wrap(c.col + baseCol, MAX_COLS) : c.col;
}

/** Resolve a sheet qualifier to sheet indices (`#REF!` for unknown sheets and external books). */
export function resolveSheetSpan(
	q: SheetQualifier | undefined,
	current: number,
	host: EvaluationHost,
): SheetSpan | CellError {
	if (q === undefined) {
		return { start: current, end: current };
	}
	if (q.workbook !== undefined) {
		return ERRORS.REF;
	}
	const a = host.sheetIndex(q.sheet);
	if (a < 0) {
		return ERRORS.REF;
	}
	if (q.sheetEnd === undefined) {
		return { start: a, end: a };
	}
	const b = host.sheetIndex(q.sheetEnd);
	if (b < 0) {
		return ERRORS.REF;
	}
	return { start: Math.min(a, b), end: Math.max(a, b) };
}

/** Build a reference value over a sheet span. */
export function makeRef(span: SheetSpan, range: RangeAddress): RefValue {
	return span.end !== span.start
		? { kind: 'ref', sheet: span.start, sheetEnd: span.end, range }
		: { kind: 'ref', sheet: span.start, range };
}

/** Range addressed by a cell node (`A1`). */
export function cellNodeRange(
	node: CellRefNode,
	inName: boolean,
	baseRow: number,
	baseCol: number,
): RangeAddress {
	const row = cornerRow(node.ref, inName, baseRow);
	const col = cornerCol(node.ref, inName, baseCol);
	return { top: row, left: col, bottom: row, right: col };
}

/** Normalised range addressed by a range node (`A1:B2`, `A:C`, `1:3`). */
export function rangeNodeRange(
	node: RangeRefNode,
	inName: boolean,
	baseRow: number,
	baseCol: number,
): RangeAddress {
	const { start, end } = node;
	let top = 0;
	let bottom = MAX_ROW_INDEX;
	let left = 0;
	let right = MAX_COL_INDEX;
	if (node.kind !== 'columns') {
		const r1 = cornerRow(start, inName, baseRow);
		const r2 = cornerRow(end, inName, baseRow);
		top = Math.min(r1, r2);
		bottom = Math.max(r1, r2);
	}
	if (node.kind !== 'rows') {
		const c1 = cornerCol(start, inName, baseCol);
		const c2 = cornerCol(end, inName, baseCol);
		left = Math.min(c1, c2);
		right = Math.max(c1, c2);
	}
	return { top, left, bottom, right };
}

/** The areas of a reference value, or `null` when it is not a reference. */
function areasOf(v: FormulaValue): readonly RefValue[] | null {
	if (isRefValue(v)) {
		return [v];
	}
	if (isMultiRefValue(v)) {
		return v.refs;
	}
	return null;
}

function sameSheets(a: RefValue, b: RefValue): boolean {
	return a.sheet === b.sheet && (a.sheetEnd ?? a.sheet) === (b.sheetEnd ?? b.sheet);
}

/** `a:b`: the smallest area covering both references (same sheet only). */
export function rangeOperator(a: FormulaValue, b: FormulaValue): FormulaValue {
	if (isCellError(a)) {
		return a;
	}
	if (isCellError(b)) {
		return b;
	}
	const areas = [...(areasOf(a) ?? []), ...(areasOf(b) ?? [])];
	if (areasOf(a) === null || areasOf(b) === null || areas.length === 0) {
		return ERRORS.VALUE;
	}
	const first = areas[0];
	let { top, left, bottom, right } = first.range;
	for (const ref of areas) {
		if (!sameSheets(first, ref)) {
			return ERRORS.VALUE;
		}
		top = Math.min(top, ref.range.top);
		left = Math.min(left, ref.range.left);
		bottom = Math.max(bottom, ref.range.bottom);
		right = Math.max(right, ref.range.right);
	}
	const span = { start: first.sheet, end: first.sheetEnd ?? first.sheet };
	return makeRef(span, { top, left, bottom, right });
}

/** `(a,b)`: a multi-area reference (all areas on one sheet). */
export function unionOperator(a: FormulaValue, b: FormulaValue): FormulaValue {
	if (isCellError(a)) {
		return a;
	}
	if (isCellError(b)) {
		return b;
	}
	const left = areasOf(a);
	const right = areasOf(b);
	if (left === null || right === null) {
		return ERRORS.VALUE;
	}
	const refs = [...left, ...right];
	for (const r of refs) {
		if (!sameSheets(refs[0], r) || r.sheetEnd !== undefined) {
			return ERRORS.VALUE;
		}
	}
	return refs.length === 1 ? refs[0] : { kind: 'refs', refs };
}

/** `a b`: the cells common to both references; `#NULL!` when they do not overlap. */
export function intersectionOperator(a: FormulaValue, b: FormulaValue): FormulaValue {
	if (isCellError(a)) {
		return a;
	}
	if (isCellError(b)) {
		return b;
	}
	const left = areasOf(a);
	const right = areasOf(b);
	if (left === null || right === null) {
		return ERRORS.VALUE;
	}
	const out: RefValue[] = [];
	for (const x of left) {
		for (const y of right) {
			if (!sameSheets(x, y)) {
				continue;
			}
			const top = Math.max(x.range.top, y.range.top);
			const leftCol = Math.max(x.range.left, y.range.left);
			const bottom = Math.min(x.range.bottom, y.range.bottom);
			const rightCol = Math.min(x.range.right, y.range.right);
			if (top <= bottom && leftCol <= rightCol) {
				const span = { start: x.sheet, end: x.sheetEnd ?? x.sheet };
				out.push(makeRef(span, { top, left: leftCol, bottom, right: rightCol }));
			}
		}
	}
	if (out.length === 0) {
		return ERRORS.NULL;
	}
	return out.length === 1 ? out[0] : { kind: 'refs', refs: out };
}

/**
 * Implicit intersection that keeps the result a reference: a single cell
 * stays itself, a single column picks the formula row, a single row picks
 * the formula column; anything else is `#VALUE!`.
 */
export function intersectRef(ref: RefValue, row: number, col: number): RefValue | CellError {
	if (ref.sheetEnd !== undefined && ref.sheetEnd !== ref.sheet) {
		return ERRORS.VALUE;
	}
	const r = ref.range;
	if (r.top === r.bottom && r.left === r.right) {
		return ref;
	}
	if (r.left === r.right && row >= r.top && row <= r.bottom) {
		return {
			kind: 'ref',
			sheet: ref.sheet,
			range: { top: row, left: r.left, bottom: row, right: r.left },
		};
	}
	if (r.top === r.bottom && col >= r.left && col <= r.right) {
		return {
			kind: 'ref',
			sheet: ref.sheet,
			range: { top: r.top, left: col, bottom: r.top, right: col },
		};
	}
	return ERRORS.VALUE;
}
