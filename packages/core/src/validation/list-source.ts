import { cellKey } from '../address/cell-address';
import { rangeCellCount, rangeIntersection } from '../address/range';
import { RelativeFormula } from '../conditional/formula-eval';
import { numberToText } from '../formula/coerce';
import type { FormulaValue, RefValue, Scalar } from '../formula/types';
import { isArrayValue, isMultiRefValue, isRefValue } from '../formula/types';
import { formatValue } from '../numfmt';
import { usedRange } from '../sheet/worksheet';
import { numberFormatCode } from '../styles/builtin-number-formats';
import type { RangeAddress } from '../types/address';
import type { Cell, CellValue } from '../types/cell';
import type { DataValidation } from '../types/data-validation';
import type { Workbook } from '../types/workbook';
import { isCellError } from '../values/errors';
import { isDefinedNameReference, resolveReference } from './reference';
import type { DataValidationListEntry, DataValidationContext } from './types';

/** Resolved items of a list validation. */
export interface ListSource {
	/** Non-blank items in source order. */
	entries: DataValidationListEntry[];
	/** The source range contains blank cells. */
	hasBlank: boolean;
	/** The source is a defined name (Excel's "Ignore blank" quirk applies). */
	named: boolean;
}

/** Ranges larger than this are clamped to the sheet's used range before reading. */
const CLAMP_AREA = 65_536;

/** Items of a literal list formula (`"Yes,No,Maybe"`), or `undefined` when it is not one. */
export function parseLiteralList(formula: string): string[] | undefined {
	const t = formula.trim();
	if (t.length < 2 || !t.startsWith('"') || !t.endsWith('"')) {
		return undefined;
	}
	const inner = t.slice(1, -1);
	if (!/^(?:[^"]|"")*$/.test(inner)) {
		return undefined;
	}
	return inner.replace(/""/g, '"').split(',');
}

/** General-format text of a value (numbers without a cell format). */
export function generalText(value: Scalar): string {
	if (typeof value === 'string') {
		return value;
	}
	if (typeof value === 'number') {
		return numberToText(value);
	}
	if (typeof value === 'boolean') {
		return value ? 'TRUE' : 'FALSE';
	}
	return value === null ? '' : value.code;
}

function displayText(
	workbook: Workbook,
	cell: Cell,
	sheetIndex: number,
	row: number,
	col: number,
	ctx: DataValidationContext,
): string {
	const value = cell.value;
	if (ctx.formatValue) {
		return ctx.formatValue(value, sheetIndex, row, col);
	}
	if (typeof value !== 'number') {
		return generalText(value);
	}
	const xf = workbook.styles.cellXfs[cell.style ?? 0];
	const code = numberFormatCode(xf?.numFmtId ?? 0, workbook.styles.numFmts);
	try {
		return formatValue(value, code, { date1904: workbook.date1904 }).text;
	} catch {
		return numberToText(value);
	}
}

function readRange(
	workbook: Workbook,
	sheetIndex: number,
	range: RangeAddress,
	named: boolean,
	ctx: DataValidationContext,
): ListSource {
	const sheet = workbook.sheets[sheetIndex];
	if (!sheet) {
		return { entries: [], hasBlank: true, named };
	}
	let r = range;
	let hasBlank = false;
	const area = rangeCellCount(range);
	if (area > CLAMP_AREA) {
		const used = usedRange(sheet);
		const clamped = used ? rangeIntersection(range, used) : null;
		if (!clamped) {
			return { entries: [], hasBlank: true, named };
		}
		hasBlank = rangeCellCount(clamped) < area;
		r = clamped;
	}
	const entries: DataValidationListEntry[] = [];
	for (let row = r.top; row <= r.bottom; row++) {
		for (let col = r.left; col <= r.right; col++) {
			const cell = sheet.cells.get(cellKey(row, col));
			const value = cell?.value ?? null;
			if (!cell || value === null || value === '') {
				hasBlank = true;
				continue;
			}
			entries.push({ value, text: displayText(workbook, cell, sheetIndex, row, col, ctx) });
		}
	}
	return { entries, hasBlank, named };
}

function fromFormulaResult(
	result: FormulaValue,
	ctx: DataValidationContext,
	named: boolean,
): ListSource {
	const refs: RefValue[] = isRefValue(result)
		? [result]
		: isMultiRefValue(result)
			? result.refs
			: [];
	if (refs.length) {
		const out: ListSource = { entries: [], hasBlank: false, named };
		if (!ctx.workbook) {
			return out;
		}
		for (const ref of refs) {
			const part = readRange(ctx.workbook, ref.sheet, ref.range, named, ctx);
			out.entries.push(...part.entries);
			out.hasBlank ||= part.hasBlank;
		}
		return out;
	}
	// A formula that fails as a whole (or yields a LAMBDA) provides no items.
	const values: Scalar[] = isArrayValue(result)
		? result.values.flat()
		: (typeof result === 'object' && result !== null && 'kind' in result) || isCellError(result)
			? []
			: [result as Scalar];
	const entries: DataValidationListEntry[] = [];
	let hasBlank = false;
	for (const value of values) {
		if (value === null || value === '') {
			hasBlank = true;
		} else {
			entries.push({ value, text: generalText(value) });
		}
	}
	return { entries, hasBlank, named };
}

/**
 * Resolve the items of a list validation: a literal list (`"Yes,No"`), a
 * reference (same sheet, other sheets, whole columns clamped to the used
 * range), a defined name, or any formula returning a reference or an array
 * (evaluated relative to the validated cell). `undefined` when the
 * validation has no source.
 */
export function resolveListSource(
	dv: DataValidation,
	ctx: DataValidationContext,
): ListSource | undefined {
	const source = dv.formula1?.trim().replace(/^=/, '');
	if (!source) {
		return undefined;
	}
	const literal = parseLiteralList(source);
	if (literal) {
		const entries = literal.filter((item) => item !== '').map((text) => ({ text, value: text }));
		return { entries, hasBlank: false, named: false };
	}
	const anchor = dv.ranges[0];
	const formula = new RelativeFormula(source, anchor?.top ?? ctx.row, anchor?.left ?? ctx.col);
	const text = formula.textAt(ctx.row, ctx.col);
	const workbook = ctx.workbook;
	if (workbook) {
		const ref = resolveReference(text, workbook, ctx.sheetIndex);
		if (ref) {
			return readRange(workbook, ref.sheetIndex, ref.range, ref.named, ctx);
		}
	}
	const named = !!workbook && isDefinedNameReference(text, workbook, ctx.sheetIndex);
	const result = formula.evaluateRaw(ctx.evaluate, ctx.sheetIndex, ctx.row, ctx.col);
	return fromFormulaResult(result, ctx, named);
}

/** Dropdown entries (display text and value) of a list validation, in order, blanks skipped. */
export function getListEntries(
	dv: DataValidation,
	ctx: DataValidationContext,
): DataValidationListEntry[] {
	if (dv.type !== 'list') {
		return [];
	}
	return resolveListSource(dv, ctx)?.entries ?? [];
}

/** Dropdown items of a list validation: the displayed values, in order, blanks skipped. */
export function getListItems(dv: DataValidation, ctx: DataValidationContext): string[] {
	return getListEntries(dv, ctx).map((entry) => entry.text);
}

function sameNumber(a: number, b: number): boolean {
	return a === b || Number(a.toPrecision(15)) === Number(b.toPrecision(15));
}

/**
 * Whether a value matches a list entry like Excel: text case-insensitively,
 * numbers numerically (literal items such as `"1,2,3"` and typed numeric text
 * included), booleans by value.
 */
export function listEntryMatches(
	entry: DataValidationListEntry,
	value: CellValue,
	toNumber: (text: string) => number | undefined,
): boolean {
	const ev = entry.value;
	if (value === null) {
		return false;
	}
	if (typeof value === 'number') {
		if (typeof ev === 'number') {
			return sameNumber(ev, value);
		}
		const n = typeof ev === 'string' ? toNumber(ev) : undefined;
		return n !== undefined && sameNumber(n, value);
	}
	if (typeof value === 'string') {
		const lower = value.toLowerCase();
		if (
			entry.text.toLowerCase() === lower ||
			(typeof ev === 'string' && ev.toLowerCase() === lower)
		) {
			return true;
		}
		if (typeof ev === 'number' || typeof ev === 'string') {
			const n = toNumber(value);
			const target = typeof ev === 'number' ? ev : n === undefined ? undefined : toNumber(ev);
			return n !== undefined && target !== undefined && sameNumber(n, target);
		}
		return typeof ev === 'boolean' && lower === (ev ? 'true' : 'false');
	}
	if (typeof value === 'boolean') {
		return typeof ev === 'boolean'
			? ev === value
			: entry.text.toUpperCase() === (value ? 'TRUE' : 'FALSE');
	}
	return isCellError(ev) ? ev.code === value.code : entry.text.toUpperCase() === value.code;
}
