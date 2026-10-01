import { cellKey } from '../address/cell-address';
import { StylePool } from '../styles/style-pool';
import { numberFormatCode } from '../styles/builtin-number-formats';
import type { Cell, CellValue } from '../types/cell';
import type { Workbook } from '../types/workbook';
import { cellError, parseErrorCode } from '../values/errors';
import { parseNumericText } from '../values/parse-number';
import type { EditTransaction } from './transaction';

/** How typed text is interpreted (Excel's data-entry rules). */
export interface ParsedInput {
	kind: 'blank' | 'formula' | 'number' | 'boolean' | 'error' | 'text';
	value: CellValue;
	/** Formula without the leading `=`. */
	formula?: string;
	/** Number format implied by the entry, applied only when the cell's format is General. */
	numFmt?: string;
	/** The entry started with an apostrophe (forced text). */
	quotePrefix?: boolean;
}

export interface ParseInputOptions {
	date1904?: boolean;
	/** Current number format of the cell. */
	currentFormat?: string;
}

/**
 * Interpret text typed into a cell: `'text`, `=formula`, `+A1` / `-A1`
 * formulas, TRUE / FALSE, error literals, numbers with formats (`1,234`,
 * `$5`, `50%`, dates, times, fractions) and plain text.
 */
export function parseCellInput(text: string, opts: ParseInputOptions = {}): ParsedInput {
	if (text === '') {
		return { kind: 'blank', value: null };
	}
	if (text.startsWith("'")) {
		return { kind: 'text', value: text.slice(1), quotePrefix: true };
	}
	const textFormat = opts.currentFormat === '@';
	if (textFormat) {
		return { kind: 'text', value: text };
	}
	if (text.startsWith('=') && text.length > 1) {
		return { kind: 'formula', value: null, formula: text.slice(1) };
	}
	const parsed = parseNumericText(text, { date1904: opts.date1904 });
	if (parsed) {
		let value = parsed.value;
		const fmt = opts.currentFormat ?? 'General';
		// Automatic percent entry: typing 5 into a percent-formatted cell means 5%.
		if (parsed.kind === 'number' && fmt.includes('%') && !/^general$/i.test(fmt)) {
			value /= 100;
		}
		return { kind: 'number', value, numFmt: parsed.numFmt };
	}
	if ((text.startsWith('+') || text.startsWith('-')) && /^[+-]\s*[A-Za-z$(_]/.test(text)) {
		return { kind: 'formula', value: null, formula: text };
	}
	const upper = text.trim().toUpperCase();
	if (upper === 'TRUE' || upper === 'FALSE') {
		return { kind: 'boolean', value: upper === 'TRUE' };
	}
	const err = parseErrorCode(text);
	if (err) {
		return { kind: 'error', value: cellError(err) };
	}
	return { kind: 'text', value: text };
}

export interface SetInputResult {
	/** The input was rejected (e.g. editing part of an array). */
	error?: 'array-part';
	parsed?: ParsedInput;
}

/**
 * Apply typed text to a cell inside a transaction: stores the value or
 * formula, applies the implied number format (only over General), keeps or
 * clears the quote prefix and turns on wrapping for multi-line text.
 * Formulas are stored as dynamic-array formulas; the calc engine computes them.
 */
export function setCellInput(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	row: number,
	col: number,
	text: string,
): SetInputResult {
	const sheet = workbook.sheets[sheetIndex];
	const key = cellKey(row, col);
	const existing = sheet.cells.get(key);
	if (
		existing?.arrayAnchor !== undefined &&
		existing.arrayAnchor !== key &&
		!isDynamicSpillMember(sheet.cells, existing)
	) {
		return { error: 'array-part' };
	}
	if (existing?.array && !existing.array.dynamic) {
		const r = existing.array.ref;
		if (r.top !== r.bottom || r.left !== r.right) {
			return { error: 'array-part' };
		}
	}
	const pool = new StylePool(workbook.styles);
	const currentFormat = numberFormatCode(
		workbook.styles.cellXfs[existing?.style ?? 0]?.numFmtId ?? 0,
		workbook.styles.numFmts,
	);
	const parsed = parseCellInput(text, { date1904: workbook.date1904, currentFormat });
	tx.touchCell(sheetIndex, row, col);

	let style = existing?.style;
	const xf = workbook.styles.cellXfs[style ?? 0];
	if (parsed.numFmt && /^general$/i.test(currentFormat)) {
		style = pool.derive(style, { numFmt: parsed.numFmt });
	}
	if (parsed.quotePrefix && !xf?.quotePrefix) {
		style = pool.derive(style, { quotePrefix: true });
	} else if (!parsed.quotePrefix && xf?.quotePrefix) {
		style = pool.derive(style, { quotePrefix: false });
	}
	if (typeof parsed.value === 'string' && parsed.value.includes('\n') && !xf?.alignment?.wrapText) {
		style = pool.derive(style, { alignment: { wrapText: true } });
	}

	if (parsed.kind === 'blank') {
		if (style === undefined || style === 0) {
			sheet.cells.delete(key);
		} else {
			sheet.cells.set(key, { value: null, style });
		}
	} else {
		const cell: Cell = { value: parsed.value };
		if (parsed.formula !== undefined) {
			cell.formula = parsed.formula;
			cell.array = { ref: { top: row, left: col, bottom: row, right: col }, dynamic: true };
		}
		if (style !== undefined && style !== 0) {
			cell.style = style;
		}
		sheet.cells.set(key, cell);
	}
	sheet.dirty = true;
	return { parsed };
}

function isDynamicSpillMember(cells: Map<number, Cell>, member: Cell): boolean {
	const anchor = member.arrayAnchor !== undefined ? cells.get(member.arrayAnchor) : undefined;
	return Boolean(anchor?.array?.dynamic);
}
