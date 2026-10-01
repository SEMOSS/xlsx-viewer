import { cellKey } from '../address/cell-address';
import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../address/constants';
import { isSingleCell, rangesIntersect } from '../address/range';
import { formatValue } from '../numfmt/format-value';
import { numberFormatCode } from '../styles/builtin-number-formats';
import { StylePool } from '../styles/style-pool';
import type { RangeAddress } from '../types/address';
import type { Cell } from '../types/cell';
import type { Worksheet } from '../types/sheet';
import type { Workbook } from '../types/workbook';
import { setCellInput } from './cell-input';
import { trimRangeToUsed } from './remove-duplicates';
import { splitTextLine, type TextSplitOptions } from './text-split';
import type { EditTransaction } from './transaction';

export type { DelimitedSplitOptions, TextSplitOptions } from './text-split';
export {
	guessFixedWidthBreaks,
	normalizeBreaks,
	parseTextToColumns,
	splitDelimitedText,
	splitFixedWidthText,
	splitTextLine,
} from './text-split';

/** Column data format of the wizard's last step. */
export type TextColumnFormat = 'general' | 'text' | 'skip';

export interface TextToColumnsOptions extends TextSplitOptions {
	/** Cell receiving the first field of the first source row. */
	destination: { row: number; col: number };
	/** Format per parsed column (missing entries are General). */
	columnFormats?: TextColumnFormat[];
}

export type TextToColumnsErrorReason = 'multipleColumns' | 'merge' | 'array';

/** Excel's messages for each refusal. */
export const TEXT_TO_COLUMNS_MESSAGES: Readonly<Record<TextToColumnsErrorReason, string>> = {
	multipleColumns:
		'Microsoft Excel can convert only one column at a time. The range can be many rows tall but ' +
		'no more than one column wide. Try again by selecting cells in one column only.',
	merge: "We can't do that to a merged cell.",
	array: "You can't change part of an array.",
};

export class TextToColumnsError extends Error {
	constructor(
		message: string,
		readonly reason: TextToColumnsErrorReason,
	) {
		super(message);
		this.name = 'TextToColumnsError';
	}
}

function fail(reason: TextToColumnsErrorReason): TextToColumnsError {
	return new TextToColumnsError(TEXT_TO_COLUMNS_MESSAGES[reason], reason);
}

/** The text a cell contributes: strings as stored, other values as displayed. */
function cellText(workbook: Workbook, cell: Cell | undefined): string {
	if (!cell || cell.value === null) {
		return '';
	}
	if (typeof cell.value === 'string') {
		return cell.value;
	}
	const xf = workbook.styles.cellXfs[cell.style ?? 0];
	const code = numberFormatCode(xf?.numFmtId ?? 0, workbook.styles.numFmts);
	return formatValue(cell.value, code, { date1904: workbook.date1904 }).text;
}

/**
 * One line of text per row of the source column (`''` for empty cells); a
 * whole-column source covers the used rows only.
 */
export function textToColumnsLines(
	workbook: Workbook,
	sheetIndex: number,
	source: RangeAddress,
): string[] {
	const sheet = workbook.sheets[sheetIndex];
	const area = trimRangeToUsed(sheet, source);
	const lines: string[] = [];
	for (let row = area.top; row <= area.bottom; row++) {
		lines.push(cellText(workbook, sheet.cells.get(cellKey(row, area.left))));
	}
	return lines;
}

interface PlannedCell {
	row: number;
	col: number;
	text: string;
	format: 'general' | 'text';
}

interface Plan {
	cells: PlannedCell[];
	columns: number;
	source: RangeAddress;
}

function plan(
	workbook: Workbook,
	sheetIndex: number,
	source: RangeAddress,
	options: TextToColumnsOptions,
): Plan {
	const area = trimRangeToUsed(workbook.sheets[sheetIndex], { ...source, right: source.left });
	const formats = options.columnFormats ?? [];
	const { row: top, col: left } = options.destination;
	const cells: PlannedCell[] = [];
	let columns = 0;
	textToColumnsLines(workbook, sheetIndex, area).forEach((line, i) => {
		const row = top + i;
		if (line === '' || row > MAX_ROW_INDEX) {
			return;
		}
		let written = 0;
		splitTextLine(line, options).forEach((text, j) => {
			const format = formats[j] ?? 'general';
			if (format === 'skip') {
				return;
			}
			const col = left + written++;
			if (col <= MAX_COL_INDEX) {
				cells.push({ row, col, text, format });
			}
		});
		columns = Math.max(columns, Math.min(written, MAX_COL_INDEX - left + 1));
	});
	return { cells, columns, source: area };
}

function hasContent(cell: Cell | undefined): boolean {
	return (
		cell !== undefined && (cell.formula !== undefined || (cell.value !== null && cell.value !== ''))
	);
}

/**
 * The conversion would replace data in cells other than the source cells
 * (Excel then asks "There's already data here. Do you want to replace it?").
 */
export function textToColumnsOverwrites(
	workbook: Workbook,
	sheetIndex: number,
	source: RangeAddress,
	options: TextToColumnsOptions,
): boolean {
	const sheet = workbook.sheets[sheetIndex];
	const p = plan(workbook, sheetIndex, source, options);
	return p.cells.some((c) => {
		const inSource = c.col === p.source.left && c.row >= p.source.top && c.row <= p.source.bottom;
		return !inSource && hasContent(sheet.cells.get(cellKey(c.row, c.col)));
	});
}

/** Part of a legacy (Ctrl+Shift+Enter) array formula spanning several cells. */
function inLegacyArray(sheet: Worksheet, row: number, col: number): boolean {
	const cell = sheet.cells.get(cellKey(row, col));
	if (cell?.arrayAnchor !== undefined) {
		return sheet.cells.get(cell.arrayAnchor)?.array?.dynamic !== true;
	}
	return cell?.array !== undefined && !cell.array.dynamic && !isSingleCell(cell.array.ref);
}

function writeText(
	tx: EditTransaction,
	workbook: Workbook,
	pool: StylePool,
	sheetIndex: number,
	c: PlannedCell,
): void {
	const sheet = workbook.sheets[sheetIndex];
	const key = cellKey(c.row, c.col);
	tx.touchCell(sheetIndex, c.row, c.col);
	let style = pool.derive(sheet.cells.get(key)?.style, { numFmt: '@' });
	if (workbook.styles.cellXfs[style]?.quotePrefix) {
		style = pool.derive(style, { quotePrefix: false });
	}
	sheet.cells.set(key, { value: c.text === '' ? null : c.text, style });
}

/**
 * Data > Text to Columns: split the text of a one-column `source` into
 * fields and write them from `options.destination` rightwards, one row per
 * source row. General fields are entered like typed input (numbers, dates
 * and percentages are recognised), Text fields are stored as text with the
 * `@` format, skipped columns are not written and empty source cells are
 * left alone. Throws `TextToColumnsError` (before changing anything) for a
 * multi-column source, merged cells or legacy arrays in the way.
 */
export function textToColumns(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	source: RangeAddress,
	options: TextToColumnsOptions,
): { columns: number } {
	if (source.left !== source.right) {
		throw fail('multipleColumns');
	}
	const sheet = workbook.sheets[sheetIndex];
	const p = plan(workbook, sheetIndex, source, options);
	if (p.cells.length === 0) {
		return { columns: 0 };
	}
	const box: RangeAddress = {
		top: p.cells[0].row,
		bottom: p.cells[p.cells.length - 1].row,
		left: options.destination.col,
		right: options.destination.col + Math.max(0, p.columns - 1),
	};
	if (sheet.merges.some((m) => rangesIntersect(m, box))) {
		throw fail('merge');
	}
	if (p.cells.some((c) => inLegacyArray(sheet, c.row, c.col))) {
		throw fail('array');
	}
	const pool = new StylePool(workbook.styles);
	for (const c of p.cells) {
		if (c.format === 'general') {
			setCellInput(tx, workbook, sheetIndex, c.row, c.col, c.text);
		} else {
			writeText(tx, workbook, pool, sheetIndex, c);
		}
	}
	sheet.dirty = true;
	return { columns: p.columns };
}
