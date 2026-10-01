import { formatCellAddress } from '../../../address/cell-address';
import { quoteSheetName } from '../../../address/sheet-name';
import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionContext, FunctionDefinition, Scalar } from '../../types';
import { isMultiRefValue, isRefValue, makeArray } from '../../types';
import { isMissingArg, textArg } from '../math/args';
import { cellFormatCode, hasNegativeColor, hasParentheses } from './cell-format';

/**
 * Optional capabilities CELL uses when the evaluation context provides them
 * (duck-typed, so hosts can add them without changing `FunctionContext`).
 * Without them CELL reports Excel's defaults: General format, general
 * alignment, locked cells, the standard 8.43 column width, unsaved workbook.
 */
export interface CellInfoProvider {
	/** Number format code of a cell; `undefined` for General. */
	getNumberFormat?(sheet: number, row: number, col: number): string | undefined;
	/** Horizontal alignment (`general`, `left`, `center`, `right`, `fill`, `justify`, ...). */
	getHorizontalAlignment?(sheet: number, row: number, col: number): string | undefined;
	/** Cell protection: locked (Excel's default) or not. */
	isCellLocked?(sheet: number, row: number, col: number): boolean;
	/** Column width in characters and whether it is the sheet default. */
	getColumnWidth?(sheet: number, col: number): { width: number; isDefault: boolean } | undefined;
	/** Workbook file name (`Book1.xlsx`) once saved. */
	getWorkbookName?(): string | undefined;
}

/** Standard column width (Calibri 11). */
const DEFAULT_COLUMN_WIDTH = 8.43;

interface Target {
	sheet: number;
	row: number;
	col: number;
}

function addressText(target: Target, ctx: FunctionContext, info: CellInfoProvider): string {
	const address = formatCellAddress({ row: target.row, col: target.col }, true, true);
	if (target.sheet === ctx.sheet) {
		return address;
	}
	const name = ctx.sheetName(target.sheet);
	const book = info.getWorkbookName?.();
	const qualified = book ? `[${book}]${name}` : name;
	const quoted = quoteSheetName(name) !== name ? `'${qualified.replace(/'/g, "''")}'` : qualified;
	return `${quoted}!${address}`;
}

/** CELL("prefix"): the label prefix implied by the alignment of a text cell. */
function prefix(value: Scalar, alignment: string | undefined): string {
	if (typeof value !== 'string') {
		return '';
	}
	switch ((alignment ?? 'general').toLowerCase()) {
		case 'right':
			return '"';
		case 'center':
		case 'centercontinuous':
			return '^';
		case 'fill':
			return '\\';
		default:
			return "'";
	}
}

function cellInfo(
	type: string,
	target: Target,
	ctx: FunctionContext,
): Scalar | ReturnType<typeof makeArray> {
	const info = ctx as FunctionContext & CellInfoProvider;
	const { sheet, row, col } = target;
	const numFmt = () => info.getNumberFormat?.(sheet, row, col);
	switch (type) {
		case 'address':
			return addressText(target, ctx, info);
		case 'col':
			return col + 1;
		case 'row':
			return row + 1;
		case 'contents': {
			const v = ctx.getCellValue(sheet, row, col);
			return v === null ? 0 : v;
		}
		case 'type': {
			const v = ctx.getCellValue(sheet, row, col);
			return v === null ? 'b' : typeof v === 'string' ? 'l' : 'v';
		}
		case 'filename': {
			const book = info.getWorkbookName?.();
			return book ? `[${book}]${ctx.sheetName(sheet)}` : '';
		}
		case 'format':
			return cellFormatCode(numFmt());
		case 'color':
			return hasNegativeColor(numFmt()) ? 1 : 0;
		case 'parentheses':
			return hasParentheses(numFmt()) ? 1 : 0;
		case 'prefix':
			return prefix(
				ctx.getCellValue(sheet, row, col),
				info.getHorizontalAlignment?.(sheet, row, col),
			);
		case 'protect':
			return info.isCellLocked?.(sheet, row, col) === false ? 0 : 1;
		case 'width': {
			// Excel 365 returns the integer width and whether it is the default width.
			const w = info.getColumnWidth?.(sheet, col);
			return makeArray([[Math.trunc(w?.width ?? DEFAULT_COLUMN_WIDTH), w?.isDefault ?? true]]);
		}
		default:
			return ERRORS.VALUE;
	}
}

const CELL: FunctionDefinition = {
	name: 'CELL',
	minArgs: 1,
	maxArgs: 2,
	params: ['value', 'ref'],
	// Formats and column widths are not tracked as dependencies: volatile, as in Excel.
	volatile: true,
	returnsArray: true,
	impl: (args, ctx) => {
		const type = textArg(args[0]);
		if (isCellError(type)) {
			return type;
		}
		let target: Target = { sheet: ctx.sheet, row: ctx.row, col: ctx.col };
		const ref = args[1];
		if (args.length > 1 && ref !== null && !isMissingArg(ref)) {
			const single = isMultiRefValue(ref) && ref.refs.length === 1 ? ref.refs[0] : ref;
			if (!isRefValue(single)) {
				return isCellError(single) ? single : ERRORS.VALUE;
			}
			target = { sheet: single.sheet, row: single.range.top, col: single.range.left };
		}
		return cellInfo(type.trim().toLowerCase(), target, ctx);
	},
};

export const CELL_FUNCTIONS: FunctionDefinition[] = [CELL];
