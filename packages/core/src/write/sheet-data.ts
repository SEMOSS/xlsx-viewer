/**
 * `<sheetData>` serialization. Built as a flat string array (no DOM) because
 * it dominates save time: rows and cells in ascending order, values through
 * the shared string table, formulas with their file prefixes restored.
 */
import { columnToLetters } from '../address/cell-address';
import { MAX_COLS, MAX_ROWS } from '../address/constants';
import { formatRange } from '../address/range';
import type { RangeAddress } from '../types/address';
import type { Cell, CellError } from '../types/cell';
import type { RowInfo, Worksheet } from '../types/sheet';
import { escapeXmlText } from '../xml';
import { fileFormula } from './formulas';
import type { CellMetadata } from './metadata';
import type { SharedStringTable } from './shared-strings';
import { escapeCellText, formatXmlNumber } from './xml-utils';

export interface SheetDataContext {
	sst: SharedStringTable;
	/** Number of `cellXfs` entries; larger style indices are dropped. */
	styleCount: number;
	metadata: CellMetadata;
}

export interface SheetDataResult {
	xml: string;
	/** Bounding range of the written cells (style-only cells included), or `null`. */
	used: RangeAddress | null;
	/** Some formula cell is marked `stale`. */
	stale: boolean;
	/** Highest row outline level written. */
	maxRowOutline: number;
	/** Cells whose style index pointed past `cellXfs` (written with the default style). */
	invalidStyles: number;
	/** Cells with non-finite numbers (written as `#NUM!`). */
	invalidNumbers: number;
}

const letters: string[] = [];

function colName(col: number): string {
	let s = letters[col];
	if (s === undefined) {
		s = columnToLetters(col);
		letters[col] = s;
	}
	return s;
}

/** Excel's maximum row height in points. */
const MAX_ROW_HEIGHT = 409;

function rowAttrs(info: RowInfo, styleCount: number): string {
	let s = '';
	if (info.customFormat && info.style !== undefined && info.style >= 0 && info.style < styleCount) {
		s += ` s="${info.style}" customFormat="1"`;
	}
	if (info.height !== undefined && Number.isFinite(info.height) && info.height >= 0) {
		s += ` ht="${formatXmlNumber(Math.min(info.height, MAX_ROW_HEIGHT))}"`;
	}
	if (info.hidden) {
		s += ' hidden="1"';
	}
	if (info.customHeight && info.height !== undefined) {
		s += ' customHeight="1"';
	}
	if (info.outlineLevel) {
		s += ` outlineLevel="${Math.min(7, Math.max(0, Math.trunc(info.outlineLevel)))}"`;
	}
	if (info.collapsed) {
		s += ' collapsed="1"';
	}
	if (info.thickTop) {
		s += ' thickTop="1"';
	}
	if (info.thickBot) {
		s += ' thickBot="1"';
	}
	return s;
}

/** Formula text as stored in the file: prefixes restored, XML-escaped. */
export function formulaText(formula: string): string {
	return escapeXmlText(fileFormula(formula));
}

interface State {
	stale: boolean;
	invalidStyles: number;
	invalidNumbers: number;
}

function metaAttrs(cell: Cell, ctx: SheetDataContext, dynamic: boolean): string {
	let s = '';
	let cm = cell.cm;
	if (dynamic) {
		cm = ctx.metadata.dynamicArrayCm(cm);
	} else if (cm !== undefined && !ctx.metadata.validCm(cm)) {
		cm = undefined;
	}
	if (cm !== undefined) {
		s += ` cm="${cm}"`;
	}
	if (cell.vm !== undefined && ctx.metadata.validVm(cell.vm)) {
		s += ` vm="${cell.vm}"`;
	}
	if (cell.ph) {
		s += ' ph="1"';
	}
	return s;
}

/** Append one `<c>`; returns false when the cell carries nothing worth writing. */
function writeCell(
	out: string[],
	ref: string,
	cell: Cell,
	ctx: SheetDataContext,
	st: State,
): boolean {
	let head = `<c r="${ref}"`;
	const style = cell.style;
	let styled = false;
	if (style) {
		if (Number.isInteger(style) && style > 0 && style < ctx.styleCount) {
			head += ` s="${style}"`;
			styled = true;
		} else {
			st.invalidStyles++;
		}
	}
	const v = cell.value;
	const formula = cell.formula;
	if (formula && cell.arrayAnchor === undefined) {
		if (cell.stale) {
			st.stale = true;
		}
		let t = '';
		let vx = '';
		if (typeof v === 'number') {
			if (Number.isFinite(v)) {
				vx = `<v>${formatXmlNumber(v)}</v>`;
			} else {
				st.invalidNumbers++;
				t = ' t="e"';
				vx = '<v>#NUM!</v>';
			}
		} else if (typeof v === 'string') {
			t = ' t="str"';
			vx = `<v>${escapeCellText(v)}</v>`;
		} else if (typeof v === 'boolean') {
			t = ' t="b"';
			vx = v ? '<v>1</v>' : '<v>0</v>';
		} else if (v !== null && v !== undefined) {
			t = ' t="e"';
			vx = `<v>${escapeXmlText((v as CellError).code)}</v>`;
		}
		const array = cell.array;
		const f = array
			? `<f t="array" ref="${formatRange(array.ref, { forceCells: true })}">${formulaText(formula)}</f>`
			: `<f>${formulaText(formula)}</f>`;
		out.push(`${head}${t}${metaAttrs(cell, ctx, !!array?.dynamic)}>${f}${vx}</c>`);
		return true;
	}
	if (v === null || v === undefined) {
		const meta = metaAttrs(cell, ctx, false);
		if (!styled && meta === '') {
			return false;
		}
		out.push(`${head}${meta}/>`);
		return true;
	}
	switch (typeof v) {
		case 'number':
			if (Number.isFinite(v)) {
				out.push(`${head}${metaAttrs(cell, ctx, false)}><v>${formatXmlNumber(v)}</v></c>`);
			} else {
				st.invalidNumbers++;
				out.push(`${head} t="e"${metaAttrs(cell, ctx, false)}><v>#NUM!</v></c>`);
			}
			return true;
		case 'string': {
			const idx = cell.richText?.length
				? ctx.sst.internRich(cell.richText)
				: ctx.sst.internPlain(v);
			out.push(`${head} t="s"${metaAttrs(cell, ctx, false)}><v>${idx}</v></c>`);
			return true;
		}
		case 'boolean':
			out.push(`${head} t="b"${metaAttrs(cell, ctx, false)}><v>${v ? 1 : 0}</v></c>`);
			return true;
		default:
			out.push(
				`${head} t="e"${metaAttrs(cell, ctx, false)}><v>${escapeXmlText((v as CellError).code)}</v></c>`,
			);
			return true;
	}
}

function hasRowContent(info: RowInfo | undefined): boolean {
	return (
		!!info &&
		((info.customFormat && info.style !== undefined) ||
			info.height !== undefined ||
			!!info.hidden ||
			!!info.outlineLevel ||
			!!info.collapsed ||
			!!info.thickTop ||
			!!info.thickBot)
	);
}

/** Serialize the cells and row records of a sheet. */
export function buildSheetData(sheet: Worksheet, ctx: SheetDataContext): SheetDataResult {
	const keys = Float64Array.from(sheet.cells.keys()).sort();
	const rowKeys = Float64Array.from(sheet.rows.keys()).sort();
	const out: string[] = ['<sheetData>'];
	const st: State = { stale: false, invalidStyles: 0, invalidNumbers: 0 };
	let top = Infinity;
	let bottom = -1;
	let left = Infinity;
	let right = -1;
	let maxOutline = 0;
	let i = 0;
	let j = 0;
	while (i < keys.length || j < rowKeys.length) {
		const cellRow = i < keys.length ? Math.floor(keys[i] / MAX_COLS) : Infinity;
		const infoRow = j < rowKeys.length ? rowKeys[j] : Infinity;
		const r = cellRow < infoRow ? cellRow : infoRow;
		if (r >= MAX_ROWS || r < 0 || !Number.isInteger(r)) {
			// Out of the grid: skip whatever sits on this row.
			if (infoRow === r) {
				j++;
			}
			while (i < keys.length && Math.floor(keys[i] / MAX_COLS) === r) {
				i++;
			}
			continue;
		}
		let info: RowInfo | undefined;
		if (infoRow === r) {
			info = sheet.rows.get(r);
			j++;
		}
		const attrs = hasRowContent(info) ? rowAttrs(info!, ctx.styleCount) : '';
		if (info?.outlineLevel && info.outlineLevel > maxOutline) {
			maxOutline = Math.min(7, Math.trunc(info.outlineLevel));
		}
		const rowNum = r + 1;
		const start = out.length;
		out.push(`<row r="${rowNum}"${attrs}>`);
		const rowText = String(rowNum);
		let wrote = false;
		const base = r * MAX_COLS;
		while (i < keys.length && keys[i] < base + MAX_COLS) {
			const key = keys[i++];
			const col = key - base;
			if (writeCell(out, colName(col) + rowText, sheet.cells.get(key)!, ctx, st)) {
				wrote = true;
				if (col < left) {
					left = col;
				}
				if (col > right) {
					right = col;
				}
			}
		}
		if (wrote) {
			out.push('</row>');
			if (r < top) {
				top = r;
			}
			bottom = r;
		} else {
			out.length = start;
			if (attrs) {
				out.push(`<row r="${rowNum}"${attrs}/>`);
			}
		}
	}
	out.push('</sheetData>');
	return {
		xml: out.length === 2 ? '<sheetData/>' : out.join(''),
		used: bottom < 0 ? null : { top, left, bottom, right },
		stale: st.stale,
		maxRowOutline: maxOutline,
		invalidStyles: st.invalidStyles,
		invalidNumbers: st.invalidNumbers,
	};
}
