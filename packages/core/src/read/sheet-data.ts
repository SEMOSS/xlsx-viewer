import { a1, cellKey } from '../address/cell-address';
import { MAX_ROWS } from '../address/constants';
import type { Cell, CellValue, RichTextRun } from '../types/cell';
import type { RowInfo } from '../types/sheet';
import { cellError } from '../values/errors';
import {
	T_B,
	T_D,
	T_E,
	T_INLINE,
	T_N,
	T_S,
	T_STR,
	isNameEnd,
	isTrue,
	newCellTag,
	parseUint,
	rowInfo,
	scanCellTag,
	scanRowTag,
	skipElement,
	skipSpecial,
	type CellTag,
} from './cell-scan';
import { errorValue, isoToSerial } from './cell-values';
import type { SharedStrings } from './context';
import { decodeCellText, readStringItem } from './rich-text';
import { FormulaCollector } from './sheet-formulas';
import { parseTagAttributes, tagEnd } from './xml-normalize';

/**
 * Fast `<sheetData>` reader. Worksheets can hold millions of cells, so rows
 * and cells are read with an `indexOf`-driven scanner over the document text
 * instead of building a DOM; only rare constructs (formulas, inline rich
 * text) go through the generic helpers.
 */

export interface SheetDataOptions {
	/** Prefix bound to the SpreadsheetML namespace (`''` for the default namespace). */
	prefix: string;
	sharedStrings: SharedStrings;
	date1904: boolean;
	/** Dynamic-array `cm` indices from `xl/metadata.xml`. */
	dynamicArrayCm?: ReadonlySet<number>;
}

export interface SheetDataIssue {
	code: string;
	message: string;
	cell?: string;
}

export interface SheetDataResult {
	cells: Map<number, Cell>;
	rows: Map<number, RowInfo>;
	/** Cells holding what-if data-table formulas (only their cached values are kept). */
	dataTableCells: number;
	issues: SheetDataIssue[];
}

/** Offsets of the children of one `<c>` (reset per cell). */
interface CellContent {
	vStart: number;
	vEnd: number;
	isStart: number;
	isContent: number;
	isContentEnd: number;
	isEnd: number;
	fAttrs: Record<string, string> | undefined;
	fText: string | undefined;
}

class SheetDataScanner {
	readonly cells = new Map<number, Cell>();
	readonly rows = new Map<number, RowInfo>();
	readonly issues: SheetDataIssue[] = [];
	private readonly p: string;
	private readonly pLen: number;
	private readonly cClose: string;
	private readonly vClose: string;
	private readonly fClose: string;
	private readonly isClose: string;
	private readonly formulas: FormulaCollector;
	private readonly tag: CellTag = newCellTag();
	private readonly content: CellContent = {
		vStart: -1,
		vEnd: -1,
		isStart: -1,
		isContent: -1,
		isContentEnd: -1,
		isEnd: -1,
		fAttrs: undefined,
		fText: undefined,
	};
	private row = -1;
	private col = -1;
	private badStringIndex = false;
	/** Rich runs of the last string value produced by {@link cellValue}. */
	private lastRich: RichTextRun[] | undefined;

	constructor(
		private readonly xml: string,
		private readonly start: number,
		private readonly end: number,
		private readonly options: SheetDataOptions,
	) {
		this.p = options.prefix ? `${options.prefix}:` : '';
		this.pLen = this.p.length;
		this.cClose = `</${this.p}c>`;
		this.vClose = `</${this.p}v>`;
		this.fClose = `</${this.p}f>`;
		this.isClose = `</${this.p}is>`;
		this.formulas = new FormulaCollector(this.cells, options.dynamicArrayCm);
	}

	run(): SheetDataResult {
		const { xml, end } = this;
		let pos = this.start;
		while (pos < end) {
			const lt = xml.indexOf('<', pos);
			if (lt === -1 || lt >= end) {
				break;
			}
			const c1 = xml.charCodeAt(lt + 1);
			if (c1 === 47 /* / */) {
				const gt = xml.indexOf('>', lt + 2);
				pos = gt === -1 ? end : gt + 1;
				continue;
			}
			if (c1 === 33 /* ! */ || c1 === 63 /* ? */) {
				pos = skipSpecial(xml, lt, end);
				continue;
			}
			let i = lt + 1;
			if (this.pLen !== 0) {
				if (!xml.startsWith(this.p, i)) {
					pos = skipElement(xml, lt, end);
					continue;
				}
				i += this.pLen;
			}
			const n0 = xml.charCodeAt(i);
			if (n0 === 99 /* c */ && isNameEnd(xml.charCodeAt(i + 1))) {
				pos = this.readCell(i + 1);
			} else if (
				n0 === 114 /* r */ &&
				xml.charCodeAt(i + 1) === 111 &&
				xml.charCodeAt(i + 2) === 119 &&
				isNameEnd(xml.charCodeAt(i + 3))
			) {
				pos = this.readRow(lt, i + 3);
			} else {
				pos = skipElement(xml, lt, end);
			}
		}
		this.issues.push(...this.formulas.finish());
		return {
			cells: this.cells,
			rows: this.rows,
			dataTableCells: this.formulas.dataTableCells,
			issues: this.issues,
		};
	}

	private readRow(lt: number, attrStart: number): number {
		const { attrs, end } = scanRowTag(this.xml, lt, attrStart);
		const r = attrs.r !== undefined ? Number(attrs.r) : NaN;
		this.row = Number.isInteger(r) && r >= 1 && r <= MAX_ROWS ? r - 1 : this.row + 1;
		this.col = -1;
		const info = rowInfo(attrs);
		if (info) {
			this.rows.set(this.row, info);
		}
		return end;
	}

	/** Read one `<c>`; `k` points just past the element name. Returns the offset after the cell. */
	private readCell(k: number): number {
		const tag = this.tag;
		k = scanCellTag(this.xml, k, this.end, tag);
		let row: number;
		let col: number;
		if (tag.col >= 0) {
			row = this.row = tag.row;
			col = tag.col;
		} else {
			row = this.row < 0 ? 0 : this.row;
			col = this.col + 1;
		}
		this.col = col;

		const content = this.content;
		content.vStart = content.vEnd = content.isStart = -1;
		content.fAttrs = content.fText = undefined;
		if (!tag.selfClosing) {
			k = this.readContent(k);
		}

		let value: CellValue = null;
		let richText: RichTextRun[] | undefined;
		if (content.vStart >= 0 || content.isStart >= 0) {
			value = this.cellValue(tag.type, row, col) ?? null;
			richText = this.lastRich;
		}

		const cell: Cell = { value };
		if (tag.style !== 0) {
			cell.style = tag.style;
		}
		if (richText) {
			cell.richText = richText;
		}
		if (tag.cm >= 0) {
			cell.cm = tag.cm;
		}
		if (tag.vm >= 0) {
			cell.vm = tag.vm;
		}
		if (tag.ph) {
			cell.ph = true;
		}
		const key = cellKey(row, col);
		if (content.fText !== undefined) {
			this.formulas.add(cell, key, row, col, content.fAttrs, content.fText, tag.cm);
		}
		this.cells.set(key, cell);
		return k;
	}

	/** Locate `<v>`, `<f>` and `<is>` inside a cell; returns the offset after `</c>`. */
	private readContent(k: number): number {
		const xml = this.xml;
		const content = this.content;
		const cEnd = xml.indexOf(this.cClose, k);
		if (cEnd === -1) {
			throw new Error('Unterminated <c> element');
		}
		let j = k;
		while (j < cEnd) {
			const lt = xml.indexOf('<', j);
			if (lt === -1 || lt >= cEnd) {
				break;
			}
			let q = lt + 1;
			const first = xml.charCodeAt(q);
			if (first === 47 || first === 33 || first === 63) {
				j = first === 47 ? xml.indexOf('>', q) + 1 : skipSpecial(xml, lt, cEnd);
				continue;
			}
			if (this.pLen !== 0) {
				if (!xml.startsWith(this.p, q)) {
					j = skipElement(xml, lt, cEnd);
					continue;
				}
				q += this.pLen;
			}
			const n0 = xml.charCodeAt(q);
			const n1 = xml.charCodeAt(q + 1);
			if (n0 === 118 /* v */ && isNameEnd(n1)) {
				const gt = xml.indexOf('>', q);
				if (xml.charCodeAt(gt - 1) === 47) {
					content.vStart = content.vEnd = j = gt + 1;
				} else {
					const close = xml.indexOf(this.vClose, gt);
					if (close === -1) {
						throw new Error('Unterminated <v> element');
					}
					content.vStart = gt + 1;
					content.vEnd = close;
					j = close + this.vClose.length;
				}
			} else if (n0 === 102 /* f */ && isNameEnd(n1)) {
				const gt = tagEnd(xml, lt);
				const selfClosing = xml.charCodeAt(gt - 2) === 47;
				const attrEnd = selfClosing ? gt - 2 : gt - 1;
				content.fAttrs =
					attrEnd > q + 1 ? parseTagAttributes(xml.slice(q + 1, attrEnd)) : undefined;
				if (selfClosing) {
					content.fText = '';
					j = gt;
				} else {
					const close = xml.indexOf(this.fClose, gt);
					if (close === -1) {
						throw new Error('Unterminated <f> element');
					}
					content.fText = xml.slice(gt, close);
					j = close + this.fClose.length;
				}
			} else if (n0 === 105 /* i */ && n1 === 115 /* s */ && isNameEnd(xml.charCodeAt(q + 2))) {
				const gt = tagEnd(xml, lt);
				content.isStart = lt;
				if (xml.charCodeAt(gt - 2) === 47) {
					content.isContent = content.isContentEnd = content.isEnd = j = gt;
				} else {
					const close = xml.indexOf(this.isClose, gt);
					if (close === -1) {
						throw new Error('Unterminated <is> element');
					}
					content.isContent = gt;
					content.isContentEnd = close;
					content.isEnd = j = close + this.isClose.length;
				}
			} else {
				j = skipElement(xml, lt, cEnd);
			}
		}
		return cEnd + this.cClose.length;
	}

	private cellValue(type: number, row: number, col: number): CellValue | undefined {
		const xml = this.xml;
		const { vStart, vEnd } = this.content;
		this.lastRich = undefined;
		switch (type) {
			case T_N:
				return this.numberValue(vStart, vEnd);
			case T_S: {
				const idx = vStart >= 0 ? parseUint(xml, vStart, vEnd) : -1;
				const s = idx >= 0 ? this.options.sharedStrings.strings[idx] : undefined;
				if (s === undefined) {
					this.badSharedString(row, col);
					return '';
				}
				this.lastRich = this.options.sharedStrings.rich.get(idx);
				return s;
			}
			case T_STR:
				return vStart >= 0 ? decodeCellText(xml.slice(vStart, vEnd)) : undefined;
			case T_B:
				return vStart >= 0 ? isTrue(xml, vStart, vEnd) : undefined;
			case T_E:
				return vStart >= 0 ? errorValue(decodeCellText(xml.slice(vStart, vEnd))) : undefined;
			case T_INLINE: {
				const c = this.content;
				if (c.isStart < 0) {
					return decodeCellText(xml.slice(vStart, vEnd));
				}
				const item = readStringItem(
					xml,
					c.isStart,
					c.isContent,
					c.isContentEnd,
					c.isEnd,
					this.options.prefix,
				);
				this.lastRich = item.runs;
				return item.text;
			}
			case T_D: {
				if (vStart < 0) {
					return undefined;
				}
				const raw = decodeCellText(xml.slice(vStart, vEnd));
				return isoToSerial(raw, this.options.date1904) ?? raw;
			}
			default:
				return undefined;
		}
	}

	private numberValue(a: number, b: number): CellValue {
		if (a < 0 || b <= a) {
			return null;
		}
		const s = this.xml.slice(a, b);
		const n = +s;
		if (n !== n) {
			// Not a number although the cell is typed as one: keep the text (CDATA unwrapped).
			const text = decodeCellText(s);
			const retry = text.trim() === '' ? NaN : +text;
			return retry === retry ? retry : text;
		}
		if (n === 0 && s.trim() === '') {
			return null;
		}
		return Number.isFinite(n) ? n : cellError('#NUM!');
	}

	private badSharedString(row: number, col: number): void {
		if (!this.badStringIndex) {
			this.badStringIndex = true;
			this.issues.push({
				code: 'invalid-shared-string',
				message: `Cell ${a1(row, col)} refers to a shared string that does not exist`,
				cell: a1(row, col),
			});
		}
	}
}

/** Read the content of `<sheetData>` (`xml.slice(start, end)`). */
export function parseSheetData(
	xml: string,
	start: number,
	end: number,
	options: SheetDataOptions,
): SheetDataResult {
	return new SheetDataScanner(xml, start, end, options).run();
}
