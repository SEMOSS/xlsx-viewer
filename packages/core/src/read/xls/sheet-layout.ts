/**
 * Row and column layout: ROW, COLINFO, DEFCOLWIDTH, STANDARDWIDTH,
 * DEFAULTROWHEIGHT, MERGEDCELLS, WSBOOL (outline summary positions) and
 * SHEETEXT (tab color).
 */
import { MAX_COL_INDEX } from '../../address/constants';
import type { ColInfo, RowInfo } from '../../types/sheet';
import { BiffCursor } from './biff-cursor';
import { RT } from './biff-opcodes';
import type { BiffRecord } from './biff-records';
import { makeRange, type SheetContext } from './sheet-context';
import { icvColor } from './xf-records';

/** Record types handled by {@link readLayout}. */
export const LAYOUT_RECORDS: readonly number[] = [
	RT.ROW,
	RT.COLINFO,
	RT.DEFCOLWIDTH,
	RT.STANDARDWIDTH,
	RT.DEFAULTROWHEIGHT,
	RT.MERGEDCELLS,
	RT.WSBOOL,
	RT.SHEETEXT,
];

function readRow(ctx: SheetContext, cur: BiffCursor): void {
	const row = cur.u16();
	cur.skip(4);
	const height = (cur.u16() & 0x7fff) / 20;
	cur.skip(4);
	const flags = cur.u16();
	const flags2 = cur.u16();
	const info: RowInfo = { height };
	if (flags & 0x40) {
		info.customHeight = true;
	}
	if (flags & 0x20) {
		info.hidden = true;
	}
	if (flags & 0x07) {
		info.outlineLevel = flags & 0x07;
	}
	if (flags & 0x10) {
		info.collapsed = true;
	}
	if (flags & 0x80) {
		const style = ctx.styles.cellXf(flags2 & 0x0fff);
		if (style) {
			info.style = style;
			info.customFormat = true;
		}
	}
	if (flags2 & 0x1000) {
		info.thickTop = true;
	}
	if (flags2 & 0x2000) {
		info.thickBot = true;
	}
	ctx.sheet.rows.set(row, info);
}

function readColInfo(ctx: SheetContext, cur: BiffCursor): void {
	const min = cur.u16();
	const max = Math.min(cur.u16(), MAX_COL_INDEX);
	const width = cur.u16() / 256;
	const style = ctx.styles.cellXf(cur.u16());
	const flags = cur.u16();
	if (max < min) {
		return;
	}
	const col: ColInfo = { min, max, width };
	if (flags & 0x02) {
		col.customWidth = true;
	}
	if (flags & 0x01) {
		col.hidden = true;
	}
	if (style) {
		col.style = style;
	}
	const level = (flags >> 8) & 0x07;
	if (level) {
		col.outlineLevel = level;
	}
	if (flags & 0x1000) {
		col.collapsed = true;
	}
	if (flags & 0x04) {
		col.bestFit = true;
	}
	ctx.sheet.cols.push(col);
}

function readMerges(ctx: SheetContext, cur: BiffCursor): void {
	const count = Math.min(cur.u16(), Math.floor(cur.remaining / 8));
	for (let i = 0; i < count; i++) {
		const range = makeRange(cur.u16(), cur.u16(), cur.u16(), cur.u16());
		if (range && (range.top !== range.bottom || range.left !== range.right)) {
			ctx.sheet.merges.push(range);
		}
	}
}

/** Handle a layout record; returns `false` for other record types. */
export function readLayout(ctx: SheetContext, record: BiffRecord): boolean {
	const cur = new BiffCursor(record.data, record.breaks);
	const sheet = ctx.sheet;
	switch (record.type) {
		case RT.ROW:
			readRow(ctx, cur);
			return true;
		case RT.COLINFO:
			readColInfo(ctx, cur);
			return true;
		case RT.DEFCOLWIDTH:
			sheet.baseColWidth = cur.u16();
			return true;
		case RT.STANDARDWIDTH:
			sheet.defaultColWidth = cur.u16() / 256;
			return true;
		case RT.DEFAULTROWHEIGHT: {
			const flags = cur.u16();
			const height = cur.u16() / 20;
			if (height > 0 && !(flags & 0x02)) {
				sheet.defaultRowHeight = height;
			}
			if (flags & 0x02) {
				sheet.zeroHeight = true;
			}
			return true;
		}
		case RT.MERGEDCELLS:
			readMerges(ctx, cur);
			return true;
		case RT.WSBOOL: {
			const flags = cur.u16();
			if (!(flags & 0x40)) {
				sheet.summaryBelow = false;
			}
			if (!(flags & 0x80)) {
				sheet.summaryRight = false;
			}
			return true;
		}
		case RT.SHEETEXT: {
			cur.skip(16);
			const color = icvColor(cur.u32() & 0x7f);
			if (color) {
				sheet.tabColor = color;
			}
			return true;
		}
		default:
			return false;
	}
}

/**
 * Tidy the layout once every record is read: row heights equal to the
 * default are dropped (unless custom), columns are sorted and made
 * non-overlapping, and the outline levels of the sheet are derived.
 */
export function finishLayout(ctx: SheetContext): void {
	const sheet = ctx.sheet;
	let rowLevel = 0;
	for (const [row, info] of sheet.rows) {
		rowLevel = Math.max(rowLevel, info.outlineLevel ?? 0);
		if (!info.customHeight && Math.abs((info.height ?? 0) - sheet.defaultRowHeight) < 0.001) {
			delete info.height;
		}
		if (Object.keys(info).length === 0) {
			sheet.rows.delete(row);
		}
	}
	sheet.cols.sort((a, b) => a.min - b.min);
	const clean: ColInfo[] = [];
	for (const c of sheet.cols) {
		const prev = clean[clean.length - 1];
		if (prev && c.min <= prev.max) {
			if (c.max <= prev.max) {
				continue;
			}
			c.min = prev.max + 1;
		}
		clean.push(c);
	}
	sheet.cols = clean;
	const colLevel = clean.reduce((m, c) => Math.max(m, c.outlineLevel ?? 0), 0);
	if (rowLevel) {
		sheet.outlineLevelRow = rowLevel;
	}
	if (colLevel) {
		sheet.outlineLevelCol = colLevel;
	}
}
