/**
 * Value cell records: NUMBER, RK, MULRK, LABELSST, LABEL, RSTRING,
 * BOOLERR, BLANK and MULBLANK. Blank cells are only kept when they carry
 * a non-default format.
 */
import { a1, cellKey } from '../../address/cell-address';
import type { CellValue, RichTextRun } from '../../types/cell';
import { cellError, ERRORS } from '../../values/errors';
import { BiffCursor } from './biff-cursor';
import { RT } from './biff-opcodes';
import type { BiffRecord } from './biff-records';
import { ERROR_CODES } from './formula';
import { makeCell, type SheetContext, sheetWarn, validCell } from './sheet-context';
import { type BiffRun, richTextRuns } from './sst';

/** Record types handled by {@link readValueCell}. */
export const VALUE_CELL_RECORDS: readonly number[] = [
	RT.NUMBER,
	RT.RK,
	RT.MULRK,
	RT.LABELSST,
	RT.LABEL,
	RT.RSTRING,
	RT.BOOLERR,
	RT.BLANK,
	RT.MULBLANK,
];

const scratch = new DataView(new ArrayBuffer(8));

/**
 * RK number: bit 0 divides by 100, bit 1 selects a signed 30-bit integer;
 * otherwise the upper 30 bits are the high bits of an IEEE double.
 */
export function decodeRk(rk: number): number {
	let v: number;
	if (rk & 0x02) {
		v = rk >> 2;
	} else {
		scratch.setUint32(0, 0, true);
		scratch.setUint32(4, rk & 0xfffffffc, true);
		v = scratch.getFloat64(0, true);
	}
	return rk & 0x01 ? v / 100 : v;
}

/** Cell value of a BIFF error code. */
export function errorValue(code: number): CellValue {
	const literal = ERROR_CODES[code];
	return literal ? cellError(literal) : ERRORS.NA;
}

function numberValue(v: number): CellValue {
	return Number.isFinite(v) ? v : ERRORS.NUM;
}

function put(
	ctx: SheetContext,
	row: number,
	col: number,
	value: CellValue,
	ixfe: number,
	richText?: RichTextRun[],
): void {
	if (!validCell(row, col)) {
		sheetWarn(
			ctx,
			'cell-out-of-range',
			`Cells outside the grid on "${ctx.sheet.name}" were dropped`,
		);
		return;
	}
	const style = ctx.styles.cellXf(ixfe);
	if (value === null && style === 0) {
		return;
	}
	const cell = makeCell(value, style);
	if (richText) {
		cell.richText = richText;
	}
	ctx.sheet.cells.set(cellKey(row, col), cell);
}

function putString(
	ctx: SheetContext,
	row: number,
	col: number,
	ixfe: number,
	text: string,
	runs: readonly BiffRun[] | undefined,
): void {
	const rich = runs?.length ? richTextRuns(text, runs, (f) => ctx.styles.font(f)) : undefined;
	put(ctx, row, col, text, ixfe, rich);
}

function labelSst(
	ctx: SheetContext,
	cur: BiffCursor,
	row: number,
	col: number,
	ixfe: number,
): void {
	const index = cur.u32();
	const text = ctx.sst.strings[index];
	if (text === undefined) {
		sheetWarn(
			ctx,
			'missing-shared-string',
			`Cells on "${ctx.sheet.name}" refer to missing shared strings and were left empty`,
			a1(row, col),
		);
		put(ctx, row, col, '', ixfe);
		return;
	}
	putString(ctx, row, col, ixfe, text, ctx.sst.runs.get(index));
}

function richLabel(
	ctx: SheetContext,
	cur: BiffCursor,
	row: number,
	col: number,
	ixfe: number,
): void {
	const text = cur.unicodeString();
	const runs: BiffRun[] = [];
	if (cur.remaining >= 2) {
		const count = Math.min(cur.u16(), Math.floor(cur.remaining / 4));
		for (let i = 0; i < count; i++) {
			runs.push({ ich: cur.u16(), ifnt: cur.u16() });
		}
	}
	putString(ctx, row, col, ixfe, text, runs);
}

/** Handle a value cell record; returns `false` for other record types. */
export function readValueCell(ctx: SheetContext, record: BiffRecord): boolean {
	const cur = new BiffCursor(record.data, record.breaks);
	switch (record.type) {
		case RT.NUMBER: {
			const [row, col, ixfe] = [cur.u16(), cur.u16(), cur.u16()];
			put(ctx, row, col, numberValue(cur.f64()), ixfe);
			return true;
		}
		case RT.RK: {
			const [row, col, ixfe] = [cur.u16(), cur.u16(), cur.u16()];
			put(ctx, row, col, decodeRk(cur.u32()), ixfe);
			return true;
		}
		case RT.MULRK: {
			const row = cur.u16();
			const first = cur.u16();
			const count = Math.floor((cur.remaining - 2) / 6);
			for (let i = 0; i < count; i++) {
				const ixfe = cur.u16();
				put(ctx, row, first + i, decodeRk(cur.u32()), ixfe);
			}
			return true;
		}
		case RT.LABELSST: {
			const [row, col, ixfe] = [cur.u16(), cur.u16(), cur.u16()];
			labelSst(ctx, cur, row, col, ixfe);
			return true;
		}
		case RT.LABEL: {
			const [row, col, ixfe] = [cur.u16(), cur.u16(), cur.u16()];
			put(ctx, row, col, cur.unicodeString(), ixfe);
			return true;
		}
		case RT.RSTRING: {
			const [row, col, ixfe] = [cur.u16(), cur.u16(), cur.u16()];
			richLabel(ctx, cur, row, col, ixfe);
			return true;
		}
		case RT.BOOLERR: {
			const [row, col, ixfe] = [cur.u16(), cur.u16(), cur.u16()];
			const v = cur.u8();
			put(ctx, row, col, cur.u8() ? errorValue(v) : v !== 0, ixfe);
			return true;
		}
		case RT.BLANK: {
			const [row, col, ixfe] = [cur.u16(), cur.u16(), cur.u16()];
			put(ctx, row, col, null, ixfe);
			return true;
		}
		case RT.MULBLANK: {
			const row = cur.u16();
			const first = cur.u16();
			const count = Math.floor((cur.remaining - 2) / 2);
			for (let i = 0; i < count; i++) {
				put(ctx, row, first + i, null, cur.u16());
			}
			return true;
		}
		default:
			return false;
	}
}
