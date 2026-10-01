/**
 * Formula cells: FORMULA records with their cached results (a string
 * result follows in a STRING record), shared formulas (SHRFMLA), array
 * formulas (ARRAY) and what-if data tables (TABLE). Members of a shared or
 * array formula only hold a ptgExp pointing at the block's top-left cell;
 * the block record follows the first member, so members are resolved once
 * the whole sheet has been read. Shared formulas are expanded so every
 * cell carries its own formula text.
 */
import { a1, cellKey } from '../../address/cell-address';
import type { RangeAddress } from '../../types/address';
import type { Cell, CellValue } from '../../types/cell';
import { BiffCursor } from './biff-cursor';
import { RT } from './biff-opcodes';
import type { BiffRecord } from './biff-records';
import { decodeRgce, type Ptg, renderFormula, XlsExternalReferenceError } from './formula';
import { errorValue } from './sheet-cells';
import {
	makeCell,
	makeRange,
	rangeHas,
	type SheetContext,
	sheetWarn,
	validCell,
} from './sheet-context';

interface Block {
	range: RangeAddress;
	ptgs?: Ptg[];
	error?: unknown;
}

interface Member {
	key: number;
	row: number;
	col: number;
	anchor: number;
}

/** Cached result of a FORMULA record; `string` means a STRING record follows. */
function cachedResult(bytes: Uint8Array): { value: CellValue; string: boolean } {
	if (bytes[6] === 0xff && bytes[7] === 0xff) {
		switch (bytes[0]) {
			case 0x00:
				return { value: '', string: true };
			case 0x01:
				return { value: bytes[2] !== 0, string: false };
			case 0x02:
				return { value: errorValue(bytes[2]), string: false };
			case 0x03:
				return { value: '', string: false };
			default:
				return { value: null, string: false };
		}
	}
	const v = new DataView(bytes.buffer, bytes.byteOffset, 8).getFloat64(0, true);
	return { value: Number.isFinite(v) ? v : errorValue(0x24), string: false };
}

/** `RefU`: rows as 16-bit, columns as 8-bit values. */
function readBlock(cur: BiffCursor): Block | undefined {
	const range = makeRange(cur.u16(), cur.u16(), cur.u8(), cur.u8());
	return range ? { range } : undefined;
}

/** Record types handled by {@link FormulaCells.read}. */
export const FORMULA_RECORDS: readonly number[] = [
	RT.FORMULA,
	RT.STRING,
	RT.SHRFMLA,
	RT.ARRAY,
	RT.TABLE,
];

export class FormulaCells {
	private readonly shared = new Map<number, Block>();
	private readonly arrays = new Map<number, Block>();
	private readonly members: Member[] = [];
	/** Formula cell waiting for its STRING record. */
	private stringTarget: number | undefined;
	private dataTables = false;
	private external: string | undefined;
	private failure: { reason: string; cell: string } | undefined;

	constructor(private readonly ctx: SheetContext) {}

	/** Handle a formula-related record; returns `false` for other record types. */
	read(record: BiffRecord): boolean {
		switch (record.type) {
			case RT.FORMULA:
				this.formula(record);
				return true;
			case RT.STRING:
				this.string(record);
				return true;
			case RT.SHRFMLA:
				this.block(record, this.shared, 2);
				return true;
			case RT.ARRAY:
				this.block(record, this.arrays, 6);
				return true;
			case RT.TABLE:
				this.dataTables = true;
				return true;
			default:
				return false;
		}
	}

	/** Another record came between a FORMULA and its STRING record. */
	endString(): void {
		this.stringTarget = undefined;
	}

	private fail(e: unknown, row: number, col: number): void {
		if (e instanceof XlsExternalReferenceError) {
			this.external ??= e.source;
		} else {
			this.failure ??= { reason: e instanceof Error ? e.message : String(e), cell: a1(row, col) };
		}
	}

	private formula(record: BiffRecord): void {
		this.stringTarget = undefined;
		const cur = new BiffCursor(record.data, record.breaks);
		const [row, col, ixfe] = [cur.u16(), cur.u16(), cur.u16()];
		const result = cachedResult(cur.bytes(8));
		cur.skip(6);
		const cce = cur.u16();
		if (!validCell(row, col)) {
			return;
		}
		const key = cellKey(row, col);
		const cell: Cell = makeCell(result.value, this.ctx.styles.cellXf(ixfe));
		this.ctx.sheet.cells.set(key, cell);
		if (result.string) {
			this.stringTarget = key;
		}
		let ptgs: Ptg[];
		try {
			ptgs = decodeRgce(cur.bytes(Math.min(cce, cur.remaining)), cur.bytes(cur.remaining));
		} catch (e) {
			this.fail(e, row, col);
			return;
		}
		const first = ptgs[0];
		if (first?.k === 'exp') {
			this.members.push({ key, row, col, anchor: cellKey(first.row, first.col) });
			return;
		}
		if (first?.k === 'tbl') {
			this.dataTables = true;
			return;
		}
		try {
			cell.formula = renderFormula(ptgs, this.ctx.env, { row, col, mode: 'cell' });
		} catch (e) {
			this.fail(e, row, col);
		}
	}

	private string(record: BiffRecord): void {
		const key = this.stringTarget;
		this.stringTarget = undefined;
		const cell = key === undefined ? undefined : this.ctx.sheet.cells.get(key);
		if (cell) {
			cell.value = new BiffCursor(record.data, record.breaks).unicodeString();
		}
	}

	/** SHRFMLA / ARRAY: the block's range and tokens, keyed by its top-left cell. */
	private block(record: BiffRecord, into: Map<number, Block>, skip: number): void {
		const cur = new BiffCursor(record.data, record.breaks);
		const block = readBlock(cur);
		cur.skip(skip);
		const cce = cur.u16();
		if (!block) {
			return;
		}
		try {
			block.ptgs = decodeRgce(cur.bytes(Math.min(cce, cur.remaining)), cur.bytes(cur.remaining));
		} catch (e) {
			block.error = e;
		}
		into.set(cellKey(block.range.top, block.range.left), block);
	}

	private findShared(row: number, col: number): Block | undefined {
		for (const b of this.shared.values()) {
			if (rangeHas(b.range, row, col)) {
				return b;
			}
		}
		return undefined;
	}

	private render(cell: Cell, block: Block, m: Member, mode: 'cell' | 'shared'): void {
		if (!block.ptgs) {
			this.fail(block.error, m.row, m.col);
			return;
		}
		try {
			cell.formula = renderFormula(block.ptgs, this.ctx.env, { row: m.row, col: m.col, mode });
		} catch (e) {
			this.fail(e, m.row, m.col);
		}
	}

	/** Resolve shared / array members and report what could not be read. */
	finish(): void {
		let orphan: string | undefined;
		for (const m of this.members) {
			const cell = this.ctx.sheet.cells.get(m.key);
			if (!cell) {
				continue;
			}
			const array = this.arrays.get(m.anchor);
			if (array && rangeHas(array.range, m.row, m.col)) {
				if (m.key === m.anchor) {
					cell.array = { ref: array.range, dynamic: false };
					this.render(cell, array, m, 'cell');
				} else {
					cell.arrayAnchor = m.anchor;
				}
				continue;
			}
			const shared = this.shared.get(m.anchor) ?? this.findShared(m.row, m.col);
			if (shared && rangeHas(shared.range, m.row, m.col)) {
				this.render(cell, shared, m, 'shared');
			} else {
				orphan ??= a1(m.row, m.col);
			}
		}
		this.report(orphan);
	}

	private report(orphan: string | undefined): void {
		const name = this.ctx.sheet.name;
		if (this.external) {
			sheetWarn(
				this.ctx,
				'external-link',
				`Formulas on "${name}" refer to ${this.external}; ` +
					'external links are not imported, so their cached values are kept',
			);
		}
		if (this.failure) {
			sheetWarn(
				this.ctx,
				'formula-unreadable',
				`Some formulas on "${name}" could not be read (${this.failure.reason}); ` +
					'their cached values are kept',
				this.failure.cell,
			);
		}
		if (orphan) {
			sheetWarn(
				this.ctx,
				'shared-formula-missing',
				`The shared formula used by ${orphan} on "${name}" is missing; the cached value is kept`,
				orphan,
			);
		}
		if (this.dataTables) {
			sheetWarn(
				this.ctx,
				'data-table',
				`"${name}" contains a what-if data table; ` +
					'its cached results are shown but not recalculated',
			);
		}
	}
}
