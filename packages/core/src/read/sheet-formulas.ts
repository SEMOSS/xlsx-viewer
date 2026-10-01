import { a1, cellKey, keyCol, keyRow } from '../address/cell-address';
import { parseRange } from '../address/range';
import { stripFormulaPrefixes } from '../formula/prefixes';
import { translateFormula } from '../formula/translate';
import type { RangeAddress } from '../types/address';
import type { Cell } from '../types/cell';
import { decodeXmlEntities } from '../xml/escape';

/** Display form of a file formula (prefixes stripped); falls back to the raw text if that fails. */
function displayFormula(raw: string): string {
	try {
		return stripFormulaPrefixes(raw);
	} catch {
		return raw;
	}
}

interface SharedMaster {
	row: number;
	col: number;
	formula: string;
}

interface PendingMember {
	key: number;
	row: number;
	col: number;
	si: number;
}

export interface FormulaIssue {
	code: string;
	message: string;
	cell?: string;
}

/**
 * Formula bookkeeping while `<sheetData>` is scanned: expands shared
 * formulas (every member gets its own text, translated from the master),
 * records array formulas and links the cells of an array result to its
 * anchor once the whole sheet has been read.
 */
export class FormulaCollector {
	dataTableCells = 0;
	private readonly masters = new Map<number, SharedMaster>();
	private readonly pending: PendingMember[] = [];
	private readonly arrays: Array<{ key: number; ref: RangeAddress }> = [];
	private readonly issues: FormulaIssue[] = [];
	private translateFailed = false;

	constructor(
		private readonly cells: Map<number, Cell>,
		private readonly dynamicArrayCm: ReadonlySet<number> | undefined,
	) {}

	/** Attach the formula of `<f>` (attributes and raw text) to a freshly built cell. */
	add(
		cell: Cell,
		key: number,
		row: number,
		col: number,
		attrs: Record<string, string> | undefined,
		rawText: string,
		cm: number,
	): void {
		const raw = rawText.indexOf('&') === -1 ? rawText : decodeXmlEntities(rawText);
		const text = raw ? displayFormula(raw) : '';
		const t = attrs?.t;
		if (t === undefined || t === 'normal') {
			if (text) {
				cell.formula = text;
			}
			return;
		}
		if (t === 'shared') {
			const si = attrs?.si !== undefined ? Number(attrs.si) : NaN;
			if (text) {
				cell.formula = text;
				if (Number.isInteger(si) && (attrs?.ref !== undefined || !this.masters.has(si))) {
					this.masters.set(si, { row, col, formula: text });
				}
				return;
			}
			if (!Number.isInteger(si)) {
				return;
			}
			const master = this.masters.get(si);
			if (master) {
				this.expand(cell, master, row, col);
			} else {
				this.pending.push({ key, row, col, si });
			}
			return;
		}
		if (t === 'array') {
			if (text) {
				cell.formula = text;
			}
			const ref = (attrs?.ref ? parseRange(attrs.ref) : null) ?? {
				top: row,
				left: col,
				bottom: row,
				right: col,
			};
			const dynamic = cm >= 0 && this.dynamicArrayCm !== undefined && this.dynamicArrayCm.has(cm);
			cell.array = { ref, dynamic };
			if (ref.top !== ref.bottom || ref.left !== ref.right) {
				this.arrays.push({ key, ref });
			}
			return;
		}
		if (t === 'dataTable') {
			this.dataTableCells++;
			return;
		}
		if (text) {
			cell.formula = text;
		}
	}

	/** Give a shared-formula member its own text, translated from the master. */
	private expand(cell: Cell, master: SharedMaster, row: number, col: number): void {
		try {
			cell.formula = translateFormula(master.formula, row - master.row, col - master.col);
		} catch (e) {
			if (!this.translateFailed) {
				this.translateFailed = true;
				const reason = e instanceof Error ? e.message : String(e);
				this.issues.push({
					code: 'shared-formula-error',
					message:
						`The shared formula "${master.formula}" could not be expanded to ${a1(row, col)} ` +
						`(${reason}); the cached value is kept`,
					cell: a1(row, col),
				});
			}
		}
	}

	/** Resolve forward shared-formula references and link array results. */
	finish(): FormulaIssue[] {
		const issues = this.issues;
		let orphan = false;
		for (const p of this.pending) {
			const cell = this.cells.get(p.key);
			const master = this.masters.get(p.si);
			if (!cell) {
				continue;
			}
			if (master) {
				this.expand(cell, master, p.row, p.col);
			} else if (!orphan) {
				orphan = true;
				issues.push({
					code: 'shared-formula-missing',
					message:
						`Shared formula ${p.si} used by ${a1(p.row, p.col)} has no master cell; ` +
						'the cached value is kept',
					cell: a1(p.row, p.col),
				});
			}
		}
		for (const { key, ref } of this.arrays) {
			this.linkArray(key, ref);
		}
		return issues;
	}

	private linkArray(anchor: number, ref: RangeAddress): void {
		const link = (k: number): void => {
			if (k === anchor) {
				return;
			}
			const cell = this.cells.get(k);
			if (cell && cell.formula === undefined && cell.arrayAnchor === undefined) {
				cell.arrayAnchor = anchor;
			}
		};
		const count = (ref.bottom - ref.top + 1) * (ref.right - ref.left + 1);
		if (count <= this.cells.size) {
			for (let r = ref.top; r <= ref.bottom; r++) {
				for (let c = ref.left; c <= ref.right; c++) {
					link(cellKey(r, c));
				}
			}
			return;
		}
		for (const k of this.cells.keys()) {
			const r = keyRow(k);
			const c = keyCol(k);
			if (r >= ref.top && r <= ref.bottom && c >= ref.left && c <= ref.right) {
				link(k);
			}
		}
	}
}
