import type { Cell, CellAddress, RangeAddress, Worksheet } from 'xlsx-viewer-core';
import { cellKey, keyCol, keyRow, MAX_COL_INDEX, MAX_ROW_INDEX } from 'xlsx-viewer-core';

export type Direction = 'up' | 'down' | 'left' | 'right';

/** Read-only view of which cells hold data, with per-line indexes built lazily. */
export class DataProbe {
	private rowIndex?: Map<number, number[]>;
	private colIndex?: Map<number, number[]>;

	constructor(
		readonly sheet: Worksheet,
		readonly isHiddenRow: (row: number) => boolean,
		readonly isHiddenCol: (col: number) => boolean,
	) {}

	hasData(row: number, col: number): boolean {
		const c: Cell | undefined = this.sheet.cells.get(cellKey(row, col));
		return c !== undefined && (c.value !== null || c.formula !== undefined);
	}

	private build(): void {
		this.rowIndex = new Map();
		this.colIndex = new Map();
		for (const [key, cell] of this.sheet.cells) {
			if (cell.value === null && cell.formula === undefined) {
				continue;
			}
			const r = keyRow(key);
			const c = keyCol(key);
			let rl = this.rowIndex.get(r);
			if (!rl) {
				rl = [];
				this.rowIndex.set(r, rl);
			}
			rl.push(c);
			let cl = this.colIndex.get(c);
			if (!cl) {
				cl = [];
				this.colIndex.set(c, cl);
			}
			cl.push(r);
		}
		for (const l of this.rowIndex.values()) {
			l.sort((a, b) => a - b);
		}
		for (const l of this.colIndex.values()) {
			l.sort((a, b) => a - b);
		}
	}

	/** Sorted columns with data in a row. */
	colsInRow(row: number): number[] {
		if (!this.rowIndex) {
			this.build();
		}
		return this.rowIndex!.get(row) ?? [];
	}

	/** Sorted rows with data in a column. */
	rowsInCol(col: number): number[] {
		if (!this.colIndex) {
			this.build();
		}
		return this.colIndex!.get(col) ?? [];
	}

	/** Last cell with data (Ctrl+End): the bottom-most row and right-most column used. */
	lastUsedCell(): CellAddress {
		let row = 0;
		let col = 0;
		for (const [key, cell] of this.sheet.cells) {
			if (cell.value === null && cell.formula === undefined && cell.style === undefined) {
				continue;
			}
			row = Math.max(row, keyRow(key));
			col = Math.max(col, keyCol(key));
		}
		return { row, col };
	}
}

function step(dir: Direction): [number, number] {
	switch (dir) {
		case 'up':
			return [-1, 0];
		case 'down':
			return [1, 0];
		case 'left':
			return [0, -1];
		case 'right':
			return [0, 1];
	}
}

function clampAddr(row: number, col: number): CellAddress {
	return {
		row: Math.max(0, Math.min(MAX_ROW_INDEX, row)),
		col: Math.max(0, Math.min(MAX_COL_INDEX, col)),
	};
}

/** One cell in a direction, skipping hidden rows / columns. */
export function moveOne(probe: DataProbe, from: CellAddress, dir: Direction): CellAddress {
	const [dr, dc] = step(dir);
	let { row, col } = from;
	do {
		row += dr;
		col += dc;
		if (row < 0 || col < 0 || row > MAX_ROW_INDEX || col > MAX_COL_INDEX) {
			return from;
		}
	} while ((dr !== 0 && probe.isHiddenRow(row)) || (dc !== 0 && probe.isHiddenCol(col)));
	return clampAddr(row, col);
}

/**
 * Ctrl+Arrow: jump to the edge of the current data region, or to the next
 * cell with data, or to the sheet edge. Uses the sparse line index, so it
 * is fast on huge sheets.
 */
export function jumpToEdge(probe: DataProbe, from: CellAddress, dir: Direction): CellAddress {
	const vertical = dir === 'up' || dir === 'down';
	const forward = dir === 'down' || dir === 'right';
	const line = vertical ? probe.rowsInCol(from.col) : probe.colsInRow(from.row);
	const pos = vertical ? from.row : from.col;
	const max = vertical ? MAX_ROW_INDEX : MAX_COL_INDEX;
	const has = new Set(line);
	const hidden = vertical ? probe.isHiddenRow : probe.isHiddenCol;
	const neighbour = (p: number) => {
		let n = p + (forward ? 1 : -1);
		while (n >= 0 && n <= max && hidden(n)) {
			n += forward ? 1 : -1;
		}
		return n;
	};
	const toAddr = (p: number): CellAddress =>
		vertical ? { row: p, col: from.col } : { row: from.row, col: p };

	const next = neighbour(pos);
	if (next < 0 || next > max) {
		return from;
	}
	if (has.has(pos) && has.has(next)) {
		// Inside a block: run to its last filled cell.
		let p = next;
		for (;;) {
			const n = neighbour(p);
			if (n < 0 || n > max || !has.has(n)) {
				return toAddr(p);
			}
			p = n;
		}
	}
	// Find the next filled cell in the direction.
	if (forward) {
		for (const p of line) {
			if (p > pos && !hidden(p)) {
				return toAddr(p);
			}
		}
		return toAddr(max);
	}
	for (let i = line.length - 1; i >= 0; i--) {
		if (line[i] < pos && !hidden(line[i])) {
			return toAddr(line[i]);
		}
	}
	return toAddr(0);
}

/**
 * Ctrl+A / Ctrl+Shift+*: the current region, i.e. the block of data around
 * the cell bounded by empty rows and columns.
 */
export function currentRegion(probe: DataProbe, at: CellAddress): RangeAddress {
	let r: RangeAddress = { top: at.row, left: at.col, bottom: at.row, right: at.col };
	let grew = true;
	let guard = 0;
	while (grew && guard++ < 10000) {
		grew = false;
		const rowHas = (row: number) => {
			for (let c = Math.max(0, r.left - 1); c <= Math.min(MAX_COL_INDEX, r.right + 1); c++) {
				if (probe.hasData(row, c)) {
					return true;
				}
			}
			return false;
		};
		const colHas = (col: number) => {
			for (let row = Math.max(0, r.top - 1); row <= Math.min(MAX_ROW_INDEX, r.bottom + 1); row++) {
				if (probe.hasData(row, col)) {
					return true;
				}
			}
			return false;
		};
		if (r.top > 0 && rowHas(r.top - 1)) {
			r = { ...r, top: r.top - 1 };
			grew = true;
		}
		if (r.bottom < MAX_ROW_INDEX && rowHas(r.bottom + 1)) {
			r = { ...r, bottom: r.bottom + 1 };
			grew = true;
		}
		if (r.left > 0 && colHas(r.left - 1)) {
			r = { ...r, left: r.left - 1 };
			grew = true;
		}
		if (r.right < MAX_COL_INDEX && colHas(r.right + 1)) {
			r = { ...r, right: r.right + 1 };
			grew = true;
		}
	}
	return r;
}
