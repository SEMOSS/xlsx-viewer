import type { CellAddress } from '../types/address';
import { MAX_COLS, MAX_ROWS } from './constants';

/**
 * Pack a 0-based (row, col) pair into a single integer map key.
 * Keys sort in row-major order, so sorting keys numerically yields sheet order.
 */
export function cellKey(row: number, col: number): number {
	return row * MAX_COLS + col;
}

/** Row of a key produced by {@link cellKey}. */
export function keyRow(key: number): number {
	return Math.floor(key / MAX_COLS);
}

/** Column of a key produced by {@link cellKey}. */
export function keyCol(key: number): number {
	return key % MAX_COLS;
}

/** Convert a key back to an address. */
export function keyToAddress(key: number): CellAddress {
	return { row: keyRow(key), col: keyCol(key) };
}

/** 0-based column index to letters: 0 -> `A`, 25 -> `Z`, 26 -> `AA`. */
export function columnToLetters(col: number): string {
	let n = col + 1;
	let out = '';
	while (n > 0) {
		const rem = (n - 1) % 26;
		out = String.fromCharCode(65 + rem) + out;
		n = Math.floor((n - 1) / 26);
	}
	return out;
}

/** Column letters to a 0-based index (`A` -> 0). Returns -1 for invalid input. */
export function lettersToColumn(letters: string): number {
	let n = 0;
	for (let i = 0; i < letters.length; i++) {
		const code = letters.charCodeAt(i);
		let v: number;
		if (code >= 65 && code <= 90) {
			v = code - 64;
		} else if (code >= 97 && code <= 122) {
			v = code - 96;
		} else {
			return -1;
		}
		n = n * 26 + v;
		if (n > MAX_COLS) {
			return -1;
		}
	}
	return letters.length === 0 ? -1 : n - 1;
}

/** Format an address as A1 notation, optionally with `$` absolute markers. */
export function formatCellAddress(addr: CellAddress, absCol = false, absRow = false): string {
	return `${absCol ? '$' : ''}${columnToLetters(addr.col)}${absRow ? '$' : ''}${addr.row + 1}`;
}

/** Shorthand for {@link formatCellAddress} with plain row/col arguments. */
export function a1(row: number, col: number): string {
	return `${columnToLetters(col)}${row + 1}`;
}

const CELL_RE = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/;

/** Parse `A1`, `$B$2` style references. Returns `null` when invalid or out of bounds. */
export function parseCellAddress(ref: string): CellAddress | null {
	const m = CELL_RE.exec(ref.trim());
	if (!m) {
		return null;
	}
	const col = lettersToColumn(m[1]);
	const row = Number(m[2]) - 1;
	if (col < 0 || row < 0 || row >= MAX_ROWS) {
		return null;
	}
	return { row, col };
}

/** Like {@link parseCellAddress} but also reports which parts were absolute. */
export function parseCellAddressWithFlags(
	ref: string,
): (CellAddress & { absRow: boolean; absCol: boolean }) | null {
	const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$/.exec(ref.trim());
	if (!m) {
		return null;
	}
	const col = lettersToColumn(m[2]);
	const row = Number(m[4]) - 1;
	if (col < 0 || row < 0 || row >= MAX_ROWS) {
		return null;
	}
	return { row, col, absCol: m[1] === '$', absRow: m[3] === '$' };
}

export function addressEquals(a: CellAddress, b: CellAddress): boolean {
	return a.row === b.row && a.col === b.col;
}
