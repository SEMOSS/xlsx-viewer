import { MAX_COL_INDEX, MAX_COLS, MAX_ROW_INDEX, MAX_ROWS } from '../../address/constants';
import type { ErrorCode } from '../../types/cell';
import type { RefCorner, SheetQualifier } from '../ast';
import {
	CH_0,
	CH_9,
	CH_BANG,
	CH_DOLLAR,
	CH_LBRACKET,
	CH_LPAREN,
	CH_MINUS,
	CH_PLUS,
	CH_RBRACKET,
	isDigit,
	isLetter,
	isNameStart,
	isWordChar,
} from './chars';
import type { RefArea, RefKind } from './tokens';

/** Every error literal, as written in formulas. */
export const ERROR_LITERALS: readonly ErrorCode[] = [
	'#NULL!',
	'#DIV/0!',
	'#VALUE!',
	'#REF!',
	'#NAME?',
	'#NUM!',
	'#N/A',
	'#GETTING_DATA',
	'#SPILL!',
	'#CALC!',
	'#FIELD!',
	'#BLOCKED!',
	'#CONNECT!',
	'#BUSY!',
	'#UNKNOWN!',
];

/** Error literal starting at `p` (case-insensitive), or `null`. */
export function matchErrorCode(src: string, p: number): ErrorCode | null {
	const head = src.slice(p, p + 13).toUpperCase();
	for (const code of ERROR_LITERALS) {
		if (head.startsWith(code)) {
			return code;
		}
	}
	return null;
}

/** End of the run of word characters (name characters and `$`) starting at `p`. */
export function scanWord(src: string, p: number): number {
	const len = src.length;
	let i = p;
	while (i < len && isWordChar(src.charCodeAt(i))) {
		i++;
	}
	return i;
}

/** `[s, e)` is a plain identifier: starts like a name and holds no `$`. */
export function isPlainName(src: string, s: number, e: number): boolean {
	if (e <= s || !isNameStart(src.charCodeAt(s))) {
		return false;
	}
	for (let i = s; i < e; i++) {
		if (src.charCodeAt(i) === CH_DOLLAR) {
			return false;
		}
	}
	return true;
}

/** One side of an A1 reference: a cell (`$A$1`), a column (`$A`) or a row (`$1`). */
interface Part {
	kind: 'cell' | 'col' | 'row';
	row: number;
	col: number;
	rowAbs: boolean;
	colAbs: boolean;
}

/** 0-based row from 1-based digits in `[i, e)`, or -1. */
function readRow(src: string, i: number, e: number): number {
	if (i >= e) {
		return -1;
	}
	let n = 0;
	for (; i < e; i++) {
		const c = src.charCodeAt(i);
		if (c < CH_0 || c > CH_9) {
			return -1;
		}
		n = n * 10 + (c - CH_0);
		if (n > MAX_ROWS) {
			return -1;
		}
	}
	return n - 1;
}

/** Classify the word `[s, e)` as an A1 reference part. */
function classifyA1Part(src: string, s: number, e: number): Part | null {
	let i = s;
	const firstAbs = src.charCodeAt(i) === CH_DOLLAR;
	if (firstAbs) {
		i++;
	}
	let col = 0;
	const lettersStart = i;
	while (i < e && isLetter(src.charCodeAt(i))) {
		if (i - lettersStart === 3) {
			return null;
		}
		col = col * 26 + ((src.charCodeAt(i) | 32) - 96);
		i++;
	}
	if (i > lettersStart) {
		col -= 1;
		if (col > MAX_COL_INDEX) {
			return null;
		}
		if (i === e) {
			return { kind: 'col', row: 0, col, rowAbs: false, colAbs: firstAbs };
		}
		const rowAbs = src.charCodeAt(i) === CH_DOLLAR;
		const row = readRow(src, rowAbs ? i + 1 : i, e);
		return row < 0 ? null : { kind: 'cell', row, col, rowAbs, colAbs: firstAbs };
	}
	const row = readRow(src, i, e);
	return row < 0 ? null : { kind: 'row', row, col: 0, rowAbs: firstAbs, colAbs: false };
}

/** A character that turns the preceding word into something other than a reference. */
function blocksReference(code: number): boolean {
	return code === CH_LPAREN || code === CH_LBRACKET || code === CH_BANG;
}

const AREA_KIND: Record<Part['kind'], RefKind> = { cell: 'cells', col: 'columns', row: 'rows' };

function partCorner(p: Part, isEnd: boolean): RefCorner {
	if (p.kind === 'col') {
		return { row: isEnd ? MAX_ROW_INDEX : 0, col: p.col, rowAbs: false, colAbs: p.colAbs };
	}
	if (p.kind === 'row') {
		return { row: p.row, col: isEnd ? MAX_COL_INDEX : 0, rowAbs: p.rowAbs, colAbs: false };
	}
	return { row: p.row, col: p.col, rowAbs: p.rowAbs, colAbs: p.colAbs };
}

function areaOf(first: Part, second: Part | null): RefArea {
	if (second === null) {
		if (first.kind === 'cell') {
			const corner = partCorner(first, false);
			return { kind: 'cell', start: corner, end: corner };
		}
		return {
			kind: AREA_KIND[first.kind],
			start: partCorner(first, false),
			end: partCorner(first, true),
		};
	}
	return {
		kind: AREA_KIND[first.kind],
		start: partCorner(first, false),
		end: partCorner(second, true),
	};
}

export interface ScannedArea {
	end: number;
	area: RefArea;
}

/**
 * Read an A1 reference at `p`: `A1`, `$A$1:B2`, `A:C`, `$1:$3`. Lone
 * columns or rows (`A`, `1`) are not references. Returns `null` when the
 * text at `p` is not a complete reference.
 */
export function scanA1Area(src: string, p: number): ScannedArea | null {
	const e1 = scanWord(src, p);
	if (e1 === p) {
		return null;
	}
	const first = classifyA1Part(src, p, e1);
	if (first === null) {
		return null;
	}
	const after = src.charCodeAt(e1);
	if (after === 58 /* : */) {
		const s2 = e1 + 1;
		const e2 = scanWord(src, s2);
		if (e2 > s2 && !blocksReference(src.charCodeAt(e2))) {
			const second = classifyA1Part(src, s2, e2);
			if (second !== null && second.kind === first.kind) {
				return { end: e2, area: areaOf(first, second) };
			}
		}
	}
	if (first.kind !== 'cell' || blocksReference(after)) {
		return null;
	}
	return { end: e1, area: areaOf(first, null) };
}

function wrap(value: number, size: number): number {
	return ((value % size) + size) % size;
}

interface R1C1Axis {
	end: number;
	abs: boolean;
	/** Absolute 1-based index, or relative offset. */
	value: number;
}

/** Number after `R` / `C`: `5`, `[-2]`, or nothing (offset 0). `null` when malformed. */
function readR1C1Axis(src: string, i: number, limit: number): R1C1Axis | null {
	const c = src.charCodeAt(i);
	if (c === CH_LBRACKET) {
		let j = i + 1;
		const sign = src.charCodeAt(j);
		if (sign === CH_PLUS || sign === CH_MINUS) {
			j++;
		}
		const digits = j;
		let n = 0;
		while (isDigit(src.charCodeAt(j))) {
			n = n * 10 + (src.charCodeAt(j) - CH_0);
			j++;
		}
		if (j === digits || src.charCodeAt(j) !== CH_RBRACKET || n >= limit) {
			return null;
		}
		return { end: j + 1, abs: false, value: sign === CH_MINUS ? -n : n };
	}
	if (isDigit(c)) {
		let j = i;
		let n = 0;
		while (isDigit(src.charCodeAt(j))) {
			n = n * 10 + (src.charCodeAt(j) - CH_0);
			j++;
		}
		return n < 1 || n > limit ? null : { end: j, abs: true, value: n };
	}
	return { end: i, abs: false, value: 0 };
}

/** One R1C1 part at `p`: `R1C1`, `R[-1]C`, `RC[2]` (cell), `R2` (row), `C[1]` (column). */
function scanR1C1Part(
	src: string,
	p: number,
	base: { row: number; col: number },
): (Part & { end: number }) | null {
	let i = p;
	let row: R1C1Axis | null = null;
	let col: R1C1Axis | null = null;
	if ((src.charCodeAt(i) | 32) === 114 /* r */) {
		row = readR1C1Axis(src, i + 1, MAX_ROWS);
		if (row === null) {
			return null;
		}
		i = row.end;
	}
	if ((src.charCodeAt(i) | 32) === 99 /* c */) {
		col = readR1C1Axis(src, i + 1, MAX_COLS);
		if (col === null) {
			return null;
		}
		i = col.end;
	}
	if (
		(row === null && col === null) ||
		isWordChar(src.charCodeAt(i)) ||
		blocksReference(src.charCodeAt(i))
	) {
		return null;
	}
	const r = row === null ? 0 : row.abs ? row.value - 1 : wrap(base.row + row.value, MAX_ROWS);
	const c = col === null ? 0 : col.abs ? col.value - 1 : wrap(base.col + col.value, MAX_COLS);
	const kind = row !== null && col !== null ? 'cell' : row !== null ? 'row' : 'col';
	return { kind, row: r, col: c, rowAbs: row?.abs ?? false, colAbs: col?.abs ?? false, end: i };
}

/**
 * Read an R1C1 reference at `p` (`R1C1`, `R[1]C:R[2]C[3]`, `R2`, `C1:C3`),
 * resolving relative parts against `base`. Relative offsets wrap around the
 * grid edges like Excel does.
 */
export function scanR1C1Area(
	src: string,
	p: number,
	base: { row: number; col: number },
): ScannedArea | null {
	const first = scanR1C1Part(src, p, base);
	if (first === null) {
		return null;
	}
	if (src.charCodeAt(first.end) === 58 /* : */) {
		const second = scanR1C1Part(src, first.end + 1, base);
		if (second !== null && second.kind === first.kind) {
			return { end: second.end, area: areaOf(first, second) };
		}
	}
	return { end: first.end, area: areaOf(first, null) };
}

/**
 * Split the text of a quoted sheet prefix (quotes removed, `''` already
 * un-doubled): `My Sheet`, `Sheet 1:Sheet 3`, `[1]Sheet 1`,
 * `C:\Docs\[Book.xlsx]Sheet1`.
 */
export function parseQualifierText(raw: string): SheetQualifier {
	let rest = raw;
	let workbook: string | undefined;
	const lb = raw.indexOf('[');
	if (lb >= 0) {
		const rb = raw.indexOf(']', lb + 1);
		if (rb > lb) {
			workbook = raw.slice(0, lb) + raw.slice(lb + 1, rb);
			rest = raw.slice(rb + 1);
		}
	}
	if (workbook === undefined && /[\\/]/.test(rest)) {
		// `'C:\Docs\Book.xlsx'!Name`: a workbook-level name in another file.
		return { sheet: '', workbook: rest };
	}
	const colon = rest.indexOf(':');
	const q: SheetQualifier =
		colon >= 0 ? { sheet: rest.slice(0, colon), sheetEnd: rest.slice(colon + 1) } : { sheet: rest };
	if (workbook !== undefined) {
		q.workbook = workbook;
	}
	return q;
}
