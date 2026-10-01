import type { RefCorner } from '../ast';
import { formatRefArea } from './references';
import { applyEdits, type TextEdit } from './token-utils';
import { tokenizeFormula } from './tokenizer';
import type { RefArea } from './tokens';

function rowPart(c: RefCorner, row: number): string {
	if (c.rowAbs) {
		return `R${c.row + 1}`;
	}
	return c.row === row ? 'R' : `R[${c.row - row}]`;
}

function colPart(c: RefCorner, col: number): string {
	if (c.colAbs) {
		return `C${c.col + 1}`;
	}
	return c.col === col ? 'C' : `C[${c.col - col}]`;
}

/**
 * R1C1 text of an area seen from the 0-based cell (row, col): `R1C1`,
 * `R[-1]C[2]`, `RC:R[3]C`, whole rows `R2:R[4]` (a single row is `R2`) and
 * whole columns `C1:C3`.
 */
export function formatRefAreaR1C1(area: RefArea, row: number, col: number): string {
	const { start, end } = area;
	switch (area.kind) {
		case 'cell':
			return rowPart(start, row) + colPart(start, col);
		case 'cells': {
			const from = rowPart(start, row) + colPart(start, col);
			return `${from}:${rowPart(end, row)}${colPart(end, col)}`;
		}
		case 'rows': {
			const a = rowPart(start, row);
			const b = rowPart(end, row);
			return a === b ? a : `${a}:${b}`;
		}
		case 'columns': {
			const a = colPart(start, col);
			const b = colPart(end, col);
			return a === b ? a : `${a}:${b}`;
		}
	}
}

/**
 * Convert the A1 references of a formula to R1C1 notation as seen from the
 * 0-based cell (row, col). Everything else (sheet prefixes, names, strings,
 * spacing) is kept.
 */
export function a1ToR1C1(formula: string, row: number, col: number): string {
	const tokens = tokenizeFormula(formula);
	const edits: TextEdit[] = [];
	for (const t of tokens) {
		if (t.type === 'ref' && t.area !== undefined) {
			const prefix = t.text.slice(0, t.qualifierLength ?? 0);
			edits.push({
				start: t.start,
				end: t.end,
				text: prefix + formatRefAreaR1C1(t.area, row, col),
			});
		}
	}
	return applyEdits(formula, edits);
}

/**
 * Convert the R1C1 references of a formula to A1 notation for the 0-based
 * cell (row, col). Relative offsets that leave the grid wrap around like
 * Excel (`R[-1]C` in row 1 is the last row).
 */
export function r1c1ToA1(formula: string, row: number, col: number): string {
	const tokens = tokenizeFormula(formula, { r1c1: { row, col } });
	const edits: TextEdit[] = [];
	for (const t of tokens) {
		if (t.type === 'ref' && t.area !== undefined) {
			const prefix = t.text.slice(0, t.qualifierLength ?? 0);
			edits.push({ start: t.start, end: t.end, text: prefix + formatRefArea(t.area) });
		}
	}
	return applyEdits(formula, edits);
}
