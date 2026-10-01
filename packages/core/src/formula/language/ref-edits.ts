import type { RefCorner } from '../ast';
import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../../address/constants';
import { formatRefArea } from './references';
import type { TextEdit } from './token-utils';
import type { FormulaToken, RefArea } from './tokens';

/**
 * A reference into the formula's own workbook on a single sheet (not an
 * external `[1]Sheet1!A1` and not a 3-D `Sheet1:Sheet3!A1` reference).
 */
export function isLocalReference(t: FormulaToken): t is FormulaToken & { area: RefArea } {
	if (t.type !== 'ref' || t.area === undefined) {
		return false;
	}
	return (
		t.sheet === undefined || (t.sheet.workbook === undefined && t.sheet.sheetEnd === undefined)
	);
}

/** Sheet a local reference points to: its qualifier, else the formula's own sheet. */
export function referenceSheet(t: FormulaToken, formulaSheet: string): string {
	return t.sheet === undefined ? formulaSheet : t.sheet.sheet;
}

/**
 * Edit replacing the reference token at `index` with `area` (keeping the
 * sheet prefix as written unless `prefix` is given), or with `#REF!` when
 * `area` is `null`. A spill operator after a dead reference is dropped.
 */
export function referenceEdit(
	tokens: readonly FormulaToken[],
	index: number,
	area: RefArea | null,
	prefix?: string,
): TextEdit {
	const t = tokens[index];
	const qualifier = prefix ?? t.text.slice(0, t.qualifierLength ?? 0);
	if (area === null) {
		const next = tokens[index + 1];
		const end = next !== undefined && next.type === 'spill' ? next.end : t.end;
		return { start: t.start, end, text: `${qualifier}#REF!` };
	}
	return { start: t.start, end: t.end, text: qualifier + formatRefArea(area) };
}

function moveCorner(c: RefCorner, dr: number, dc: number, all: boolean): RefCorner | null {
	const row = c.rowAbs && !all ? c.row : c.row + dr;
	const col = c.colAbs && !all ? c.col : c.col + dc;
	if (row < 0 || row > MAX_ROW_INDEX || col < 0 || col > MAX_COL_INDEX) {
		return null;
	}
	return { row, col, rowAbs: c.rowAbs, colAbs: c.colAbs };
}

/**
 * Offset an area. `all` moves absolute parts too (cut / paste); otherwise
 * only relative parts move (copy / fill). Whole columns ignore the row
 * offset and whole rows the column offset. Returns `null` when a corner
 * leaves the grid and `area` itself when nothing changes.
 */
export function offsetArea(
	area: RefArea,
	rowDelta: number,
	colDelta: number,
	all = false,
): RefArea | null {
	const dr = area.kind === 'columns' ? 0 : rowDelta;
	const dc = area.kind === 'rows' ? 0 : colDelta;
	const start = moveCorner(area.start, dr, dc, all);
	if (start === null) {
		return null;
	}
	if (area.kind === 'cell') {
		return start.row === area.start.row && start.col === area.start.col
			? area
			: { kind: 'cell', start, end: start };
	}
	const end = moveCorner(area.end, dr, dc, all);
	if (end === null) {
		return null;
	}
	const same =
		start.row === area.start.row &&
		start.col === area.start.col &&
		end.row === area.end.row &&
		end.col === area.end.col;
	return same ? area : { kind: area.kind, start, end };
}
