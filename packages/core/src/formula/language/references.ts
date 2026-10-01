import { columnToLetters } from '../../address/cell-address';
import { quoteSheetName } from '../../address/sheet-name';
import type { RangeAddress } from '../../types/address';
import type { RefCorner, SheetQualifier } from '../ast';
import { tokenizeFormula } from './tokenizer';
import type { RefArea, TokenizeOptions } from './tokens';

/** A1 text of a cell corner, with `$` markers: `$A$1`. */
export function formatRefCorner(c: RefCorner): string {
	return `${c.colAbs ? '$' : ''}${columnToLetters(c.col)}${c.rowAbs ? '$' : ''}${c.row + 1}`;
}

function formatColumnPart(c: RefCorner): string {
	return `${c.colAbs ? '$' : ''}${columnToLetters(c.col)}`;
}

function formatRowPart(c: RefCorner): string {
	return `${c.rowAbs ? '$' : ''}${c.row + 1}`;
}

/** A1 text of a reference area (no sheet): `A1`, `$A$1:B2`, `A:C`, `$1:$3`. */
export function formatRefArea(area: RefArea): string {
	switch (area.kind) {
		case 'cell':
			return formatRefCorner(area.start);
		case 'cells':
			return `${formatRefCorner(area.start)}:${formatRefCorner(area.end)}`;
		case 'columns':
			return `${formatColumnPart(area.start)}:${formatColumnPart(area.end)}`;
		case 'rows':
			return `${formatRowPart(area.start)}:${formatRowPart(area.end)}`;
	}
}

function needsQuotes(name: string): boolean {
	return quoteSheetName(name) !== name;
}

/**
 * Sheet prefix including the trailing `!`: `Sheet1!`, `'My Sheet'!`,
 * `'Sheet 1:Sheet 3'!`, `[1]Sheet1!`, `'[Book 1.xlsx]Data'!`, `[1]!`.
 */
export function formatSheetQualifier(q: SheetQualifier): string {
	const sheets = q.sheetEnd !== undefined ? `${q.sheet}:${q.sheetEnd}` : q.sheet;
	let quote =
		(q.sheet !== '' && needsQuotes(q.sheet)) ||
		(q.sheetEnd !== undefined && needsQuotes(q.sheetEnd));
	let text = sheets;
	if (q.workbook !== undefined) {
		const slash = Math.max(q.workbook.lastIndexOf('\\'), q.workbook.lastIndexOf('/'));
		const path = q.workbook.slice(0, slash + 1);
		const book = q.workbook.slice(slash + 1);
		if (path !== '' || !/^[A-Za-z0-9_.]+$/.test(book)) {
			quote = true;
		}
		if (q.sheet === '' && q.sheetEnd === undefined) {
			text = path !== '' ? `${path}${book}` : `[${book}]`;
		} else {
			text = `${path}[${book}]${sheets}`;
		}
	}
	return quote ? `'${text.replace(/'/g, "''")}'!` : `${text}!`;
}

/** Full reference text: optional sheet prefix plus the area. */
export function formatReference(sheet: SheetQualifier | undefined, area: RefArea): string {
	return sheet === undefined
		? formatRefArea(area)
		: formatSheetQualifier(sheet) + formatRefArea(area);
}

/** Normalised range covered by an area (whole columns / rows span the grid). */
export function refAreaToRange(area: RefArea): RangeAddress {
	const { start, end } = area;
	return {
		top: Math.min(start.row, end.row),
		left: Math.min(start.col, end.col),
		bottom: Math.max(start.row, end.row),
		right: Math.max(start.col, end.col),
	};
}

/** Two sheet names are the same sheet (Excel compares case-insensitively). */
export function sameSheetName(a: string, b: string): boolean {
	return a === b || a.toLowerCase() === b.toLowerCase();
}

export interface ParsedReference {
	sheet?: SheetQualifier;
	area: RefArea;
}

/**
 * Parse a standalone reference such as `Sheet1!$A$1:$B$2`, `'My Sheet'!C:C`
 * or (with `r1c1`) `R1C1:R[2]C`, as used by INDIRECT and defined names.
 * Returns `null` unless the whole text is exactly one reference.
 */
export function parseReferenceText(
	text: string,
	options?: TokenizeOptions,
): ParsedReference | null {
	let found: ParsedReference | null = null;
	for (const token of tokenizeFormula(text.trim(), options)) {
		if (token.type === 'equals') {
			continue;
		}
		if (token.type !== 'ref' || found !== null || token.area === undefined) {
			return null;
		}
		found =
			token.sheet === undefined ? { area: token.area } : { sheet: token.sheet, area: token.area };
	}
	return found;
}
