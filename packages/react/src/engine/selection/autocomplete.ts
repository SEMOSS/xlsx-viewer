import type { Worksheet } from 'xlsx-viewer-core';
import { cellKey } from 'xlsx-viewer-core';

/** Cells scanned up and down the column at most. */
const MAX_SCAN = 10_000;

/**
 * Excel's AutoComplete for cell values: the single text entry of the
 * column's list (the contiguous cells above and below, up to a blank) that
 * starts with what was typed, case-insensitively. Null when there is no
 * match, several different matches, or the entry is a number or formula.
 */
export function autoCompleteCandidate(
	sheet: Worksheet,
	row: number,
	col: number,
	typed: string,
): string | null {
	if (
		typed === '' ||
		typed.startsWith('=') ||
		typed.includes('\n') ||
		/^[\d\s.,+\-/:%$()]+$/.test(typed)
	) {
		return null;
	}
	const prefix = typed.toLowerCase();
	let found: string | null = null;
	const visit = (r: number): 'stop' | 'next' | 'conflict' => {
		const v = sheet.cells.get(cellKey(r, col))?.value;
		if (v === null || v === undefined || v === '') {
			return 'stop';
		}
		if (typeof v === 'string' && v.length > typed.length && v.toLowerCase().startsWith(prefix)) {
			if (found === null) {
				found = v;
			} else if (found.toLowerCase() !== v.toLowerCase()) {
				return 'conflict';
			}
		}
		return 'next';
	};
	for (const dir of [-1, 1]) {
		for (let i = 1, r = row + dir; i <= MAX_SCAN && r >= 0; i++, r += dir) {
			const step = visit(r);
			if (step === 'conflict') {
				return null;
			}
			if (step === 'stop') {
				break;
			}
		}
	}
	return found;
}
