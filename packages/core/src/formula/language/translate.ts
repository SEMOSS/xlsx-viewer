import { offsetArea, referenceEdit } from './ref-edits';
import { applyEdits, type TextEdit } from './token-utils';
import { tokenizeFormula } from './tokenizer';

/**
 * Shift the relative references of a formula by (rowDelta, colDelta), as
 * when copying a formula, filling, or expanding a shared formula. Absolute
 * parts (`$A`, `$1`) stay; whole columns / rows only move along their own
 * axis. A reference with a corner pushed off the grid becomes `#REF!`
 * (keeping its sheet prefix). Strings, names and table references are not
 * touched and the formatting of the rest of the formula is preserved.
 * Never throws: on unexpected input the formula is returned unchanged.
 */
export function translateFormula(formula: string, rowDelta: number, colDelta: number): string {
	if ((rowDelta === 0 && colDelta === 0) || formula === '') {
		return formula;
	}
	try {
		const tokens = tokenizeFormula(formula);
		const edits: TextEdit[] = [];
		for (let i = 0; i < tokens.length; i++) {
			const t = tokens[i];
			if (t.type !== 'ref' || t.area === undefined) {
				continue;
			}
			const moved = offsetArea(t.area, rowDelta, colDelta);
			if (moved !== t.area) {
				edits.push(referenceEdit(tokens, i, moved));
			}
		}
		return applyEdits(formula, edits);
	} catch {
		return formula;
	}
}
