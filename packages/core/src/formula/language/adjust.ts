import {
	adjustAreaForStructureChange,
	type FormulaAdjustment,
	type StructureChange,
} from './adjust-range';
import { isLocalReference, referenceEdit, referenceSheet } from './ref-edits';
import { sameSheetName } from './references';
import { applyEdits, type TextEdit } from './token-utils';
import { tokenizeFormula } from './tokenizer';

/**
 * Rewrite the references of a formula after rows or columns were inserted
 * or deleted on `change.sheet`, the way Excel does:
 *
 * - relative and absolute references both move;
 * - inserting at or before a reference shifts it, inserting inside a range
 *   (after its first line) expands it;
 * - deleting removes references inside the deleted band (`#REF!`), shrinks
 *   ranges that overlap it and shifts references after it;
 * - whole columns ignore row changes and whole rows ignore column changes.
 *
 * Unqualified references belong to `formulaSheet`. External references and
 * 3-D references (`Sheet1:Sheet3!A1`) are left alone, as Excel does when a
 * single sheet of the span changes. Names, tables and strings are not touched.
 */
export function adjustFormulaForStructureChange(
	formula: string,
	change: StructureChange,
	formulaSheet: string,
): FormulaAdjustment {
	if (change.count <= 0 || formula === '') {
		return { formula, changed: false };
	}
	const tokens = tokenizeFormula(formula);
	const edits: TextEdit[] = [];
	for (let i = 0; i < tokens.length; i++) {
		const t = tokens[i];
		if (!isLocalReference(t) || !sameSheetName(referenceSheet(t, formulaSheet), change.sheet)) {
			continue;
		}
		const adjusted = adjustAreaForStructureChange(t.area, change);
		if (adjusted !== t.area) {
			edits.push(referenceEdit(tokens, i, adjusted));
		}
	}
	if (edits.length === 0) {
		return { formula, changed: false };
	}
	return { formula: applyEdits(formula, edits), changed: true };
}
