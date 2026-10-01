import { rangeContainsRange } from '../../address/range';
import type { RangeAddress } from '../../types/address';
import type { FormulaAdjustment } from './adjust-range';
import { isLocalReference, offsetArea, referenceEdit, referenceSheet } from './ref-edits';
import { formatSheetQualifier, refAreaToRange, sameSheetName } from './references';
import { applyEdits, type TextEdit } from './token-utils';
import { tokenizeFormula } from './tokenizer';

/** Cells cut from a sheet. */
export interface FormulaMoveSource {
	sheet: string;
	range: RangeAddress;
}

/** Top-left cell where the cut cells are pasted. */
export interface FormulaMoveTarget {
	sheet: string;
	row: number;
	col: number;
}

export interface FormulaMoveOptions {
	/**
	 * Sheet the formula itself lives on after the move, when the formula
	 * cell is part of the moved block and lands on another sheet. Its
	 * unqualified references to cells that stay behind get a sheet prefix.
	 */
	formulaTargetSheet?: string;
}

/**
 * Rewrite a formula for a cut / paste (move) of `source` to `target`, with
 * Excel semantics:
 *
 * - references entirely inside the moved block follow it (absolute parts
 *   too), gaining a sheet prefix when the block lands on another sheet;
 * - references entirely inside the destination block (and not moved) become
 *   `#REF!`, because the pasted cells replace those cells;
 * - every other reference is unchanged (partial overlaps included).
 *
 * Call it for every formula in the workbook, including the moved ones (a
 * moved formula keeps pointing at the same cells unless those moved too).
 */
export function adjustFormulaForMove(
	formula: string,
	formulaSheet: string,
	source: FormulaMoveSource,
	target: FormulaMoveTarget,
	options: FormulaMoveOptions = {},
): FormulaAdjustment {
	if (formula === '') {
		return { formula, changed: false };
	}
	const src = source.range;
	const dr = target.row - src.top;
	const dc = target.col - src.left;
	const dest: RangeAddress = {
		top: target.row,
		left: target.col,
		bottom: target.row + (src.bottom - src.top),
		right: target.col + (src.right - src.left),
	};
	const home = options.formulaTargetSheet ?? formulaSheet;
	const tokens = tokenizeFormula(formula);
	const edits: TextEdit[] = [];
	for (let i = 0; i < tokens.length; i++) {
		const t = tokens[i];
		if (!isLocalReference(t)) {
			continue;
		}
		const refSheet = referenceSheet(t, formulaSheet);
		const range = refAreaToRange(t.area);
		const written = t.text.slice(0, t.qualifierLength ?? 0);
		if (sameSheetName(refSheet, source.sheet) && rangeContainsRange(src, range)) {
			const moved = offsetArea(t.area, dr, dc, true);
			const crossSheet = !sameSheetName(refSheet, target.sheet);
			let prefix = written;
			if (crossSheet) {
				prefix =
					t.sheet === undefined && sameSheetName(target.sheet, home)
						? ''
						: formatSheetQualifier({ sheet: target.sheet });
			} else if (t.sheet === undefined && !sameSheetName(refSheet, home)) {
				prefix = formatSheetQualifier({ sheet: refSheet });
			}
			if (moved !== t.area || prefix !== written) {
				edits.push(referenceEdit(tokens, i, moved, prefix));
			}
			continue;
		}
		if (sameSheetName(refSheet, target.sheet) && rangeContainsRange(dest, range)) {
			edits.push(referenceEdit(tokens, i, null));
			continue;
		}
		if (t.sheet === undefined && !sameSheetName(refSheet, home)) {
			edits.push(referenceEdit(tokens, i, t.area, formatSheetQualifier({ sheet: refSheet })));
		}
	}
	if (edits.length === 0) {
		return { formula, changed: false };
	}
	return { formula: applyEdits(formula, edits), changed: true };
}
