import { quoteSheetName } from '../address/sheet-name';
import { isNumericLiteral } from '../conditional/formula-eval';
import { parseFormula } from '../formula/language/parser';
import { applyEdits, type TextEdit } from '../formula/language/token-utils';
import { tokenizeFormula } from '../formula/language/tokenizer';

/** Excel's message for a "Refers to" entry that is not a valid formula. */
export const REFERS_TO_ERROR_MESSAGE = "There's a problem with this formula.";

/**
 * Prefix every unqualified cell reference with `sheetName!`, the way
 * Excel completes a "Refers to" entry such as `=$A$1:$B$2` to
 * `=Sheet1!$A$1:$B$2`. Names, strings and qualified references stay.
 */
export function qualifyFormulaReferences(formula: string, sheetName: string): string {
	const prefix = `${quoteSheetName(sheetName)}!`;
	const edits: TextEdit[] = [];
	for (const t of tokenizeFormula(formula)) {
		if (t.type === 'ref' && t.sheet === undefined) {
			edits.push({ start: t.start, end: t.start, text: prefix });
		}
	}
	return applyEdits(formula, edits);
}

/**
 * Turn the "Refers to" text of the name dialogs into a stored formula (no
 * leading `=`). `=...` must be a valid formula; text without `=` becomes a
 * constant like in Excel (`5` stays `5`, `abc` becomes `"abc"`). With
 * `sheetName`, unqualified references are qualified with that sheet.
 */
export function parseRefersTo(
	text: string,
	sheetName?: string,
): { formula: string } | { error: string } {
	const trimmed = text.trim();
	if (trimmed === '') {
		return { error: REFERS_TO_ERROR_MESSAGE };
	}
	if (!trimmed.startsWith('=')) {
		if (isNumericLiteral(trimmed)) {
			return { formula: trimmed };
		}
		if (/^(?:true|false)$/i.test(trimmed)) {
			return { formula: trimmed.toUpperCase() };
		}
		return { formula: `"${trimmed.replace(/"/g, '""')}"` };
	}
	const formula = trimmed.slice(1).trim();
	if (formula === '') {
		return { error: REFERS_TO_ERROR_MESSAGE };
	}
	try {
		parseFormula(formula);
	} catch {
		return { error: REFERS_TO_ERROR_MESSAGE };
	}
	return {
		formula: sheetName === undefined ? formula : qualifyFormulaReferences(formula, sheetName),
	};
}
