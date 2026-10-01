import type { SheetQualifier } from '../ast';
import { formatSheetQualifier, sameSheetName } from './references';
import { applyEdits, type TextEdit } from './token-utils';
import { tokenizeFormula } from './tokenizer';
import type { FormulaToken } from './tokens';

/** Qualifier of a token when it names a sheet of this workbook. */
function localQualifier(t: FormulaToken): SheetQualifier | undefined {
	return t.sheet !== undefined && t.sheet.workbook === undefined ? t.sheet : undefined;
}

function mentions(q: SheetQualifier, name: string): boolean {
	return (
		sameSheetName(q.sheet, name) || (q.sheetEnd !== undefined && sameSheetName(q.sheetEnd, name))
	);
}

/**
 * Rename a sheet in every sheet-qualified reference and name of a formula
 * (`Sheet1!A1`, `'Sheet 1:Sheet3'!A1`, `Sheet1!MyName`, `Sheet1!#REF!`),
 * quoting the new name where needed. Strings (INDIRECT text) and external
 * references are not touched.
 */
export function renameSheetInFormula(formula: string, oldName: string, newName: string): string {
	if (formula.indexOf('!') < 0) {
		return formula;
	}
	const edits: TextEdit[] = [];
	for (const t of tokenizeFormula(formula)) {
		const q = localQualifier(t);
		if (q === undefined || !mentions(q, oldName)) {
			continue;
		}
		const renamed: SheetQualifier = { sheet: sameSheetName(q.sheet, oldName) ? newName : q.sheet };
		if (q.sheetEnd !== undefined) {
			renamed.sheetEnd = sameSheetName(q.sheetEnd, oldName) ? newName : q.sheetEnd;
		}
		edits.push({
			start: t.start,
			end: t.start + (t.qualifierLength ?? 0),
			text: formatSheetQualifier(renamed),
		});
	}
	return applyEdits(formula, edits);
}

/** The formula names `name` in a sheet prefix (directly or as a 3-D endpoint). */
export function formulaReferencesSheet(formula: string, name: string): boolean {
	if (formula.indexOf('!') < 0) {
		return false;
	}
	for (const t of tokenizeFormula(formula)) {
		const q = localQualifier(t);
		if (q !== undefined && mentions(q, name)) {
			return true;
		}
	}
	return false;
}

/**
 * After deleting sheet `name`, turn every reference to it into `#REF!`
 * (`=Gone!A1+1` -> `=#REF!+1`). A 3-D reference keeps working unless the
 * deleted sheet was its only sheet: Excel then narrows the span, which
 * needs the sheet order and is left to the caller.
 */
export function invalidateSheetInFormula(formula: string, name: string): string {
	if (formula.indexOf('!') < 0) {
		return formula;
	}
	const tokens = tokenizeFormula(formula);
	const edits: TextEdit[] = [];
	for (let i = 0; i < tokens.length; i++) {
		const t = tokens[i];
		const q = localQualifier(t);
		if (q === undefined || !sameSheetName(q.sheet, name)) {
			continue;
		}
		if (q.sheetEnd !== undefined && !sameSheetName(q.sheetEnd, name)) {
			continue;
		}
		const next = tokens[i + 1];
		const end = next !== undefined && next.type === 'spill' ? next.end : t.end;
		edits.push({ start: t.start, end, text: '#REF!' });
	}
	return applyEdits(formula, edits);
}
