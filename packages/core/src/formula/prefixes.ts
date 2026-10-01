/**
 * File-format prefixes for newer functions. Excel stores functions added
 * after Excel 2007 with an `_xlfn.` prefix (FILTER and SORT use
 * `_xlfn._xlws.`), LET / LAMBDA parameters with `_xlpm.`, the spill
 * operator `A1#` as `_xlfn.ANCHORARRAY(A1)` and the implicit-intersection
 * operator `@x` as `_xlfn.SINGLE(x)`. The UI shows formulas without them.
 */
import { addPrefixesWithTokens, stripPrefixesWithTokens } from './language/prefix-rewrite';
import { addPrefixesTextually, stripPrefixesTextually } from './language/legacy-prefixes';

export { filePrefixFor, XLFN_FUNCTIONS, XLWS_FUNCTIONS } from './language/function-prefixes';
export { mapCodeSegments } from './language/legacy-prefixes';

/**
 * File form -> display form: remove `_xlfn.`, `_xlws.` and `_xlpm.`, turn
 * `_xlfn.ANCHORARRAY(A1)` into `A1#` and `_xlfn.SINGLE(x)` into `@x`.
 * Only the affected tokens change; strings and spacing are preserved.
 * Never throws.
 */
export function stripFormulaPrefixes(formula: string): string {
	if (!/_xl/i.test(formula)) {
		return formula;
	}
	try {
		return stripPrefixesWithTokens(formula) ?? stripPrefixesTextually(formula);
	} catch {
		return stripPrefixesTextually(formula);
	}
}

/**
 * Display form -> file form: prefix newer functions (`_xlfn.`,
 * `_xlfn._xlws.`), LET / LAMBDA parameters and their uses (`_xlpm.`), and
 * write `A1#` as `_xlfn.ANCHORARRAY(A1)`, `@x` as `_xlfn.SINGLE(x)` and the
 * `[@Col]` table shorthand as `[[#This Row],[Col]]`. Never throws.
 * Formulas without `(`, `#` or `@` need no prefix and return immediately,
 * so callers do not need a shortcut of their own.
 */
export function addFormulaPrefixes(formula: string): string {
	if (!/[(#@]/.test(formula)) {
		return formula;
	}
	try {
		return addPrefixesWithTokens(formula) ?? addPrefixesTextually(formula);
	} catch {
		return addPrefixesTextually(formula);
	}
}
