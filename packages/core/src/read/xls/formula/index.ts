/**
 * BIFF8 formulas: token decoding (`ptg-decode`), the function tables and
 * the conversion to formula text in display form (no leading `=`, no file
 * prefixes). Every rendered formula is checked with the formula parser.
 */
import { parseFormula } from '../../../formula/language/parser';
import { serializeFormula } from '../../../formula/language/serializer';
import { type FormulaEnv, type FormulaPos, ptgsToAst } from './ptg-ast';
import { type Ptg, XlsFormulaError } from './ptg-decode';

export { CETAB } from './cetab';
export { FTAB, type FtabEntry, UDF_INDEX } from './ftab';
export {
	type FormulaEnv,
	type FormulaPos,
	ptgsToAst,
	type SupBook,
	XlsExternalReferenceError,
	type Xti,
} from './ptg-ast';
export {
	type ArrayValue,
	decodeRgce,
	ERROR_CODES,
	errorCode,
	type Ptg,
	type PtgLoc,
	XlsFormulaError,
} from './ptg-decode';

/**
 * Formula text for decoded tokens at `pos`. Throws {@link XlsFormulaError}
 * when the tokens are malformed or the result would not parse.
 */
export function renderFormula(ptgs: readonly Ptg[], env: FormulaEnv, pos: FormulaPos): string {
	const text = serializeFormula(ptgsToAst(ptgs, env, pos));
	try {
		parseFormula(text);
	} catch (e) {
		const reason = e instanceof Error ? e.message : String(e);
		throw new XlsFormulaError(`The formula "${text}" is not valid (${reason})`);
	}
	return text;
}
