/**
 * File-format prefixes for newer functions. Excel stores functions added
 * after Excel 2007 with an `_xlfn.` prefix (FILTER and SORT use
 * `_xlfn._xlws.`) and LET / LAMBDA parameters with `_xlpm.`. The UI shows
 * formulas without them.
 */

/** Functions stored as `_xlfn.NAME`. */
// prettier-ignore
export const XLFN_FUNCTIONS: ReadonlySet<string> = new Set([
	'ACOT', 'ACOTH', 'AGGREGATE', 'ARABIC', 'BASE', 'BETA.DIST', 'BETA.INV', 'BINOM.DIST', 'BINOM.DIST.RANGE', 'BINOM.INV',
	'BITAND', 'BITLSHIFT', 'BITOR', 'BITRSHIFT', 'BITXOR', 'CEILING.MATH', 'CEILING.PRECISE', 'CHISQ.DIST',
	'CHISQ.DIST.RT', 'CHISQ.INV', 'CHISQ.INV.RT', 'CHISQ.TEST', 'COMBINA', 'CONCAT', 'CONFIDENCE.NORM',
	'CONFIDENCE.T', 'COT', 'COTH', 'COVARIANCE.P', 'COVARIANCE.S', 'CSC', 'CSCH', 'DAYS', 'DECIMAL', 'ENCODEURL', 'FIELDVALUE',
	'ECMA.CEILING', 'ERF.PRECISE', 'ERFC.PRECISE', 'EXPON.DIST', 'F.DIST', 'F.DIST.RT', 'F.INV', 'F.INV.RT',
	'F.TEST', 'FILTERXML', 'FLOOR.MATH', 'FLOOR.PRECISE', 'FORECAST.ETS', 'FORECAST.ETS.CONFINT',
	'FORECAST.ETS.SEASONALITY', 'FORECAST.ETS.STAT', 'FORECAST.LINEAR', 'FORMULATEXT', 'GAMMA', 'GAMMA.DIST',
	'GAMMA.INV', 'GAMMALN.PRECISE', 'GAUSS', 'HYPGEOM.DIST', 'IFNA', 'IFS', 'IMCOSH', 'IMCOT', 'IMCSC',
	'IMCSCH', 'IMSEC', 'IMSECH', 'IMSINH', 'IMTAN', 'ISFORMULA', 'ISOWEEKNUM', 'LOGNORM.DIST',
	'LOGNORM.INV', 'MAXIFS', 'MINIFS', 'MODE.MULT', 'MODE.SNGL', 'MUNIT', 'NEGBINOM.DIST',
	'NETWORKDAYS.INTL', 'NORM.DIST', 'NORM.INV', 'NORM.S.DIST', 'NORM.S.INV', 'NUMBERVALUE', 'PDURATION',
	'PERCENTILE.EXC', 'PERCENTILE.INC', 'PERCENTRANK.EXC', 'PERCENTRANK.INC', 'PERMUTATIONA', 'PHI',
	'POISSON.DIST', 'QUARTILE.EXC', 'QUARTILE.INC', 'QUERYSTRING', 'RANK.AVG', 'RANK.EQ', 'RRI', 'SEC',
	'SECH', 'SHEET', 'SHEETS', 'SKEW.P', 'STDEV.P', 'STDEV.S', 'SWITCH', 'T.DIST', 'T.DIST.2T', 'T.DIST.RT',
	'T.INV', 'T.INV.2T', 'T.TEST', 'TEXTJOIN', 'UNICHAR', 'UNICODE', 'VAR.P', 'VAR.S', 'WEBSERVICE',
	'WEIBULL.DIST', 'WORKDAY.INTL', 'XOR', 'Z.TEST',
	'XLOOKUP', 'XMATCH', 'SEQUENCE', 'RANDARRAY', 'UNIQUE', 'SORTBY', 'LET', 'LAMBDA', 'MAP', 'REDUCE',
	'SCAN', 'MAKEARRAY', 'BYROW', 'BYCOL', 'ISOMITTED', 'CHOOSECOLS', 'CHOOSEROWS', 'DROP', 'EXPAND',
	'HSTACK', 'VSTACK', 'TAKE', 'TEXTAFTER', 'TEXTBEFORE', 'TEXTSPLIT', 'TOCOL', 'TOROW', 'WRAPCOLS',
	'WRAPROWS', 'VALUETOTEXT', 'ARRAYTOTEXT', 'IMAGE', 'STOCKHISTORY', 'GROUPBY', 'PIVOTBY', 'PERCENTOF',
	'REGEXTEST', 'REGEXEXTRACT', 'REGEXREPLACE', 'TRIMRANGE', 'ANCHORARRAY', 'SINGLE',
]);

/** Functions stored as `_xlfn._xlws.NAME`. */
export const XLWS_FUNCTIONS: ReadonlySet<string> = new Set(['FILTER', 'SORT']);

/** File prefix required for a function name (upper case), or `''`. */
export function filePrefixFor(name: string): string {
	const upper = name.toUpperCase();
	if (XLWS_FUNCTIONS.has(upper)) {
		return '_xlfn._xlws.';
	}
	if (XLFN_FUNCTIONS.has(upper)) {
		return '_xlfn.';
	}
	return '';
}
