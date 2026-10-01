/**
 * Excel release that introduced each function added after Excel 2007, from
 * the "new functions" lists of each version. Functions that reached
 * Microsoft 365 after Excel 2021 (some were later bundled with Excel 2024)
 * are tagged `365`. Functions that are not listed predate Excel 2010.
 */

/** Versions in release order. */
export const CATALOG_VERSIONS: readonly string[] = ['2010', '2013', '2016', '2019', '2021', '365'];

// prettier-ignore
const ADDED_IN: ReadonlyArray<readonly [string, readonly string[]]> = [
	['2010', [
		'AGGREGATE', 'BETA.DIST', 'BETA.INV', 'BINOM.DIST', 'BINOM.INV', 'CEILING.PRECISE',
		'CHISQ.DIST', 'CHISQ.DIST.RT', 'CHISQ.INV', 'CHISQ.INV.RT', 'CHISQ.TEST', 'CONFIDENCE.NORM',
		'CONFIDENCE.T', 'COVARIANCE.P', 'COVARIANCE.S', 'ERF.PRECISE', 'ERFC.PRECISE', 'EXPON.DIST',
		'F.DIST', 'F.DIST.RT', 'F.INV', 'F.INV.RT', 'F.TEST', 'FLOOR.PRECISE', 'GAMMA.DIST',
		'GAMMA.INV', 'GAMMALN.PRECISE', 'HYPGEOM.DIST', 'ISO.CEILING', 'LOGNORM.DIST', 'LOGNORM.INV',
		'MODE.MULT', 'MODE.SNGL', 'NEGBINOM.DIST', 'NETWORKDAYS.INTL', 'NORM.DIST', 'NORM.INV',
		'NORM.S.DIST', 'NORM.S.INV', 'PERCENTILE.EXC', 'PERCENTILE.INC', 'PERCENTRANK.EXC',
		'PERCENTRANK.INC', 'POISSON.DIST', 'QUARTILE.EXC', 'QUARTILE.INC', 'RANK.AVG', 'RANK.EQ',
		'STDEV.P', 'STDEV.S', 'T.DIST', 'T.DIST.2T', 'T.DIST.RT', 'T.INV', 'T.INV.2T', 'T.TEST',
		'VAR.P', 'VAR.S', 'WEIBULL.DIST', 'WORKDAY.INTL', 'Z.TEST',
	]],
	['2013', [
		'ACOT', 'ACOTH', 'ARABIC', 'BASE', 'BINOM.DIST.RANGE', 'BITAND', 'BITLSHIFT', 'BITOR',
		'BITRSHIFT', 'BITXOR', 'CEILING.MATH', 'COMBINA', 'COT', 'COTH', 'CSC', 'CSCH', 'DAYS',
		'DBCS', 'DECIMAL', 'ENCODEURL', 'FILTERXML', 'FLOOR.MATH', 'FORMULATEXT', 'GAMMA', 'GAUSS',
		'IFNA', 'IMCOSH', 'IMCOT', 'IMCSC', 'IMCSCH', 'IMSEC', 'IMSECH', 'IMSINH', 'IMTAN',
		'ISFORMULA', 'ISOWEEKNUM', 'MUNIT', 'NUMBERVALUE', 'PDURATION', 'PERMUTATIONA', 'PHI', 'RRI',
		'SEC', 'SECH', 'SHEET', 'SHEETS', 'SKEW.P', 'UNICHAR', 'UNICODE', 'WEBSERVICE', 'XOR',
	]],
	['2016', [
		'FORECAST.ETS', 'FORECAST.ETS.CONFINT', 'FORECAST.ETS.SEASONALITY', 'FORECAST.ETS.STAT',
		'FORECAST.LINEAR',
	]],
	['2019', [
		'CONCAT', 'IFS', 'MAXIFS', 'MINIFS', 'SWITCH', 'TEXTJOIN',
	]],
	['2021', [
		'FILTER', 'LET', 'RANDARRAY', 'SEQUENCE', 'SORT', 'SORTBY', 'UNIQUE', 'XLOOKUP', 'XMATCH',
	]],
	['365', [
		'ARRAYTOTEXT', 'BYCOL', 'BYROW', 'CHOOSECOLS', 'CHOOSEROWS', 'DETECTLANGUAGE', 'DROP',
		'EXPAND', 'FIELDVALUE', 'GROUPBY', 'HSTACK', 'IMAGE', 'ISOMITTED', 'LAMBDA', 'MAKEARRAY',
		'MAP', 'PERCENTOF', 'PIVOTBY', 'REDUCE', 'REGEXEXTRACT', 'REGEXREPLACE', 'REGEXTEST', 'SCAN',
		'STOCKHISTORY', 'TAKE', 'TEXTAFTER', 'TEXTBEFORE', 'TEXTSPLIT', 'TOCOL', 'TOROW', 'TRANSLATE',
		'TRIMRANGE', 'VALUETOTEXT', 'VSTACK', 'WRAPCOLS', 'WRAPROWS',
	]],
];

const SINCE = new Map<string, string>();
for (const [version, names] of ADDED_IN) {
	for (const name of names) {
		SINCE.set(name, version);
	}
}

/** Version that introduced `name` (upper case), or undefined for Excel 2007 functions. */
export function addedIn(name: string): string | undefined {
	return SINCE.get(name);
}

/** Every `[name, version]` pair of the table, in table order (duplicates included). */
export function versionTable(): Array<[string, string]> {
	return ADDED_IN.flatMap(([version, names]) =>
		names.map((name): [string, string] => [name, version]),
	);
}
