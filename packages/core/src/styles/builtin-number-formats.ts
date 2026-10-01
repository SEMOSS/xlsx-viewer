/**
 * Built-in number formats (ids 0-49) as Excel displays them in the en-US
 * locale. Ids 5-8 and 37-44 are locale-dependent currency / accounting codes.
 */
export const BUILTIN_NUMBER_FORMATS: Readonly<Record<number, string>> = {
	0: 'General',
	1: '0',
	2: '0.00',
	3: '#,##0',
	4: '#,##0.00',
	5: '"$"#,##0_);\\("$"#,##0\\)',
	6: '"$"#,##0_);[Red]\\("$"#,##0\\)',
	7: '"$"#,##0.00_);\\("$"#,##0.00\\)',
	8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
	9: '0%',
	10: '0.00%',
	11: '0.00E+00',
	12: '# ?/?',
	13: '# ??/??',
	14: 'm/d/yyyy',
	15: 'd-mmm-yy',
	16: 'd-mmm',
	17: 'mmm-yy',
	18: 'h:mm AM/PM',
	19: 'h:mm:ss AM/PM',
	20: 'h:mm',
	21: 'h:mm:ss',
	22: 'm/d/yyyy h:mm',
	27: '[$-404]e/m/d',
	30: 'm/d/yy',
	36: '[$-404]e/m/d',
	37: '#,##0_);(#,##0)',
	38: '#,##0_);[Red](#,##0)',
	39: '#,##0.00_);(#,##0.00)',
	40: '#,##0.00_);[Red](#,##0.00)',
	41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
	42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
	43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
	44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
	45: 'mm:ss',
	46: '[h]:mm:ss',
	47: 'mmss.0',
	48: '##0.0E+0',
	49: '@',
	50: '[$-404]e/m/d',
	57: '[$-404]e/m/d',
	59: 't0',
	60: 't0.00',
	61: 't#,##0',
	62: 't#,##0.00',
	67: 't0%',
	68: 't0.00%',
	69: 't# ?/?',
	70: 't# ??/??',
};

/** First id available for custom number formats. */
export const FIRST_CUSTOM_NUMFMT_ID = 164;

/** Format code for an id, consulting custom formats first. Unknown ids fall back to General. */
export function numberFormatCode(id: number, custom: Readonly<Record<number, string>>): string {
	return custom[id] ?? BUILTIN_NUMBER_FORMATS[id] ?? 'General';
}

/** Built-in id whose code equals `code`, if any. */
export function builtinNumberFormatId(code: string): number | undefined {
	for (const [id, c] of Object.entries(BUILTIN_NUMBER_FORMATS)) {
		if (c === code) {
			return Number(id);
		}
	}
	return undefined;
}
