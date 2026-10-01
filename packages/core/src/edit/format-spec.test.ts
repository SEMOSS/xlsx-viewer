import { describe, expect, it } from 'vitest';

import { buildNumberFormat } from '../numfmt/builder';
import {
	CURRENCY_SYMBOL_PRESETS,
	DATE_FORMAT_PRESETS,
	FRACTION_FORMAT_PRESETS,
	NEGATIVE_NUMBER_STYLES,
	SPECIAL_FORMAT_PRESETS,
	TIME_FORMAT_PRESETS,
} from '../numfmt/presets';
import {
	customNumberFormatList,
	describeNumberFormat,
	formatCodeKey,
	sameFormatCode,
} from './format-spec';

describe('formatCodeKey / sameFormatCode', () => {
	it('ignores escapes, quotes, the en-US locale tag and a trailing text section', () => {
		expect(sameFormatCode('"$"#,##0.00', '\\$#,##0.00')).toBe(true);
		expect(sameFormatCode('[$-409]h:mm AM/PM;@', 'h:mm AM/PM')).toBe(true);
		expect(sameFormatCode('0.00', '0.000')).toBe(false);
		expect(formatCodeKey('[$-F800]dddd')).toBe('[$-F800]dddd');
	});
});

describe('describeNumberFormat', () => {
	it('recognises General and Text', () => {
		expect(describeNumberFormat('General')).toEqual({ category: 'general' });
		expect(describeNumberFormat('@')).toEqual({ category: 'text', code: '@' });
	});

	it('round-trips every Number option the dialog can build', () => {
		for (let decimals = 0; decimals <= 4; decimals++) {
			for (const useThousands of [false, true]) {
				for (const n of NEGATIVE_NUMBER_STYLES) {
					const spec = {
						category: 'number' as const,
						decimals,
						useThousands,
						negativeStyle: n.style,
					};
					expect(describeNumberFormat(buildNumberFormat(spec))).toEqual(spec);
				}
			}
		}
	});

	it('round-trips Currency and Accounting with every symbol', () => {
		for (const symbol of CURRENCY_SYMBOL_PRESETS) {
			for (const n of NEGATIVE_NUMBER_STYLES) {
				const spec = {
					category: 'currency' as const,
					decimals: 2,
					currencySymbol: symbol.id,
					negativeStyle: n.style,
				};
				const back = describeNumberFormat(buildNumberFormat(spec));
				expect(buildNumberFormat(back)).toBe(buildNumberFormat(spec));
				// Without a symbol the code is a Number format, which is how Excel reopens it.
				expect(back.category).toBe(symbol.id === 'none' ? 'number' : 'currency');
			}
			const acc = { category: 'accounting' as const, decimals: 1, currencySymbol: symbol.id };
			const back = describeNumberFormat(buildNumberFormat(acc));
			expect(back.category).toBe('accounting');
			expect(buildNumberFormat(back)).toBe(buildNumberFormat(acc));
		}
	});

	it('maps built-in codes to their categories', () => {
		expect(describeNumberFormat('#,##0.00')).toMatchObject({
			category: 'number',
			decimals: 2,
			useThousands: true,
			negativeStyle: 'minus',
		});
		expect(describeNumberFormat('0%')).toEqual({ category: 'percentage', decimals: 0 });
		expect(describeNumberFormat('0.00E+00')).toEqual({ category: 'scientific', decimals: 2 });
		expect(describeNumberFormat('"$"#,##0.00_);[Red]\\("$"#,##0.00\\)')).toMatchObject({
			category: 'currency',
			decimals: 2,
			currencySymbol: '$',
			negativeStyle: 'redParens',
		});
		expect(
			describeNumberFormat('_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)'),
		).toMatchObject({
			category: 'accounting',
			decimals: 2,
		});
	});

	it('selects the matching type of date, time, fraction and special presets', () => {
		expect(describeNumberFormat('m/d/yyyy')).toEqual({
			category: 'date',
			code: DATE_FORMAT_PRESETS[0].code,
		});
		expect(describeNumberFormat('h:mm AM/PM')).toEqual({
			category: 'time',
			code: '[$-409]h:mm AM/PM;@',
		});
		expect(describeNumberFormat('h:mm')).toEqual({ category: 'time', code: 'h:mm;@' });
		for (const p of [...DATE_FORMAT_PRESETS, ...TIME_FORMAT_PRESETS]) {
			expect(['date', 'time']).toContain(describeNumberFormat(p.code).category);
		}
		for (const p of FRACTION_FORMAT_PRESETS) {
			expect(describeNumberFormat(p.code)).toEqual({ category: 'fraction', code: p.code });
		}
		for (const p of SPECIAL_FORMAT_PRESETS) {
			expect(describeNumberFormat(p.code)).toEqual({ category: 'special', code: p.code });
		}
	});

	it('falls back to Custom for codes the categories cannot produce', () => {
		expect(describeNumberFormat('0.0,,"M"')).toEqual({ category: 'custom', code: '0.0,,"M"' });
		expect(describeNumberFormat('[Blue]0.00')).toEqual({ category: 'custom', code: '[Blue]0.00' });
		expect(describeNumberFormat('yyyy-mm-dd')).toEqual({ category: 'custom', code: 'yyyy-mm-dd' });
	});

	it('treats unescaped parentheses like the escaped ones the dialog writes', () => {
		expect(describeNumberFormat('#,##0.00_);(#,##0.00)')).toMatchObject({
			category: 'number',
			negativeStyle: 'parens',
		});
	});
});

describe('customNumberFormatList', () => {
	it('lists the built-in codes first, then the workbook codes without duplicates', () => {
		const list = customNumberFormatList({
			numFmts: { 165: 'yyyy-mm-dd', 164: '0.0%', 166: '0.00' },
		});
		expect(list[0]).toBe('General');
		expect(list).toContain('[h]:mm:ss');
		expect(list.slice(-2)).toEqual(['0.0%', 'yyyy-mm-dd']);
		expect(list.filter((c) => c === '0.00')).toHaveLength(1);
	});
});
