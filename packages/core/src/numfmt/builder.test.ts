import { describe, expect, it } from 'vitest';

import { BUILTIN_NUMBER_FORMATS as B } from '../styles/builtin-number-formats';
import {
	buildNumberFormat,
	classifyFormat,
	CURRENCY_SYMBOL_PRESETS,
	DATE_FORMAT_PRESETS,
	formatValue,
	FRACTION_FORMAT_PRESETS,
	NEGATIVE_NUMBER_STYLES,
	SPECIAL_FORMAT_PRESETS,
	TIME_FORMAT_PRESETS,
} from './index';
import type { NumberFormatSpec } from './index';

/** Wednesday, March 14, 2012 1:30:55.2 PM, the Format Cells dialog sample. */
const TIME_OF_DAY = (13 * 3600 + 30 * 60 + 55.2) / 86_400;
const SAMPLE = 40982 + TIME_OF_DAY;

describe('buildNumberFormat', () => {
	it.each<[NumberFormatSpec, string]>([
		[{ category: 'general' }, 'General'],
		[{ category: 'number' }, '0.00'],
		[{ category: 'number', decimals: 0 }, '0'],
		[{ category: 'number', decimals: 2, useThousands: true }, '#,##0.00'],
		[{ category: 'number', negativeStyle: 'red' }, '0.00;[Red]0.00'],
		[{ category: 'number', negativeStyle: 'parens' }, '0.00_);\\(0.00\\)'],
		[
			{ category: 'number', useThousands: true, negativeStyle: 'redParens' },
			'#,##0.00_);[Red]\\(#,##0.00\\)',
		],
		[{ category: 'number', negativeStyle: 'redMinus' }, '0.00_ ;[Red]\\-0.00\\ '],
		[{ category: 'number', decimals: 99 }, '0.' + '0'.repeat(30)],
		[{ category: 'currency' }, '"$"#,##0.00'],
		[{ category: 'currency', negativeStyle: 'parens' }, B[7]],
		[{ category: 'currency', negativeStyle: 'redParens' }, B[8]],
		[{ category: 'currency', decimals: 0, negativeStyle: 'parens' }, B[5]],
		[{ category: 'currency', decimals: 0, negativeStyle: 'redParens' }, B[6]],
		[{ category: 'currency', negativeStyle: 'red' }, '"$"#,##0.00;[Red]"$"#,##0.00'],
		[{ category: 'currency', currencySymbol: '€' }, '[$€-2]\\ #,##0.00'],
		[{ category: 'currency', currencySymbol: '€-after' }, '#,##0.00\\ [$€-1]'],
		[{ category: 'currency', currencySymbol: '£', decimals: 0 }, '[$£-809]#,##0'],
		[{ category: 'currency', currencySymbol: 'none' }, '#,##0.00'],
		[{ category: 'currency', currencySymbol: '' }, '#,##0.00'],
		[{ category: 'currency', currencySymbol: 'kr', symbolPosition: 'after' }, '#,##0.00\\ "kr"'],
		[{ category: 'currency', currencySymbol: 'Br' }, '"Br"#,##0.00'],
		[{ category: 'accounting' }, B[44]],
		[{ category: 'accounting', decimals: 0 }, B[42]],
		[{ category: 'accounting', currencySymbol: 'none' }, B[43]],
		[{ category: 'accounting', currencySymbol: 'none', decimals: 0 }, B[41]],
		[
			{ category: 'accounting', currencySymbol: '€' },
			'_([$€-2]\\ * #,##0.00_);_([$€-2]\\ * \\(#,##0.00\\);_([$€-2]\\ * "-"??_);_(@_)',
		],
		[{ category: 'percentage' }, '0.00%'],
		[{ category: 'percentage', decimals: 0 }, '0%'],
		[{ category: 'scientific' }, '0.00E+00'],
		[{ category: 'scientific', decimals: 0 }, '0E+00'],
		[{ category: 'fraction' }, '# ?/?'],
		[{ category: 'fraction', fractionDigits: 2 }, '# ??/??'],
		[{ category: 'fraction', fractionDenominator: 16 }, '# ??/16'],
		[{ category: 'fraction', fractionDenominator: 8 }, '# ?/8'],
		[{ category: 'fraction', fractionDenominator: 100 }, '# ??/100'],
		[{ category: 'date' }, 'm/d/yyyy'],
		[{ category: 'date', code: '[$-409]d-mmm-yy;@' }, '[$-409]d-mmm-yy;@'],
		[{ category: 'time' }, '[$-F400]h:mm:ss\\ AM/PM'],
		[{ category: 'text' }, '@'],
		[{ category: 'special' }, '00000'],
		[{ category: 'custom', code: '0 "kg"' }, '0 "kg"'],
		[{ category: 'custom' }, 'General'],
	])('%j -> %s', (spec, code) => {
		expect(buildNumberFormat(spec)).toBe(code);
	});

	it('produces codes classified back into their category', () => {
		const specs: NumberFormatSpec[] = [
			{ category: 'general' },
			{ category: 'number', useThousands: true, negativeStyle: 'redMinus' },
			{ category: 'currency', currencySymbol: '€-after', negativeStyle: 'red' },
			{ category: 'accounting', currencySymbol: '£' },
			{ category: 'percentage', decimals: 1 },
			{ category: 'scientific', decimals: 3 },
			{ category: 'fraction', fractionDenominator: 4 },
			{ category: 'date' },
			{ category: 'time' },
			{ category: 'text' },
			{ category: 'special', code: '000-00-0000' },
		];
		for (const spec of specs) {
			expect(classifyFormat(buildNumberFormat(spec))).toBe(spec.category);
		}
	});

	it('renders the accounting and currency options as Excel does', () => {
		const euro = buildNumberFormat({ category: 'currency', currencySymbol: '€' });
		expect(formatValue(1234.5, euro).text).toBe('€ 1,234.50');
		const after = buildNumberFormat({ category: 'currency', currencySymbol: '€-after' });
		expect(formatValue(-1234.5, after).text).toBe('-1,234.50 €');
		const redMinus = buildNumberFormat({ category: 'number', negativeStyle: 'redMinus' });
		expect(formatValue(-1234.1, redMinus)).toEqual({ text: '-1234.10 ', color: '#FF0000' });
		expect(formatValue(1234.1, redMinus).text).toBe('1234.10 ');
	});
});

describe('Format Cells presets', () => {
	it.each(DATE_FORMAT_PRESETS.map((p) => [p.code, p.example]))('date %s -> %s', (code, example) => {
		expect(formatValue(SAMPLE, code).text).toBe(example);
		expect(classifyFormat(code)).toBe('date');
	});

	it.each(TIME_FORMAT_PRESETS.map((p) => [p.code, p.example]))('time %s -> %s', (code, example) => {
		const value = code.startsWith('[h]') ? 1 + TIME_OF_DAY : SAMPLE;
		expect(formatValue(value, code).text).toBe(example);
		expect(['time', 'date']).toContain(classifyFormat(code));
	});

	it('lists fraction and special presets of their categories', () => {
		for (const p of FRACTION_FORMAT_PRESETS) {
			expect(classifyFormat(p.code)).toBe('fraction');
		}
		for (const p of SPECIAL_FORMAT_PRESETS) {
			expect(classifyFormat(p.code)).toBe('special');
		}
		expect(FRACTION_FORMAT_PRESETS).toHaveLength(9);
		expect(SPECIAL_FORMAT_PRESETS.map((p) => p.label)).toContain('Phone Number');
	});

	it('shows the fraction preset samples', () => {
		const samples = [0.25, 0.84, 312 / 943, 0.5, 0.5, 0.5, 0.5, 0.3, 0.3];
		const expected = ['1/4', '21/25', '312/943', '1/2', '2/4', '4/8', '8/16', '3/10', '30/100'];
		FRACTION_FORMAT_PRESETS.forEach((p, i) => {
			expect(formatValue(samples[i], p.code).text.trim()).toBe(expected[i]);
		});
		expect(formatValue(0.33086, '# ???/???').text).toBe(' 223/674');
	});

	it('offers currency symbols and negative styles', () => {
		expect(CURRENCY_SYMBOL_PRESETS.map((p) => p.id)).toEqual(
			expect.arrayContaining(['none', '$', '€', '€-after', '£', '¥']),
		);
		expect(NEGATIVE_NUMBER_STYLES.map((s) => s.style)).toEqual([
			'minus',
			'red',
			'parens',
			'redParens',
			'redMinus',
		]);
	});
});
