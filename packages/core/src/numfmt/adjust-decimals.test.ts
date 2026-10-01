import { describe, expect, it } from 'vitest';

import { BUILTIN_NUMBER_FORMATS as B } from '../styles/builtin-number-formats';
import { adjustDecimals } from './index';

describe('adjustDecimals', () => {
	it.each<[string, number, string]>([
		['0', 1, '0.0'],
		['0.0', -1, '0'],
		['0.00', 1, '0.000'],
		['0.00', 2, '0.0000'],
		['0.00', -2, '0'],
		['0.00', -5, '0'],
		['#,##0', 1, '#,##0.0'],
		['#,##0.00_);(#,##0.00)', 1, '#,##0.000_);(#,##0.000)'],
		['#,##0_);[Red](#,##0)', 1, '#,##0.0_);[Red](#,##0.0)'],
		['0%', 1, '0.0%'],
		['0.0%', -1, '0%'],
		['0.00E+00', 1, '0.000E+00'],
		['0.0E+00', -1, '0E+00'],
		['0E+00', 1, '0.0E+00'],
		['##0.0E+0', -1, '##0E+0'],
		['"$"#,##0.00', -1, '"$"#,##0.0'],
		['"$"#,##0_);\\("$"#,##0\\)', 1, '"$"#,##0.0_);\\("$"#,##0.0\\)'],
		['0.0,,"M"', 1, '0.00,,"M"'],
		['#,##0,', 1, '#,##0.0,'],
		['0.00" pts"', 1, '0.000" pts"'],
		['"0.0"0', 1, '"0.0"0.0'],
		['[>=1000]#,##0.0,"K";0', 1, '[>=1000]#,##0.00,"K";0.0'],
		['0.00;-0.00;0;@', 1, '0.000;-0.000;0.0;@'],
		['[$€-407] #,##0.00', 1, '[$€-407] #,##0.000'],
		['h:mm:ss', 1, 'h:mm:ss.0'],
		['[h]:mm:ss', 2, '[h]:mm:ss.00'],
		['h:mm:ss.000', 1, 'h:mm:ss.000'],
		['mm:ss.0', -1, 'mm:ss'],
		['mm:ss.00', -1, 'mm:ss.0'],
		['m/d/yyyy', 1, 'm/d/yyyy'],
		['h:mm', 1, 'h:mm'],
		['# ?/?', 1, '# ?/?'],
		['@', 1, '@'],
		['0', -1, '0'],
		['0.00', 0, '0.00'],
		['[Red]General', 1, '[Red]General'],
	])('%s %+i -> %s', (code, delta, expected) => {
		expect(adjustDecimals(code, delta)).toBe(expected);
	});

	it('walks between the accounting built-ins', () => {
		expect(adjustDecimals(B[41], 2)).toBe(B[43]);
		expect(adjustDecimals(B[43], -2)).toBe(B[41]);
		expect(adjustDecimals(B[42], 2)).toBe(B[44]);
		expect(adjustDecimals(B[44], -2)).toBe(B[42]);
		expect(adjustDecimals(B[44], 1)).toBe(
			'_("$"* #,##0.000_);_("$"* \\(#,##0.000\\);_("$"* "-"???_);_(@_)',
		);
		expect(adjustDecimals(B[42], 1)).toBe(
			'_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)',
		);
		expect(adjustDecimals('_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@_-', 1)).toBe(
			'_-* #,##0.000_-;-* #,##0.000_-;_-* "-"???_-;_-@_-',
		);
	});

	it('walks between the currency and number built-ins', () => {
		expect(adjustDecimals(B[5], 2)).toBe(B[7]);
		expect(adjustDecimals(B[8], -2)).toBe(B[6]);
		expect(adjustDecimals(B[37], 2)).toBe(B[39]);
		expect(adjustDecimals(B[1], 2)).toBe(B[2]);
		expect(adjustDecimals(B[9], 2)).toBe(B[10]);
	});

	it('starts General from the decimals it shows for the value', () => {
		expect(adjustDecimals('General', 1)).toBe('0.0');
		expect(adjustDecimals('General', -1)).toBe('0');
		expect(adjustDecimals('', 1)).toBe('0.0');
		expect(adjustDecimals('General', 1, 1.5)).toBe('0.00');
		expect(adjustDecimals('General', 1, 5)).toBe('0.0');
		expect(adjustDecimals('General', -1, 1.234)).toBe('0.00');
		expect(adjustDecimals('General', -1, 7)).toBe('0');
		expect(adjustDecimals('General', 1, 123456789012)).toBe('0.000000E+00');
		expect(adjustDecimals('General', 0, 1.5)).toBe('General');
	});

	it('caps at 30 decimals', () => {
		const max = '0.' + '0'.repeat(30);
		expect(adjustDecimals(max, 1)).toBe(max);
		expect(adjustDecimals('0.' + '0'.repeat(29), 5)).toBe(max);
	});
});
