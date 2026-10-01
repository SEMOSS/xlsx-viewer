import { describe, expect, it } from 'vitest';

import { ERRORS } from '../values/errors';
import { formatValueForEdit } from './index';

describe('formatValueForEdit', () => {
	it.each<[number, string, string]>([
		[45000, 'm/d/yyyy', '3/15/2023'],
		[45000, 'd-mmm-yy', '3/15/2023'],
		[45000.75, 'm/d/yyyy', '3/15/2023 6:00:00 PM'],
		[45000.5, 'm/d/yyyy h:mm', '3/15/2023 12:00:00 PM'],
		[45000 + 1 / 86_400, 'yyyy-mm-dd', '3/15/2023 12:00:01 AM'],
		[0.75, 'h:mm', '6:00:00 PM'],
		[0, 'h:mm', '12:00:00 AM'],
		[0.999999, 'h:mm:ss', '1/1/1900 12:00:00 AM'],
		[0.4999, 'h:mm:ss', '11:59:51 AM'],
		[1.5, '[h]:mm', '1/1/1900 12:00:00 PM'],
		[1, 'h:mm', '1/1/1900 12:00:00 AM'],
		[0.125, '0.00%', '12.5%'],
		[0.07, '0%', '7%'],
		[-0.5, '0%', '-50%'],
		[1234.5, '#,##0.00', '1234.5'],
		[1234.5, '"$"#,##0.00_);\\("$"#,##0.00\\)', '1234.5'],
		[0.5, '# ?/?', '0.5'],
		[12345, '0.00E+00', '12345'],
		[0.1 + 0.2, 'General', '0.3'],
		[1e20, 'General', '1E+20'],
		[123456789012, 'General', '123456789012'],
		[-1, 'm/d/yyyy', '-1'],
		[2958466, 'm/d/yyyy', '2958466'],
		[5, '@', '5'],
	])('%s with %s -> %s', (value, code, expected) => {
		expect(formatValueForEdit(value, code)).toBe(expected);
	});

	it('uses the 1904 epoch', () => {
		expect(formatValueForEdit(0, 'm/d/yyyy', { date1904: true })).toBe('1/1/1904');
		expect(formatValueForEdit(45000.5, 'm/d/yyyy', { date1904: true })).toBe(
			'3/16/2027 12:00:00 PM',
		);
	});

	it('shows non-numeric values as typed', () => {
		expect(formatValueForEdit(true, '0.00')).toBe('TRUE');
		expect(formatValueForEdit(false)).toBe('FALSE');
		expect(formatValueForEdit('abc', 'm/d/yyyy')).toBe('abc');
		expect(formatValueForEdit(ERRORS.NA)).toBe('#N/A');
		expect(formatValueForEdit(null)).toBe('');
	});
});
