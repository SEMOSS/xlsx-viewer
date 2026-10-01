import { describe, expect, it } from 'vitest';

import { baseName, savedFileName } from './download';

describe('file names', () => {
	it('keeps Open XML names and converts others to .xlsx', () => {
		expect(savedFileName('Budget.xlsx')).toBe('Budget.xlsx');
		expect(savedFileName('Macros.xlsm')).toBe('Macros.xlsm');
		expect(savedFileName('Legacy.xls')).toBe('Legacy.xlsx');
		expect(savedFileName('export.csv')).toBe('export.xlsx');
		expect(savedFileName('Untitled')).toBe('Untitled.xlsx');
		expect(baseName('Legacy.XLS')).toBe('Legacy');
	});
});
