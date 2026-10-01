import { describe, expect, it } from 'vitest';

import type { Table } from '../types/table';
import { tableCellFormat } from './table-format';
import { BUILTIN_TABLE_STYLES, getBuiltinTableStyle, getTableStyle } from './table-styles';

const table: Table = {
	id: 1,
	name: 'Table1',
	displayName: 'Table1',
	ref: { top: 0, left: 0, bottom: 4, right: 2 },
	headerRowCount: 1,
	totalsRowCount: 1,
	columns: [
		{ id: 1, name: 'A' },
		{ id: 2, name: 'B' },
		{ id: 3, name: 'C' },
	],
	style: {
		name: 'TableStyleMedium2',
		showFirstColumn: false,
		showLastColumn: false,
		showRowStripes: true,
		showColumnStripes: false,
	},
};

describe('built-in table styles', () => {
	it('defines all 60 built-in styles', () => {
		expect(BUILTIN_TABLE_STYLES).toHaveLength(60);
		for (const name of BUILTIN_TABLE_STYLES) {
			expect(getBuiltinTableStyle(name), name).toBeDefined();
		}
		expect(getBuiltinTableStyle('TableStyleMedium29')).toBeUndefined();
		expect(getBuiltinTableStyle('Nope')).toBeUndefined();
	});

	it('resolves custom styles through dxfs', () => {
		const def = getTableStyle('Mine', {
			numFmts: {},
			fonts: [],
			fills: [],
			borders: [],
			cellStyleXfs: [],
			cellXfs: [],
			cellStyles: [],
			dxfs: [{ font: { bold: true } }],
			customTableStyles: [{ name: 'Mine', elements: [{ type: 'headerRow', dxfId: 0 }] }],
		});
		expect(def?.elements.headerRow).toEqual({ font: { bold: true } });
	});
});

describe('tableCellFormat (Medium 2)', () => {
	const style = getBuiltinTableStyle('TableStyleMedium2')!;

	it('formats the header row with an accent fill and white bold text', () => {
		const f = tableCellFormat(table, style, 0, 1);
		expect(f?.font).toMatchObject({ bold: true, color: { theme: 0 } });
		expect(f?.fill).toMatchObject({ patternType: 'solid', bgColor: { theme: 4 } });
	});

	it('bands data rows starting with the first stripe', () => {
		expect(tableCellFormat(table, style, 1, 0)?.fill).toMatchObject({
			bgColor: { theme: 4, tint: 0.7999 },
		});
		expect(tableCellFormat(table, style, 2, 0)?.fill).toBeUndefined();
		expect(tableCellFormat(table, style, 3, 0)?.fill).toMatchObject({ bgColor: { theme: 4 } });
	});

	it('puts a double rule above the total row', () => {
		const f = tableCellFormat(table, style, 4, 2);
		expect(f?.font).toMatchObject({ bold: true });
		expect(f?.border?.top).toMatchObject({ style: 'double' });
	});

	it('ignores cells outside the table', () => {
		expect(tableCellFormat(table, style, 5, 0)).toBeUndefined();
	});
});
