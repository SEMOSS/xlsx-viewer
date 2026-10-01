import { describe, expect, it } from 'vitest';

import type { Table } from '../../types/table';
import { ERRORS } from '../../values/errors';
import { parseStructuredReference } from '../language/structured';
import { structuredRange, tableColumnIndex, tableDataRange } from './structured';

/** Table1 at B2:D6: header row 2, data rows 3-5, totals row 6; columns Qty, Unit Price, Total. */
const table: Table = {
	id: 1,
	name: 'Table1',
	displayName: 'Table1',
	ref: { top: 1, left: 1, bottom: 5, right: 3 },
	headerRowCount: 1,
	totalsRowCount: 1,
	columns: [
		{ id: 1, name: 'Qty' },
		{ id: 2, name: 'Unit Price' },
		{ id: 3, name: 'Total' },
	],
};

const range = (text: string, row = 3) =>
	structuredRange(parseStructuredReference(text)!, table, row);

describe('structured references', () => {
	it('finds columns case-insensitively', () => {
		expect(tableColumnIndex(table, 'unit price')).toBe(1);
		expect(tableColumnIndex(table, 'Nope')).toBe(-1);
	});

	it('addresses the data body by default', () => {
		expect(range('Table1[Qty]')).toEqual({ top: 2, left: 1, bottom: 4, right: 1 });
		expect(range('Table1[]')).toEqual({ top: 2, left: 1, bottom: 4, right: 3 });
		expect(range('Table1[[Qty]:[Total]]')).toEqual({ top: 2, left: 1, bottom: 4, right: 3 });
		expect(range('Table1[[Total]:[Qty]]')).toEqual({ top: 2, left: 1, bottom: 4, right: 3 });
		expect(tableDataRange(table)).toEqual({ top: 2, left: 1, bottom: 4, right: 3 });
	});

	it('handles the special items', () => {
		expect(range('Table1[#All]')).toEqual({ top: 1, left: 1, bottom: 5, right: 3 });
		expect(range('Table1[#Headers]')).toEqual({ top: 1, left: 1, bottom: 1, right: 3 });
		expect(range('Table1[#Totals]')).toEqual({ top: 5, left: 1, bottom: 5, right: 3 });
		expect(range('Table1[[#Headers],[Total]]')).toEqual({ top: 1, left: 3, bottom: 1, right: 3 });
		expect(range('Table1[[#Headers],[#Data],[Qty]]')).toEqual({
			top: 1,
			left: 1,
			bottom: 4,
			right: 1,
		});
		expect(range('Table1[[#Data],[#Totals],[Qty]]')).toEqual({
			top: 2,
			left: 1,
			bottom: 5,
			right: 1,
		});
		expect(range('Table1[[#All],[Unit Price]]')).toEqual({ top: 1, left: 2, bottom: 5, right: 2 });
	});

	it('resolves this-row references against the formula row', () => {
		expect(range('[@Qty]', 3)).toEqual({ top: 3, left: 1, bottom: 3, right: 1 });
		expect(range('[@[Unit Price]]', 4)).toEqual({ top: 4, left: 2, bottom: 4, right: 2 });
		expect(range('Table1[@]', 2)).toEqual({ top: 2, left: 1, bottom: 2, right: 3 });
		expect(range('[@Qty]', 1)).toBe(ERRORS.VALUE);
		expect(range('[@Qty]', 5)).toBe(ERRORS.VALUE);
	});

	it('reports unknown columns and missing parts as #REF!', () => {
		expect(range('Table1[Nope]')).toBe(ERRORS.REF);
		expect(range('Table1[[#Headers],[#Totals]]')).toBe(ERRORS.REF);
		const bare: Table = { ...table, headerRowCount: 0, totalsRowCount: 0 };
		expect(structuredRange(parseStructuredReference('Table1[#Headers]')!, bare, 2)).toBe(
			ERRORS.REF,
		);
		expect(structuredRange(parseStructuredReference('Table1[#Totals]')!, bare, 2)).toBe(ERRORS.REF);
		expect(structuredRange(parseStructuredReference('Table1[Qty]')!, bare, 2)).toEqual({
			top: 1,
			left: 1,
			bottom: 5,
			right: 1,
		});
	});
});
