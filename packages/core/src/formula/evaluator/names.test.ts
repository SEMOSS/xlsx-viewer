import { describe, expect, it } from 'vitest';

import { calculated, edit, evalIn, valueAt } from '../../calc/test-utils/workbook';
import type { Table } from '../../types/table';
import type { Workbook } from '../../types/workbook';
import { ERRORS } from '../../values/errors';

function withNames(wb: Workbook): void {
	wb.definedNames.push(
		{ name: 'Rate', formula: '0.25' },
		{ name: 'Greeting', formula: '"hello"' },
		{ name: 'Data', formula: 'Sheet1!$A$1:$A$3' },
		{ name: 'Total', formula: 'SUM(Data)' },
		{ name: 'Double', formula: 'LAMBDA(x,x*2)' },
		{ name: 'Fact', formula: 'LAMBDA(n,IF(n<=1,1,n*Fact(n-1)))' },
		{ name: 'Local', formula: '"sheet2 only"', localSheet: 1 },
		{ name: 'Rate', formula: '0.5', localSheet: 1 },
		{ name: 'Left', formula: 'Sheet1!A1' },
		{ name: 'Loop', formula: 'Loop+1' },
		{ name: 'Consts', formula: '{1,2,3}' },
		{ name: '_xlnm.Print_Area', formula: 'Sheet1!$A$1:$B$3', localSheet: 0 },
		{ name: 'Broken', formula: 'SUM(' },
	);
}

describe('defined names', () => {
	const { engine, wb } = calculated(
		{
			Sheet1: { A1: 1, A2: 2, A3: 3, B1: 10 },
			Sheet2: { A1: 'b' },
		},
		{},
		withNames,
	);
	const ev = (f: string, at = 'Sheet1!E2') => evalIn(engine, f, at);

	it('evaluates constants, ranges and formulas', () => {
		expect(ev('Rate*4')).toBe(1);
		expect(ev('rate*4')).toBe(1);
		expect(ev('Greeting')).toBe('hello');
		expect(ev('SUM(Data)')).toBe(6);
		expect(ev('Data*10')).toEqual([[10], [20], [30]]);
		expect(ev('Total')).toBe(6);
		expect(ev('ROWS(Data)')).toBe(3);
		expect(ev('SUM(Consts)')).toBe(6);
		expect(ev('ISREF(Data)')).toBe(true);
	});

	it('prefers sheet-scoped names on their sheet', () => {
		expect(ev('Rate')).toBe(0.25);
		expect(ev('Rate', 'Sheet2!E2')).toBe(0.5);
		expect(ev('Local', 'Sheet2!E2')).toBe('sheet2 only');
		expect(ev('Local')).toEqual(ERRORS.NAME);
		expect(ev('Sheet2!Local')).toBe('sheet2 only');
		expect(ev('Sheet2!Rate')).toBe(0.5);
		expect(ev('Sheet1!Rate')).toBe(0.25);
	});

	it('finds built-in names with or without the _xlnm. prefix', () => {
		expect(ev('ROWS(Print_Area)')).toBe(3);
		expect(ev('COLUMNS(_xlnm.Print_Area)')).toBe(2);
	});

	it('shifts relative references in names by the formula cell', () => {
		// `Left` is Sheet1!A1 relative to A1: from B2 it means B2, from A2 it means A2.
		expect(ev('Left', 'Sheet1!A2')).toBe(2);
		expect(ev('Left', 'Sheet1!B1')).toBe(10);
	});

	it('calls named LAMBDAs, recursively', () => {
		expect(ev('Double(21)')).toBe(42);
		expect(ev('Fact(5)')).toBe(120);
		expect(ev('MAP(Data,Double)')).toEqual([[2], [4], [6]]);
	});

	it('guards recursive and broken names', () => {
		expect(ev('Loop')).toEqual(ERRORS.NAME);
		expect(ev('Broken')).toEqual(ERRORS.NAME);
		expect(ev('NoSuchName')).toEqual(ERRORS.NAME);
	});

	it('recalculates formulas through names when cells change', () => {
		edit(engine, 'Sheet1!C1', '=Total*Rate');
		expect(valueAt(wb, 'C1')).toBe(1.5);
		edit(engine, 'Sheet1!A3', 13);
		expect(valueAt(wb, 'C1')).toBe(4);
	});
});

const salesTable = (): Table => ({
	id: 1,
	name: 'Sales',
	displayName: 'Sales',
	ref: { top: 0, left: 0, bottom: 3, right: 2 },
	headerRowCount: 1,
	totalsRowCount: 1,
	columns: [
		{ id: 1, name: 'Item' },
		{ id: 2, name: 'Qty' },
		{ id: 3, name: 'Amount' },
	],
});

describe('structured references', () => {
	const { engine, wb } = calculated(
		{
			Sheet1: {
				A1: 'Item',
				B1: 'Qty',
				C1: 'Amount',
				A2: 'pen',
				B2: 2,
				C2: '=[@Qty]*1.5',
				A3: 'ink',
				B3: 4,
				C3: '=[@Qty]*1.5',
				A4: 'Total',
				B4: '=SUBTOTAL(109,[Qty])',
				C4: '=SUM(Sales[Amount])',
			},
			Sheet2: { A1: '=SUM(Sales[Qty])', A2: '=ROWS(Sales)', A3: '=Sales[[#Headers],[Amount]]' },
		},
		{},
		(w) => w.sheets[0].tables.push(salesTable()),
	);
	const ev = (f: string, at = 'Sheet2!E2') => evalIn(engine, f, at);

	it('computes calculated columns and totals', () => {
		expect(valueAt(wb, 'C2')).toBe(3);
		expect(valueAt(wb, 'C3')).toBe(6);
		expect(valueAt(wb, 'B4')).toBe(6);
		expect(valueAt(wb, 'C4')).toBe(9);
		expect(valueAt(wb, 'Sheet2!A1')).toBe(6);
		expect(valueAt(wb, 'Sheet2!A2')).toBe(2);
		expect(valueAt(wb, 'Sheet2!A3')).toBe('Amount');
	});

	it('evaluates the table specifiers', () => {
		expect(ev('ROWS(Sales[#All])')).toBe(4);
		expect(ev('Sales[[#Totals],[Qty]]')).toBe(6);
		expect(ev('COUNTA(Sales[[#Headers],[Item]:[Qty]])')).toBe(2);
		expect(ev('sales[QTY]')).toEqual([[2], [4]]);
		expect(ev('SUM(Sales)')).toBe(15);
		expect(ev('Sales[Nope]')).toEqual(ERRORS.REF);
		expect(ev('NoTable[Qty]')).toEqual(ERRORS.REF);
		expect(ev('Sales[@Qty]', 'Sheet1!E3')).toBe(4);
		expect(ev('Sales[@Qty]', 'Sheet1!E9')).toEqual(ERRORS.VALUE);
	});

	it('recalculates only the affected row of a calculated column', () => {
		const r = edit(engine, 'Sheet1!B2', 10);
		expect(valueAt(wb, 'C2')).toBe(15);
		expect(valueAt(wb, 'C3')).toBe(6);
		const changed = r.changed.map((c) => `${c.sheet}:${c.row}:${c.col}`);
		expect(changed).toContain('0:1:2');
		expect(changed).not.toContain('0:2:2');
		expect(valueAt(wb, 'C4')).toBe(21);
		expect(valueAt(wb, 'Sheet2!A1')).toBe(14);
	});
});
