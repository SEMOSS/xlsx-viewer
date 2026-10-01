import { describe, expect, it } from 'vitest';

import { cellKey } from '../address/cell-address';
import { parseRange } from '../address/range';
import { createWorkbook } from '../sheet/workbook';
import { StyleResolver } from '../styles/resolve-style';
import type { Workbook } from '../types/workbook';
import { setCellInput } from './cell-input';
import { autoExpandTable } from './table-expand';
import { convertTableToRange, isValidTableName, renameTable } from './table-rename-ops';
import {
	setTableFilterButton,
	setTableHeaderRow,
	setTableStyle,
	setTableStyleOptions,
	setTableTotalsRow,
	tableIndexAt,
} from './table-design-ops';
import { createTable, TableError } from './table-ops';
import { EditTransaction } from './transaction';

function setup(): Workbook {
	const wb = createWorkbook();
	const tx = new EditTransaction(wb, 'Setup');
	const rows = [
		['Name', 'Qty', 'Price'],
		['Pen', '3', '1.5'],
		['Ink', '5', '7'],
		['Pad', '2', '4.25'],
	];
	rows.forEach((r, i) => r.forEach((v, j) => setCellInput(tx, wb, 0, i, j, v)));
	setCellInput(tx, wb, 0, 0, 5, '=SUM(Table1[Qty])');
	setCellInput(tx, wb, 0, 1, 5, '=COUNTA(Table1)');
	createTable(tx, wb, 0, parseRange('A1:C4')!);
	tx.commit();
	return wb;
}

const run = <T>(wb: Workbook, fn: (tx: EditTransaction) => T) => {
	const tx = new EditTransaction(wb, 'Test');
	const out = fn(tx);
	return { out, entry: tx.commit() };
};

const at = (wb: Workbook, a1: string) => {
	const r = parseRange(a1)!;
	return wb.sheets[0].cells.get(cellKey(r.top, r.left));
};

describe('table design operations', () => {
	it('finds the table at a cell', () => {
		const wb = setup();
		expect(tableIndexAt(wb, 0, 2, 1)).toBe(0);
		expect(tableIndexAt(wb, 0, 5, 1)).toBe(-1);
	});

	it('adds and removes a totals row with SUBTOTAL formulas, with undo', () => {
		const wb = setup();
		const { entry } = run(wb, (tx) => setTableTotalsRow(tx, wb, 0, 0, true));
		const table = wb.sheets[0].tables[0];
		expect(table.ref).toEqual(parseRange('A1:C5'));
		expect(table.totalsRowCount).toBe(1);
		expect(at(wb, 'A5')?.value).toBe('Total');
		expect(at(wb, 'C5')?.formula).toBe('SUBTOTAL(109,Table1[Price])');
		expect(table.autoFilter?.ref).toEqual(parseRange('A1:C4'));
		entry.undo();
		expect(wb.sheets[0].tables[0].ref).toEqual(parseRange('A1:C4'));
		expect(at(wb, 'A5')).toBeUndefined();
		entry.redo();
		run(wb, (tx) => setTableTotalsRow(tx, wb, 0, 0, false));
		expect(wb.sheets[0].tables[0].totalsRowCount).toBe(0);
		expect(at(wb, 'C5')).toBeUndefined();
	});

	it('refuses a totals row over data', () => {
		const wb = setup();
		run(wb, (tx) => setCellInput(tx, wb, 0, 4, 1, 'busy'));
		expect(() => run(wb, (tx) => setTableTotalsRow(tx, wb, 0, 0, true))).toThrow(TableError);
	});

	it('hides and restores the header row', () => {
		const wb = setup();
		run(wb, (tx) => setTableHeaderRow(tx, wb, 0, 0, false));
		expect(wb.sheets[0].tables[0].ref).toEqual(parseRange('A2:C4'));
		expect(wb.sheets[0].tables[0].autoFilter).toBeUndefined();
		expect(at(wb, 'A1')).toBeUndefined();
		run(wb, (tx) => setTableHeaderRow(tx, wb, 0, 0, true));
		expect(wb.sheets[0].tables[0].ref).toEqual(parseRange('A1:C4'));
		expect(at(wb, 'B1')?.value).toBe('Qty');
		expect(wb.sheets[0].tables[0].autoFilter?.ref).toEqual(parseRange('A1:C4'));
	});

	it('changes style and options and toggles the filter buttons', () => {
		const wb = setup();
		run(wb, (tx) => {
			setTableStyle(tx, wb, 0, 0, 'TableStyleLight9');
			setTableStyleOptions(tx, wb, 0, 0, { showColumnStripes: true, showRowStripes: false });
			setTableFilterButton(tx, wb, 0, 0, false);
		});
		const table = wb.sheets[0].tables[0];
		expect(table.style).toMatchObject({
			name: 'TableStyleLight9',
			showColumnStripes: true,
			showRowStripes: false,
		});
		expect(table.autoFilter).toBeUndefined();
		run(wb, (tx) => setTableFilterButton(tx, wb, 0, 0, true));
		expect(wb.sheets[0].tables[0].autoFilter?.ref).toEqual(parseRange('A1:C4'));
	});

	it('validates table names', () => {
		expect(isValidTableName('Sales_2026')).toBe(true);
		expect(isValidTableName('_x.y')).toBe(true);
		expect(isValidTableName('A1')).toBe(false);
		expect(isValidTableName('R1C1')).toBe(false);
		expect(isValidTableName('c')).toBe(false);
		expect(isValidTableName('two words')).toBe(false);
		expect(isValidTableName('1st')).toBe(false);
	});

	it('renames a table and its structured references', () => {
		const wb = setup();
		run(wb, (tx) => renameTable(tx, wb, 0, 0, 'Orders'));
		expect(wb.sheets[0].tables[0]).toMatchObject({ name: 'Orders', displayName: 'Orders' });
		expect(at(wb, 'F1')?.formula).toBe('SUM(Orders[Qty])');
		expect(at(wb, 'F2')?.formula).toBe('COUNTA(Orders)');
		expect(() => run(wb, (tx) => renameTable(tx, wb, 0, 0, 'B2'))).toThrow(TableError);
		wb.definedNames.push({ name: 'Taken', formula: 'Sheet1!$A$1' });
		expect(() => run(wb, (tx) => renameTable(tx, wb, 0, 0, 'taken'))).toThrow(/already in use/);
	});

	it('converts to a range: references become A1 and the style becomes cell formatting', () => {
		const wb = setup();
		run(wb, (tx) => convertTableToRange(tx, wb, 0, 0));
		expect(wb.sheets[0].tables).toHaveLength(0);
		expect(at(wb, 'F1')?.formula).toBe('SUM($B$2:$B$4)');
		expect(at(wb, 'F2')?.formula).toBe('COUNTA($A$2:$C$4)');
		const resolver = new StyleResolver(wb.styles, wb.theme);
		const header = resolver.resolve(at(wb, 'A1')?.style);
		expect(header.fill).toBeDefined();
		expect(header.font.bold).toBe(true);
		const stripe = resolver.resolve(at(wb, 'A2')?.style);
		expect(stripe.fill).toBeDefined();
	});
});

describe('table auto-expansion', () => {
	it('grows a table when typing directly below or to the right of it', () => {
		const wb = setup();
		run(wb, (tx) => {
			setCellInput(tx, wb, 0, 4, 0, 'Cap');
			expect(autoExpandTable(tx, wb, 0, 4, 0)).toBe(true);
		});
		expect(wb.sheets[0].tables[0].ref).toEqual(parseRange('A1:C5'));
		expect(wb.sheets[0].tables[0].autoFilter?.ref).toEqual(parseRange('A1:C5'));
		run(wb, (tx) => {
			setCellInput(tx, wb, 0, 2, 3, '9');
			expect(autoExpandTable(tx, wb, 0, 2, 3)).toBe(true);
		});
		const table = wb.sheets[0].tables[0];
		expect(table.ref).toEqual(parseRange('A1:D5'));
		expect(table.columns.map((c) => c.name)).toEqual(['Name', 'Qty', 'Price', 'Column4']);
		expect(at(wb, 'D1')?.value).toBe('Column4');
	});

	it('leaves cells elsewhere alone', () => {
		const wb = setup();
		run(wb, (tx) => expect(autoExpandTable(tx, wb, 0, 6, 0)).toBe(false));
		expect(wb.sheets[0].tables[0].ref).toEqual(parseRange('A1:C4'));
	});
});
