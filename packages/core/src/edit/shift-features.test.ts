import { describe, expect, it } from 'vitest';

import { parseRange } from '../address/range';
import { createWorkbook } from '../sheet/workbook';
import type { ChartModel } from '../types/chart';
import type { DrawingObject } from '../types/drawing';
import type { Table } from '../types/table';
import type { Workbook } from '../types/workbook';
import { checkCellShift, deleteCellsShift, insertCellsShift } from './shift-cells';
import { EditTransaction } from './transaction';

const r = (a1: string) => parseRange(a1)!;

function book(): Workbook {
	return createWorkbook({ sheetNames: ['Sheet1', 'Sheet2'] });
}

function run(wb: Workbook, fn: (tx: EditTransaction) => void) {
	const tx = new EditTransaction(wb, 'shift');
	fn(tx);
	return tx.commit();
}

function table(ref: string, extra: Partial<Table> = {}): Table {
	const range = r(ref);
	return {
		id: 1,
		name: 'Table1',
		displayName: 'Table1',
		ref: range,
		headerRowCount: 1,
		totalsRowCount: 0,
		columns: [],
		autoFilter: { ref: range, columns: [] },
		...extra,
	};
}

function chart(values: string): DrawingObject {
	const model: ChartModel = {
		title: { ref: 'Sheet1!$B$1' },
		plots: [
			{
				type: 'bar',
				axisIds: [],
				series: [{ idx: 0, order: 0, values: { kind: 'num', ref: values } }],
			},
		],
		axes: [],
	};
	const at = { col: 0, colOff: 0, row: 0, rowOff: 0 };
	return { id: '1', kind: 'chart', anchor: { type: 'twoCell', from: at, to: at }, chart: model };
}

describe('cell shifts: rules, validations, links and notes', () => {
	it('adjusts conditional format and validation ranges and formulas with the band rule', () => {
		const wb = book();
		const s = wb.sheets[0];
		s.conditionalFormats.push(
			{ ranges: [r('B5:B10')], rules: [{ type: 'expression', priority: 1, formulas: ['B5>0'] }] },
			{ ranges: [r('A1:C10')], rules: [{ type: 'expression', priority: 2, formulas: ['$B1>0'] }] },
		);
		s.dataValidations.push(
			{ ranges: [r('B1:B10')], type: 'list', formula1: '$B$20:$B$22' },
			{ ranges: [r('B3')], type: 'whole', formula1: '1', formula2: 'B4' },
		);
		wb.sheets[1].conditionalFormats.push({
			ranges: [r('A1')],
			rules: [{ type: 'expression', priority: 1, formulas: ['Sheet1!B5=1'] }],
		});
		run(wb, (tx) => insertCellsShift(tx, wb, 0, r('B3'), 'down'));
		expect(s.conditionalFormats.map((cf) => cf.ranges)).toEqual([[r('B6:B11')], [r('A1:C10')]]);
		expect(s.conditionalFormats.map((cf) => cf.rules[0].formulas)).toEqual([['B6>0'], ['$B1>0']]);
		expect(s.dataValidations.map((dv) => dv.ranges)).toEqual([[r('B1:B11')], [r('B4')]]);
		expect(s.dataValidations[0].formula1).toBe('$B$21:$B$23');
		expect(s.dataValidations[1].formula2).toBe('B5');
		expect(wb.sheets[1].conditionalFormats[0].rules[0].formulas).toEqual(['Sheet1!B6=1']);

		run(wb, (tx) => deleteCellsShift(tx, wb, 0, r('B4'), 'up'));
		expect(s.dataValidations.map((dv) => dv.ranges)).toEqual([[r('B1:B10')]]);
	});

	it('moves hyperlinks, notes and threaded comments with their cells', () => {
		const wb = book();
		const s = wb.sheets[0];
		s.hyperlinks.push(
			{ ref: r('B4'), target: 'https://a.example' },
			{ ref: r('C4'), location: 'A1' },
		);
		s.notes.push(
			{ ref: { row: 2, col: 1 }, author: 'A', text: 'moves' },
			{ ref: { row: 2, col: 2 }, author: 'A', text: 'stays' },
		);
		s.threadedComments.push({ id: 't1', ref: { row: 5, col: 1 }, personId: 'p', text: 'hi' });
		s.commentsDirty = false;
		run(wb, (tx) => insertCellsShift(tx, wb, 0, r('B3'), 'down'));
		expect(s.hyperlinks.map((h) => h.ref)).toEqual([r('B5'), r('C4')]);
		expect(s.notes.map((n) => n.ref)).toEqual([
			{ row: 3, col: 1 },
			{ row: 2, col: 2 },
		]);
		expect(s.threadedComments[0].ref).toEqual({ row: 6, col: 1 });
		expect(s.commentsDirty).toBe(true);

		run(wb, (tx) => deleteCellsShift(tx, wb, 0, r('B4:B5'), 'up'));
		expect(s.hyperlinks.map((h) => h.ref)).toEqual([r('C4')]);
		expect(s.notes.map((n) => n.text)).toEqual(['stays']);
		expect(s.threadedComments[0].ref).toEqual({ row: 4, col: 1 });
	});

	it('moves the sheet filter, sparklines and chart references', () => {
		const wb = book();
		const s = wb.sheets[0];
		s.autoFilter = { ref: r('B1:B10'), columns: [] };
		s.sparklineGroups.push({
			type: 'line',
			sparklines: [
				{ formula: 'Sheet1!B1:B4', location: { row: 4, col: 1 } },
				{ formula: 'Sheet1!A1:A4', location: { row: 4, col: 0 } },
			],
		});
		wb.sheets[1].drawings.push(chart('Sheet1!$B$1:$B$5'));
		wb.sheets[1].drawingsDirty = false;
		run(wb, (tx) => insertCellsShift(tx, wb, 0, r('B3'), 'down'));
		expect(s.autoFilter?.ref).toEqual(r('B1:B11'));
		expect(s.sparklineGroups[0].sparklines).toEqual([
			{ formula: 'Sheet1!B1:B5', location: { row: 5, col: 1 } },
			{ formula: 'Sheet1!A1:A4', location: { row: 4, col: 0 } },
		]);
		const moved = wb.sheets[1].drawings[0];
		expect(moved.kind === 'chart' && moved.chart.plots[0].series[0].values?.ref).toBe(
			'Sheet1!$B$1:$B$6',
		);
		expect(moved.kind === 'chart' && moved.chart.title?.ref).toBe('Sheet1!$B$1');
		expect(moved.kind === 'chart' && moved.chart.dirty).toBe(true);
		expect(wb.sheets[1].drawingsDirty).toBe(true);
	});
});

describe('cell shifts: tables', () => {
	it('moves tables inside the block and refuses to split them', () => {
		const wb = book();
		const s = wb.sheets[0];
		s.tables.push(table('B10:B12'));
		run(wb, (tx) => insertCellsShift(tx, wb, 0, r('B3:B4'), 'down'));
		expect(s.tables[0]).toMatchObject({
			ref: r('B12:B14'),
			autoFilter: { ref: r('B12:B14') },
			dirty: true,
		});
		s.tables[0] = table('A10:C12');
		const e = checkCellShift(wb, 0, r('B3'), 'down');
		expect(e?.reason).toBe('table');
		expect(e?.message).toBe(
			'This operation is not allowed. The operation is attempting to shift cells in a table ' +
				'on your worksheet.',
		);
		expect(checkCellShift(wb, 0, r('D11'), 'right')).toBeNull();
		expect(checkCellShift(wb, 0, r('B11'), 'right')?.reason).toBe('table');
	});

	it('inserts and deletes table rows when the band spans the whole table width', () => {
		const wb = book();
		const s = wb.sheets[0];
		s.tables.push(table('A10:C15'));
		run(wb, (tx) => insertCellsShift(tx, wb, 0, r('A12:C13'), 'down'));
		expect(s.tables[0].ref).toEqual(r('A10:C17'));
		run(wb, (tx) => deleteCellsShift(tx, wb, 0, r('A11:D12'), 'up'));
		expect(s.tables[0].ref).toEqual(r('A10:C15'));
		expect(s.tables[0].autoFilter?.ref).toEqual(r('A10:C15'));
		// The header row and the last data row cannot go.
		expect(checkCellShift(wb, 0, r('A10:C10'), 'up')?.reason).toBe('table');
		expect(checkCellShift(wb, 0, r('A11:C15'), 'up')?.reason).toBe('table');
		expect(checkCellShift(wb, 0, r('A11:C14'), 'up')).toBeNull();
	});

	it('keeps totals rows out of the deleted rows and removes whole tables', () => {
		const wb = book();
		const s = wb.sheets[0];
		s.tables.push(
			table('A1:B5', { totalsRowCount: 1, autoFilter: { ref: r('A1:B4'), columns: [] } }),
		);
		expect(checkCellShift(wb, 0, r('A5:B5'), 'up')?.reason).toBe('table');
		run(wb, (tx) => insertCellsShift(tx, wb, 0, r('A5:B5'), 'down'));
		expect(s.tables[0]).toMatchObject({ ref: r('A1:B6'), autoFilter: { ref: r('A1:B5') } });
		run(wb, (tx) => deleteCellsShift(tx, wb, 0, r('A1:C6'), 'up'));
		expect(s.tables).toEqual([]);
	});
});
