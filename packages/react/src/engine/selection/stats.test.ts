import { describe, expect, it } from 'vitest';
import type { RangeAddress, Workbook, Worksheet } from 'xlsx-viewer-core';
import {
	cellError,
	createWorkbook,
	EditTransaction,
	parseRange,
	putCell,
	setCellInput,
	setRowsHidden,
} from 'xlsx-viewer-core';

import { selectAll } from './selection';
import { selectionStats } from './stats';

const r = (ref: string): RangeAddress => parseRange(ref)!;

/** A sheet with values typed into cells (`{ A1: '10', B2: 'abc' }`). */
function sheetWith(
	typed: Record<string, string>,
	build?: (tx: EditTransaction, wb: Workbook) => void,
): Worksheet {
	const wb = createWorkbook();
	const tx = new EditTransaction(wb, 'setup');
	for (const [ref, text] of Object.entries(typed)) {
		const { top, left } = r(ref);
		setCellInput(tx, wb, 0, top, left, text);
	}
	build?.(tx, wb);
	tx.commit();
	return wb.sheets[0];
}

describe('selectionStats', () => {
	it('sums, averages and finds the extremes of numbers', () => {
		const sheet = sheetWith({ A1: '10', A2: '20', A3: '30.5', A4: '-4' });
		expect(selectionStats(sheet, [r('A1:A4')])).toEqual({
			count: 4,
			numericCount: 4,
			sum: 56.5,
			min: -4,
			max: 30.5,
			average: 56.5 / 4,
		});
	});

	it('counts text, booleans and errors without adding them up', () => {
		const sheet = sheetWith({ A1: '5', A2: 'abc', A3: 'TRUE', A4: '#DIV/0!', A5: '7' });
		// Numbers stored as text are counted, not summed (like Excel's status bar).
		putCell(sheet, 5, 0, { value: '100' });
		expect(selectionStats(sheet, [r('A1:A6')])).toEqual({
			count: 6,
			numericCount: 2,
			sum: 12,
			min: 5,
			max: 7,
			average: 6,
		});
	});

	it('ignores blank and empty-string cells, even when formatted', () => {
		const sheet = sheetWith({ A1: '1' });
		putCell(sheet, 1, 0, { value: null, style: 1 });
		putCell(sheet, 2, 0, { value: '' });
		expect(selectionStats(sheet, [r('A1:A10')])).toMatchObject({
			count: 1,
			numericCount: 1,
			sum: 1,
		});
	});

	it('reports no minimum, maximum or average without numbers', () => {
		const text = sheetWith({ A1: 'x', B1: 'FALSE' });
		expect(selectionStats(text, [r('A1:B1')])).toEqual({
			count: 2,
			numericCount: 0,
			sum: 0,
			min: null,
			max: null,
			average: null,
		});
		expect(selectionStats(text, [])).toMatchObject({ count: 0, sum: 0, average: null });
		expect(selectionStats(text, [r('D4:E9')]).count).toBe(0);
	});

	it('uses the cached values of formulas', () => {
		const sheet = sheetWith({});
		putCell(sheet, 0, 0, { value: 4, formula: '2*2' });
		putCell(sheet, 1, 0, { value: 'x', formula: '"x"' });
		putCell(sheet, 2, 0, { value: null, formula: 'NOW()' });
		putCell(sheet, 3, 0, { value: cellError('#REF!'), formula: '#REF!+1' });
		expect(selectionStats(sheet, [r('A1:A4')])).toMatchObject({
			count: 3,
			numericCount: 1,
			sum: 4,
		});
	});

	it('counts cells covered by overlapping ranges once', () => {
		const sheet = sheetWith({ A1: '1', B2: '10', C3: '100' });
		expect(selectionStats(sheet, [r('A1:B2'), r('B2:C3')])).toMatchObject({
			count: 3,
			sum: 111,
		});
		expect(selectionStats(sheet, [r('B2'), r('B2'), r('A1:C3')]).sum).toBe(111);
	});

	it('combines disjoint ranges', () => {
		const sheet = sheetWith({ A1: '2', A2: '4', C1: '6', C2: 'skip me' });
		expect(selectionStats(sheet, [r('A1:A2'), r('C1')])).toMatchObject({
			count: 3,
			sum: 12,
			average: 4,
		});
	});

	it('skips hidden rows, like Excel', () => {
		const sheet = sheetWith({ A1: '1', A2: '10', A3: '100' }, (tx, wb) =>
			setRowsHidden(tx, wb, 0, 1, 1, true),
		);
		expect(selectionStats(sheet, [r('A1:A3')])).toMatchObject({ count: 2, sum: 101, max: 100 });
		expect(selectionStats(sheet, [r('A:A')])).toMatchObject({ count: 2, sum: 101 });
	});

	it('handles whole-column and select-all ranges through the stored cells', () => {
		const sheet = sheetWith({ A1: '1', A500000: '2', B1: '40', XFD1048576: '8' });
		expect(selectionStats(sheet, [r('A:A')])).toMatchObject({ count: 2, sum: 3, max: 2 });
		expect(selectionStats(sheet, [r('A:B')])).toMatchObject({ count: 3, sum: 43 });
		expect(selectionStats(sheet, [r('1:1')])).toMatchObject({ count: 2, sum: 41 });
		expect(selectionStats(sheet, [selectAll()])).toMatchObject({ count: 4, sum: 51, max: 40 });
		// The same answer as cell-by-cell iteration over a small range.
		expect(selectionStats(sheet, [r('A1:B1')])).toEqual(selectionStats(sheet, [r('1:1'), r('A1')]));
	});
});
