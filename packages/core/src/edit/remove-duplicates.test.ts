import { describe, expect, it } from 'vitest';

import { cellKey } from '../address/cell-address';
import { MAX_ROW_INDEX } from '../address/constants';
import { parseRange } from '../address/range';
import { createWorkbook } from '../sheet/workbook';
import { getCell } from '../sheet/worksheet';
import { StylePool } from '../styles/style-pool';
import type { Workbook } from '../types/workbook';
import { setCellInput } from './cell-input';
import {
	findDuplicateRows,
	RemoveDuplicatesError,
	removeDuplicateRows,
	trimRangeToUsed,
} from './remove-duplicates';
import { EditTransaction } from './transaction';

const r = (a1: string) => parseRange(a1)!;

/** Fill rows starting at A1 (each entry typed like user input). */
function sheetWith(rows: string[][]): Workbook {
	const wb = createWorkbook();
	const tx = new EditTransaction(wb, 'seed');
	rows.forEach((row, i) => row.forEach((v, j) => setCellInput(tx, wb, 0, i, j, v)));
	tx.commit();
	return wb;
}

const values = (wb: Workbook, ref: string) => {
	const a = r(ref);
	const out: unknown[][] = [];
	for (let row = a.top; row <= a.bottom; row++) {
		const line: unknown[] = [];
		for (let col = a.left; col <= a.right; col++) {
			line.push(getCell(wb.sheets[0], row, col)?.value ?? null);
		}
		out.push(line);
	}
	return out;
};

function remove(wb: Workbook, ref: string, columns: number[], hasHeader: boolean) {
	const tx = new EditTransaction(wb, 'Remove Duplicates');
	const result = removeDuplicateRows(tx, wb, 0, r(ref), { columns, hasHeader });
	return { result, entry: tx.commit() };
}

const DATA = [
	['Name', 'City', 'Qty'],
	['Ann', 'Oslo', '1'],
	['Bob', 'Rome', '2'],
	['ann', 'OSLO', '3'],
	['Cid', 'Rome', '2'],
	['Bob', 'Rome', '5'],
];

describe('findDuplicateRows', () => {
	it('keeps first occurrences, comparing text case-insensitively', () => {
		const wb = sheetWith(DATA);
		const found = findDuplicateRows(wb, 0, r('A1:C6'), { columns: [0, 1], hasHeader: true });
		expect(found.kept).toEqual([1, 2, 4]);
		expect(found.duplicates).toEqual([3, 5]);
		const all = findDuplicateRows(wb, 0, r('A1:C6'), { columns: [0, 1, 2], hasHeader: true });
		expect(all.duplicates).toEqual([]);
	});

	it('compares displayed values rather than stored values', () => {
		const wb = sheetWith([['1'], ['1.0'], ['1/1/2024'], ['45292'], ['TRUE'], ['true']]);
		const pct = new StylePool(wb.styles).derive(undefined, { numFmt: '0%' });
		wb.sheets[0].cells.set(cellKey(1, 0), { value: 1, style: pct });
		const found = findDuplicateRows(wb, 0, r('A1:A6'), { columns: [0], hasHeader: false });
		// 1 vs 100% differ; the date and its serial number display differently.
		expect(found.duplicates).toEqual([5]);
	});

	it('treats blanks as values and ignores columns outside the range', () => {
		const wb = sheetWith([
			['a', 'x'],
			['', 'y'],
			['a', 'z'],
			['', 'w'],
		]);
		const found = findDuplicateRows(wb, 0, r('A1:A4'), { columns: [0, 1], hasHeader: false });
		expect(found.duplicates).toEqual([2, 3]);
		expect(
			findDuplicateRows(wb, 0, r('A1:A4'), { columns: [], hasHeader: false }).duplicates,
		).toEqual([]);
	});

	it('limits whole columns to the used rows', () => {
		const wb = sheetWith([['a'], ['a'], ['b']]);
		const whole = { top: 0, left: 0, bottom: MAX_ROW_INDEX, right: 0 };
		expect(trimRangeToUsed(wb.sheets[0], whole)).toEqual(r('A1:A3'));
		expect(trimRangeToUsed(wb.sheets[0], r('A1:B9'))).toEqual(r('A1:B9'));
		expect(findDuplicateRows(wb, 0, whole, { columns: [0], hasHeader: false })).toEqual({
			kept: [0, 2],
			duplicates: [1],
			range: r('A1:A3'),
		});
	});

	it('refuses merged cells and multi-cell arrays', () => {
		const wb = sheetWith(DATA);
		wb.sheets[0].merges.push(r('C5:D5'));
		expect(() => findDuplicateRows(wb, 0, r('A1:C6'), { columns: [0], hasHeader: true })).toThrow(
			"We can't do that to a merged cell.",
		);
		wb.sheets[0].merges = [];
		wb.sheets[0].cells.set(cellKey(2, 2), { value: 2, arrayAnchor: cellKey(1, 2) });
		let error: unknown;
		try {
			findDuplicateRows(wb, 0, r('A1:C6'), { columns: [0], hasHeader: true });
		} catch (e) {
			error = e;
		}
		expect(error).toBeInstanceOf(RemoveDuplicatesError);
		expect((error as RemoveDuplicatesError).reason).toBe('array');
	});
});

describe('removeDuplicateRows', () => {
	it('compacts kept rows upwards and empties the freed rows', () => {
		const wb = sheetWith([...DATA, ['Bob', 'Rome', '9']]);
		setCellInput(new EditTransaction(wb, 'x'), wb, 0, 1, 4, 'outside');
		const { result } = remove(wb, 'A1:C7', [0, 1], true);
		expect(result).toEqual({ removed: 3, remaining: 3 });
		expect(values(wb, 'A1:C7')).toEqual([
			['Name', 'City', 'Qty'],
			['Ann', 'Oslo', 1],
			['Bob', 'Rome', 2],
			['Cid', 'Rome', 2],
			[null, null, null],
			[null, null, null],
			[null, null, null],
		]);
		expect(getCell(wb.sheets[0], 1, 4)?.value).toBe('outside');
	});

	it('moves formats and translates formulas as if copied', () => {
		const wb = sheetWith([
			['a', '=ROW()'],
			['a', '1'],
			['b', '=A3&"!"'],
			['c', '=$A$1'],
		]);
		const bold = new StylePool(wb.styles).derive(undefined, { font: { bold: true } });
		const b3 = getCell(wb.sheets[0], 2, 0)!;
		wb.sheets[0].cells.set(cellKey(2, 0), { ...b3, style: bold });
		remove(wb, 'A1:B4', [0], false);
		expect(getCell(wb.sheets[0], 1, 0)).toEqual({ value: 'b', style: bold });
		expect(getCell(wb.sheets[0], 1, 1)).toMatchObject({
			formula: 'A2&"!"',
			array: { ref: r('B2') },
		});
		expect(getCell(wb.sheets[0], 2, 1)?.formula).toBe('$A$1');
		expect(getCell(wb.sheets[0], 3, 0)).toBeUndefined();
	});

	it('changes nothing without duplicates and undoes / redoes', () => {
		const wb = sheetWith(DATA);
		const none = remove(wb, 'A1:C6', [0, 1, 2], true);
		expect(none.result).toEqual({ removed: 0, remaining: 5 });
		expect(none.entry.sheets).toEqual([]);

		const before = values(wb, 'A1:C6');
		const { entry } = remove(wb, 'A1:C6', [1], true);
		const after = values(wb, 'A1:C6');
		expect(after.map((row) => row[1])).toEqual(['City', 'Oslo', 'Rome', null, null, null]);
		entry.undo();
		expect(values(wb, 'A1:C6')).toEqual(before);
		entry.redo();
		expect(values(wb, 'A1:C6')).toEqual(after);
	});
});
