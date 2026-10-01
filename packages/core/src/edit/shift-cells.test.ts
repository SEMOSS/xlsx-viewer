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
	CELL_SHIFT_MESSAGES,
	CellShiftError,
	checkCellShift,
	deleteCellsShift,
	insertCellsShift,
} from './shift-cells';
import { EditTransaction } from './transaction';

const r = (a1: string) => parseRange(a1)!;

function book(): Workbook {
	return createWorkbook({ sheetNames: ['Sheet1', 'Sheet2'] });
}

function seed(wb: Workbook, entries: Record<string, string>, sheet = 0): void {
	const tx = new EditTransaction(wb, 'seed');
	for (const [ref, text] of Object.entries(entries)) {
		const a = r(ref);
		setCellInput(tx, wb, sheet, a.top, a.left, text);
	}
	tx.commit();
}

const cell = (wb: Workbook, ref: string, sheet = 0) => {
	const a = r(ref);
	return getCell(wb.sheets[sheet], a.top, a.left);
};
const val = (wb: Workbook, ref: string, sheet = 0) => cell(wb, ref, sheet)?.value ?? null;
const fx = (wb: Workbook, ref: string, sheet = 0) => cell(wb, ref, sheet)?.formula;

function run(wb: Workbook, fn: (tx: EditTransaction) => void) {
	const tx = new EditTransaction(wb, 'shift');
	fn(tx);
	return tx.commit();
}

function refusal(fn: () => void): CellShiftError {
	try {
		fn();
	} catch (e) {
		if (e instanceof CellShiftError) {
			return e;
		}
		throw e;
	}
	throw new Error('expected a CellShiftError');
}

describe('insertCellsShift: shift cells down', () => {
	it('moves the cells of the band down and leaves other columns alone', () => {
		const wb = book();
		seed(wb, { A3: 'a3', B2: 'b2', B3: 'b3', B4: '4', C3: 'c3' });
		run(wb, (tx) => insertCellsShift(tx, wb, 0, r('B3:B4'), 'down'));
		expect(['A3', 'B2', 'B3', 'B4', 'B5', 'B6', 'C3'].map((a) => val(wb, a))).toEqual([
			'a3',
			'b2',
			null,
			null,
			'b3',
			4,
			'c3',
		]);
	});

	it('gives the inserted cells the format of the cells above (none on the first row)', () => {
		const wb = book();
		const bold = new StylePool(wb.styles).derive(undefined, { font: { bold: true } });
		wb.sheets[0].cells.set(cellKey(1, 1), { value: 'x', style: bold });
		wb.sheets[0].cells.set(cellKey(1, 2), { value: 'y', style: bold });
		run(wb, (tx) => insertCellsShift(tx, wb, 0, r('B3:B4'), 'down'));
		expect(cell(wb, 'B3')).toEqual({ value: null, style: bold });
		expect(cell(wb, 'B4')).toEqual({ value: null, style: bold });
		expect(cell(wb, 'C3')).toBeUndefined();
		run(wb, (tx) => insertCellsShift(tx, wb, 0, r('C1'), 'down'));
		expect(cell(wb, 'C1')).toBeUndefined();
		expect(cell(wb, 'C3')).toEqual({ value: 'y', style: bold });
	});

	it('adjusts formulas with the band rule', () => {
		const wb = book();
		seed(wb, {
			D1: '=B5',
			D2: '=SUM(B1:B5)',
			D3: '=SUM(A1:C5)',
			D4: '=B2',
			D5: '=$B$5*2',
			D6: '=SUM(B:B)',
			B5: '=A5+B4',
		});
		run(wb, (tx) => insertCellsShift(tx, wb, 0, r('B3'), 'down'));
		expect(['D1', 'D2', 'D3', 'D4', 'D5', 'D6'].map((a) => fx(wb, a))).toEqual([
			'B6',
			'SUM(B1:B6)',
			'SUM(A1:C5)',
			'B2',
			'$B$6*2',
			'SUM(B:B)',
		]);
		// The moved formula keeps pointing at the same cells.
		expect(fx(wb, 'B6')).toBe('A5+B5');
		expect(cell(wb, 'B6')?.array?.ref).toEqual(r('B6'));
	});

	it('adjusts references from other sheets and defined names', () => {
		const wb = book();
		seed(wb, { B5: '7' });
		seed(wb, { A1: '=Sheet1!B5+B5', A2: "='Sheet1'!B1:B5" }, 1);
		wb.definedNames.push(
			{ name: 'Data', formula: 'Sheet1!$B$1:$B$5' },
			{ name: 'Wide', formula: 'Sheet1!$A$1:$C$5' },
			{ name: 'Local', formula: 'B5', localSheet: 1 },
		);
		run(wb, (tx) => insertCellsShift(tx, wb, 0, r('B3'), 'down'));
		expect(fx(wb, 'A1', 1)).toBe('Sheet1!B6+B5');
		expect(fx(wb, 'A2', 1)).toBe("'Sheet1'!B1:B6");
		expect(wb.definedNames.map((n) => n.formula)).toEqual([
			'Sheet1!$B$1:$B$6',
			'Sheet1!$A$1:$C$5',
			'B5',
		]);
	});

	it('moves merges inside the block and refuses merges the shift would split', () => {
		const wb = book();
		wb.sheets[0].merges.push(r('B5:B6'), r('D1:E2'));
		run(wb, (tx) => insertCellsShift(tx, wb, 0, r('B3'), 'down'));
		expect(wb.sheets[0].merges).toEqual([r('B6:B7'), r('D1:E2')]);
		wb.sheets[0].merges.push(r('A9:C9'));
		const e = refusal(() => run(wb, (tx) => insertCellsShift(tx, wb, 0, r('B3'), 'down')));
		expect(e.reason).toBe('merge');
		expect(e.message).toBe("We can't do that to a merged cell.");
	});

	it('refuses to push nonblank cells off the grid but drops formats', () => {
		const wb = book();
		const sheet = wb.sheets[0];
		sheet.cells.set(cellKey(MAX_ROW_INDEX, 1), { value: null, style: 1 });
		run(wb, (tx) => insertCellsShift(tx, wb, 0, r('B3'), 'down'));
		expect(sheet.cells.has(cellKey(MAX_ROW_INDEX, 1))).toBe(false);
		sheet.cells.set(cellKey(MAX_ROW_INDEX, 1), { value: 'last' });
		const e = refusal(() => run(wb, (tx) => insertCellsShift(tx, wb, 0, r('B3'), 'down')));
		expect(e.reason).toBe('dataLoss');
		expect(e.message).toBe(CELL_SHIFT_MESSAGES.dataLoss);
		expect(e.message).toContain('cannot shift nonblank cells off of the worksheet');
		// Cells above the insertion point do not move.
		expect(checkCellShift(wb, 0, r(`B${MAX_ROW_INDEX + 1}`), 'down')?.reason).toBe('dataLoss');
		expect(checkCellShift(wb, 0, r('C3'), 'down')).toBeNull();
	});

	it('refuses to split legacy array formulas and moves whole ones', () => {
		const wb = book();
		const sheet = wb.sheets[0];
		const anchor = cellKey(4, 1);
		sheet.cells.set(anchor, {
			value: 1,
			formula: 'A5:A6',
			array: { ref: r('B5:B6'), dynamic: false },
		});
		sheet.cells.set(cellKey(5, 1), { value: 2, arrayAnchor: anchor });
		run(wb, (tx) => insertCellsShift(tx, wb, 0, r('B3'), 'down'));
		expect(cell(wb, 'B6')).toMatchObject({ formula: 'A5:A6', array: { ref: r('B6:B7') } });
		expect(cell(wb, 'B7')).toEqual({ value: 2, arrayAnchor: cellKey(5, 1) });
		const e = refusal(() => run(wb, (tx) => insertCellsShift(tx, wb, 0, r('B7'), 'down')));
		expect(e.reason).toBe('array');
		expect(e.message).toBe("You can't change part of an array.");
	});

	it('lets a split spill range spill again from its formula cell', () => {
		const wb = book();
		const sheet = wb.sheets[0];
		const anchor = cellKey(0, 1);
		sheet.cells.set(anchor, {
			value: 1,
			formula: 'SEQUENCE(4)',
			array: { ref: r('B1:B4'), dynamic: true },
		});
		for (let row = 1; row < 4; row++) {
			sheet.cells.set(cellKey(row, 1), {
				value: row + 1,
				arrayAnchor: anchor,
				style: row === 3 ? 1 : undefined,
			});
		}
		run(wb, (tx) => insertCellsShift(tx, wb, 0, r('B3'), 'down'));
		expect(cell(wb, 'B1')).toMatchObject({ formula: 'SEQUENCE(4)', array: { ref: r('B1') } });
		expect(cell(wb, 'B2')).toBeUndefined();
		expect(cell(wb, 'B5')).toEqual({ value: null, style: 1 });
	});
});

describe('insertCellsShift: shift cells right', () => {
	it('moves the cells of the rows right with formats from the left', () => {
		const wb = book();
		seed(wb, {
			A2: 'a',
			B2: 'b',
			C2: 'c',
			B3: 'x',
			E5: '=C2',
			E6: '=SUM(A2:D2)',
			E7: '=SUM(A1:D3)',
		});
		const fill = new StylePool(wb.styles).derive(undefined, {
			fill: { type: 'pattern', patternType: 'solid' },
		});
		wb.sheets[0].cells.set(cellKey(1, 0), { value: 'a', style: fill });
		run(wb, (tx) => insertCellsShift(tx, wb, 0, r('B2'), 'right'));
		expect(['A2', 'B2', 'C2', 'D2', 'B3'].map((a) => val(wb, a))).toEqual([
			'a',
			null,
			'b',
			'c',
			'x',
		]);
		expect(cell(wb, 'B2')).toEqual({ value: null, style: fill });
		expect(['E5', 'E6', 'E7'].map((a) => fx(wb, a))).toEqual(['D2', 'SUM(A2:E2)', 'SUM(A1:D3)']);
	});
});

describe('deleteCellsShift', () => {
	it('shift up: removes the cells, moves the cells below up and kills references', () => {
		const wb = book();
		seed(wb, { B3: 'gone', B4: 'b4', B6: 'b6', C4: 'c4', D1: '=B3', D2: '=SUM(B1:B5)', D3: '=B6' });
		run(wb, (tx) => deleteCellsShift(tx, wb, 0, r('B3'), 'up'));
		expect(['B3', 'B5', 'B6', 'C4'].map((a) => val(wb, a))).toEqual(['b4', 'b6', null, 'c4']);
		expect(['D1', 'D2', 'D3'].map((a) => fx(wb, a))).toEqual(['#REF!', 'SUM(B1:B4)', 'B5']);
	});

	it('shift left: transposed', () => {
		const wb = book();
		seed(wb, { B2: 'gone', C2: 'c2', D2: 'd2', C3: 'c3', A5: '=C2', A6: '=B2&D2' });
		run(wb, (tx) => deleteCellsShift(tx, wb, 0, r('B2'), 'left'));
		expect(['B2', 'C2', 'D2', 'C3'].map((a) => val(wb, a))).toEqual(['c2', 'd2', null, 'c3']);
		expect(fx(wb, 'A5')).toBe('B2');
		expect(fx(wb, 'A6')).toBe('#REF!&C2');
	});

	it('removes merges inside the deleted cells and moves those below', () => {
		const wb = book();
		wb.sheets[0].merges.push(r('B3:C3'), r('B6:C7'));
		run(wb, (tx) => deleteCellsShift(tx, wb, 0, r('B3:C4'), 'up'));
		expect(wb.sheets[0].merges).toEqual([r('B4:C5')]);
		wb.sheets[0].merges.push(r('C8:D8'));
		expect(checkCellShift(wb, 0, r('B3:C4'), 'up')?.reason).toBe('merge');
	});
});

describe('refusals and undo', () => {
	it('changes nothing when refusing', () => {
		const wb = book();
		seed(wb, { B5: '1', D1: '=B5' });
		wb.sheets[0].merges.push(r('A9:C9'));
		wb.sheets.forEach((s) => (s.dirty = false));
		const tx = new EditTransaction(wb, 'shift');
		expect(() => insertCellsShift(tx, wb, 0, r('B3'), 'down')).toThrow(CellShiftError);
		const entry = tx.commit();
		expect(entry.sheets).toEqual([]);
		expect(val(wb, 'B5')).toBe(1);
		expect(fx(wb, 'D1')).toBe('B5');
		expect(wb.sheets[0].dirty).toBe(false);
	});

	it('undoes and redoes everything', () => {
		const wb = book();
		seed(wb, { B3: 'b3', B4: '=B3', D1: '=SUM(B1:B5)' });
		seed(wb, { A1: '=Sheet1!B4' }, 1);
		wb.definedNames.push({ name: 'Data', formula: 'Sheet1!$B$1:$B$5' });
		wb.sheets[0].merges.push(r('B8:B9'));
		const before = new Map(wb.sheets[0].cells);
		const entry = run(wb, (tx) => insertCellsShift(tx, wb, 0, r('B2:B3'), 'down'));
		expect(entry.structural).toBe(true);
		expect(fx(wb, 'B6')).toBe('B5');
		const after = new Map(wb.sheets[0].cells);
		entry.undo();
		expect(wb.sheets[0].cells).toEqual(before);
		expect(fx(wb, 'A1', 1)).toBe('Sheet1!B4');
		expect(wb.definedNames[0].formula).toBe('Sheet1!$B$1:$B$5');
		expect(wb.sheets[0].merges).toEqual([r('B8:B9')]);
		entry.redo();
		expect(wb.sheets[0].cells).toEqual(after);
		expect(fx(wb, 'A1', 1)).toBe('Sheet1!B6');
		expect(fx(wb, 'D1')).toBe('SUM(B1:B7)');
		expect(wb.definedNames[0].formula).toBe('Sheet1!$B$1:$B$7');
		expect(wb.sheets[0].merges).toEqual([r('B10:B11')]);
	});

	it('undoes a deletion', () => {
		const wb = book();
		seed(wb, { B2: 'x', B3: 'y', C1: '=B2' });
		const entry = run(wb, (tx) => deleteCellsShift(tx, wb, 0, r('B2'), 'up'));
		expect(val(wb, 'B2')).toBe('y');
		expect(fx(wb, 'C1')).toBe('#REF!');
		entry.undo();
		expect([val(wb, 'B2'), val(wb, 'B3'), fx(wb, 'C1')]).toEqual(['x', 'y', 'B2']);
		entry.redo();
		expect([val(wb, 'B2'), val(wb, 'B3'), fx(wb, 'C1')]).toEqual(['y', null, '#REF!']);
	});
});
