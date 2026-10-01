import { describe, expect, it } from 'vitest';

import { parseRange } from '../address/range';
import { getPrintArea, getPrintTitles } from '../print/print-names';
import { createWorkbook } from '../sheet/workbook';
import type { Workbook } from '../types/workbook';
import {
	addToPrintArea,
	canRemovePageBreak,
	clearPrintArea,
	EditTransaction,
	insertPageBreak,
	insertRows,
	deleteRows,
	removePageBreak,
	resetAllPageBreaks,
	setFitTo,
	setHeaderFooter,
	setMargins,
	setOrientation,
	setPageSetup,
	setPaperSize,
	setPrintArea,
	setPrintGridlines,
	setPrintHeadings,
	setPrintTitles,
	setScale,
	shiftPageBreaks,
} from './index';

const r = (ref: string) => parseRange(ref)!;

function run(wb: Workbook, fn: (tx: EditTransaction) => void) {
	const tx = new EditTransaction(wb, 'test');
	fn(tx);
	return tx.commit();
}

describe('page setup ops', () => {
	it('sets orientation and paper size, undoably', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0];
		sheet.dirty = false;
		const entry = run(wb, (tx) => {
			setOrientation(tx, wb, 0, 'landscape');
			setPaperSize(tx, wb, 0, 9);
		});
		expect(sheet.pageSetup).toEqual({ orientation: 'landscape', paperSize: 9 });
		expect(sheet.dirty).toBe(true);
		entry.undo();
		expect(sheet.pageSetup).toBeUndefined();
		entry.redo();
		expect(sheet.pageSetup?.orientation).toBe('landscape');
	});

	it('applies margin presets and custom margins', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0];
		const entry = run(wb, (tx) => setMargins(tx, wb, 0, 'wide'));
		expect(sheet.pageMargins).toEqual({
			top: 1,
			bottom: 1,
			left: 1,
			right: 1,
			header: 0.5,
			footer: 0.5,
		});
		run(wb, (tx) => setMargins(tx, wb, 0, 'narrow'));
		expect(sheet.pageMargins?.left).toBe(0.25);
		run(wb, (tx) =>
			setMargins(tx, wb, 0, {
				left: 2,
				right: -1,
				top: 1,
				bottom: 1,
				header: 0,
				footer: Number.NaN,
			}),
		);
		expect(sheet.pageMargins).toEqual({
			left: 2,
			right: 0,
			top: 1,
			bottom: 1,
			header: 0,
			footer: 0,
		});
		entry.undo();
		expect(sheet.pageMargins).toBeUndefined();
	});

	it('switches between Adjust to and Fit to', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0];
		run(wb, (tx) => setFitTo(tx, wb, 0, 1, 0));
		expect(sheet.pageSetup).toEqual({ fitToPage: true, fitToWidth: 1, fitToHeight: 0 });
		run(wb, (tx) => setScale(tx, wb, 0, 1000));
		expect(sheet.pageSetup).toEqual({
			fitToPage: false,
			fitToWidth: 1,
			fitToHeight: 0,
			scale: 400,
		});
		run(wb, (tx) => setFitTo(tx, wb, 0, 2, 3));
		expect(sheet.pageSetup).toMatchObject({ fitToPage: true, fitToWidth: 2, fitToHeight: 3 });
		run(wb, (tx) => setFitTo(tx, wb, 0, 0, 0));
		expect(sheet.pageSetup).toEqual({ fitToPage: false });
		run(wb, (tx) => setPageSetup(tx, wb, 0, { fitToPage: undefined, pageOrder: 'overThenDown' }));
		expect(sheet.pageSetup).toEqual({ pageOrder: 'overThenDown' });
	});

	it('turns gridline and heading printing on and off', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0];
		run(wb, (tx) => {
			setPrintGridlines(tx, wb, 0, true);
			setPrintHeadings(tx, wb, 0, true);
		});
		expect(sheet.printOptions).toEqual({ gridLines: true, headings: true });
		const entry = run(wb, (tx) => {
			setPrintGridlines(tx, wb, 0, false);
			setPrintHeadings(tx, wb, 0, false);
		});
		expect(sheet.printOptions).toBeUndefined();
		entry.undo();
		expect(sheet.printOptions).toEqual({ gridLines: true, headings: true });
	});

	it('sets headers and footers, dropping empty parts', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0];
		run(wb, (tx) =>
			setHeaderFooter(tx, wb, 0, {
				oddHeader: '&CPage &P',
				oddFooter: '',
				differentFirst: false,
				scaleWithDoc: false,
			}),
		);
		expect(sheet.headerFooter).toEqual({ oddHeader: '&CPage &P', scaleWithDoc: false });
		run(wb, (tx) => setHeaderFooter(tx, wb, 0, { oddHeader: '' }));
		expect(sheet.headerFooter).toBeUndefined();
	});
});

describe('print area and titles', () => {
	it('sets, extends and clears the print area name', () => {
		const wb = createWorkbook({ sheetNames: ['Data', 'My Sheet'] });
		wb.dirty = false;
		const entry = run(wb, (tx) => setPrintArea(tx, wb, 1, [r('A1:C10')]));
		expect(wb.definedNames).toEqual([
			{ name: '_xlnm.Print_Area', formula: "'My Sheet'!$A$1:$C$10", localSheet: 1, builtin: true },
		]);
		expect(wb.dirty).toBe(true);
		// D1:E10 continues the area; G1 starts a new one; B2 is already inside.
		run(wb, (tx) => addToPrintArea(tx, wb, 1, [r('D1:E10'), r('G1'), r('B2')]));
		expect(getPrintArea(wb, 1)).toEqual([r('A1:E10'), r('G1')]);
		run(wb, (tx) => addToPrintArea(tx, wb, 1, [r('A11:E12')]));
		expect(getPrintArea(wb, 1)).toEqual([r('A1:E12'), r('G1')]);
		run(wb, (tx) => clearPrintArea(tx, wb, 1));
		expect(wb.definedNames).toEqual([]);
		entry.undo();
		expect(wb.definedNames).toEqual([]);
	});

	it('keeps defined names sorted and undoes print area changes', () => {
		const wb = createWorkbook();
		wb.definedNames.push({ name: 'Alpha', formula: '1' }, { name: 'Zeta', formula: '2' });
		const entry = run(wb, (tx) => setPrintArea(tx, wb, 0, [r('A1:B2'), r('D:E')]));
		expect(wb.definedNames.map((n) => n.name)).toEqual(['Alpha', '_xlnm.Print_Area', 'Zeta']);
		expect(wb.definedNames[1].formula).toBe('Sheet1!$A$1:$B$2,Sheet1!$D:$E');
		entry.undo();
		expect(wb.definedNames.map((n) => n.name)).toEqual(['Alpha', 'Zeta']);
	});

	it('sets and clears print titles', () => {
		const wb = createWorkbook();
		run(wb, (tx) =>
			setPrintTitles(tx, wb, 0, { rows: { start: 2, end: 0 }, cols: { start: 1, end: 1 } }),
		);
		expect(wb.definedNames[0]).toMatchObject({
			name: '_xlnm.Print_Titles',
			formula: 'Sheet1!$B:$B,Sheet1!$1:$3',
			localSheet: 0,
		});
		expect(getPrintTitles(wb, 0)).toEqual({
			rows: { start: 0, end: 2 },
			cols: { start: 1, end: 1 },
		});
		run(wb, (tx) => setPrintTitles(tx, wb, 0, { rows: { start: 0, end: 0 } }));
		expect(wb.definedNames[0].formula).toBe('Sheet1!$1:$1');
		run(wb, (tx) => setPrintTitles(tx, wb, 0, {}));
		expect(wb.definedNames).toEqual([]);
	});
});

describe('page breaks', () => {
	it('inserts breaks above and left of the active cell', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0];
		const entry = run(wb, (tx) => insertPageBreak(tx, wb, 0, { row: 9, col: 3 }));
		expect(sheet.rowBreaks).toEqual([9]);
		expect(sheet.colBreaks).toEqual([3]);
		run(wb, (tx) => insertPageBreak(tx, wb, 0, { row: 4, col: 0 }));
		expect(sheet.rowBreaks).toEqual([4, 9]);
		expect(sheet.colBreaks).toEqual([3]);
		run(wb, (tx) => insertPageBreak(tx, wb, 0, { row: 0, col: 0 }));
		run(wb, (tx) => insertPageBreak(tx, wb, 0, { row: 20, col: 6 }, 'row'));
		expect(sheet.rowBreaks).toEqual([4, 9, 20]);
		expect(sheet.colBreaks).toEqual([3]);
		entry.undo();
		expect(sheet.rowBreaks).toBeUndefined();
		expect(sheet.colBreaks).toBeUndefined();
	});

	it('removes the breaks next to the active cell and resets all', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0];
		sheet.rowBreaks = [4, 9];
		sheet.colBreaks = [3];
		expect(canRemovePageBreak(sheet, { row: 9, col: 1 })).toBe(true);
		expect(canRemovePageBreak(sheet, { row: 8, col: 1 })).toBe(false);
		run(wb, (tx) => removePageBreak(tx, wb, 0, { row: 9, col: 3 }));
		expect(sheet.rowBreaks).toEqual([4]);
		expect(sheet.colBreaks).toEqual([]);
		const entry = run(wb, (tx) => resetAllPageBreaks(tx, wb, 0));
		expect(sheet.rowBreaks).toEqual([]);
		entry.undo();
		expect(sheet.rowBreaks).toEqual([4]);
	});

	it('moves breaks with inserted and deleted rows', () => {
		expect(shiftPageBreaks([3, 10], true, 5, 2)).toEqual([3, 12]);
		expect(shiftPageBreaks([3, 10], false, 2, 2)).toEqual([8]);
		expect(shiftPageBreaks(undefined, true, 0, 1)).toBeUndefined();
		const wb = createWorkbook();
		const sheet = wb.sheets[0];
		sheet.rowBreaks = [10];
		run(wb, (tx) => insertRows(tx, wb, 0, 2, 3));
		expect(sheet.rowBreaks).toEqual([13]);
		run(wb, (tx) => deleteRows(tx, wb, 0, 0, 1));
		expect(sheet.rowBreaks).toEqual([12]);
	});
});
