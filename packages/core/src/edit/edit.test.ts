import { describe, expect, it } from 'vitest';

import { cellKey } from '../address/cell-address';
import { parseRange } from '../address/range';
import { createWorkbook } from '../sheet/workbook';
import { getCell } from '../sheet/worksheet';
import { numberFormatCode } from '../styles/builtin-number-formats';
import type { Workbook } from '../types/workbook';
import { ERRORS } from '../values/errors';
import {
	addSheet,
	applyBorderPreset,
	applyStylePatch,
	clearRanges,
	deleteSheet,
	deleteColumns,
	EditTransaction,
	fillRange,
	filterColumnValues,
	insertRows,
	setFilterColumn,
	sortRange,
	mergeRange,
	moveSheet,
	parseCellInput,
	renameSheet,
	renameSheetReferences,
	setCellInput,
	setColumnWidth,
	setFreezePanes,
	unmergeRange,
} from './index';

const r = (a1: string) => parseRange(a1)!;

function numFmtOf(wb: Workbook, row: number, col: number): string {
	const cell = getCell(wb.sheets[0], row, col);
	return numberFormatCode(wb.styles.cellXfs[cell?.style ?? 0].numFmtId, wb.styles.numFmts);
}

describe('parseCellInput', () => {
	it('recognises Excel entry kinds', () => {
		expect(parseCellInput('')).toEqual({ kind: 'blank', value: null });
		expect(parseCellInput("'00123")).toEqual({ kind: 'text', value: '00123', quotePrefix: true });
		expect(parseCellInput('=SUM(A1:A3)')).toMatchObject({ kind: 'formula', formula: 'SUM(A1:A3)' });
		expect(parseCellInput('+A1')).toMatchObject({ kind: 'formula', formula: '+A1' });
		expect(parseCellInput('-5')).toMatchObject({ kind: 'number', value: -5 });
		expect(parseCellInput('true')).toEqual({ kind: 'boolean', value: true });
		expect(parseCellInput('#N/A')).toEqual({ kind: 'error', value: ERRORS.NA });
		expect(parseCellInput('12%')).toMatchObject({ kind: 'number', value: 0.12, numFmt: '0%' });
		expect(parseCellInput('hello')).toEqual({ kind: 'text', value: 'hello' });
	});

	it('keeps text in text-formatted cells and applies automatic percent entry', () => {
		expect(parseCellInput('=A1', { currentFormat: '@' })).toEqual({ kind: 'text', value: '=A1' });
		expect(parseCellInput('5', { currentFormat: '0%' })).toMatchObject({ value: 0.05 });
	});
});

describe('EditTransaction with setCellInput', () => {
	it('stores values with implied formats and undoes / redoes exactly', () => {
		const wb = createWorkbook();
		const tx = new EditTransaction(wb, 'Typing');
		setCellInput(tx, wb, 0, 0, 0, '1,234.5');
		setCellInput(tx, wb, 0, 1, 0, '=A1*2');
		const entry = tx.commit();
		expect(getCell(wb.sheets[0], 0, 0)?.value).toBe(1234.5);
		expect(numFmtOf(wb, 0, 0)).toBe('#,##0.00');
		expect(getCell(wb.sheets[0], 1, 0)).toMatchObject({
			formula: 'A1*2',
			array: { dynamic: true },
		});
		expect(entry.changedCells).toHaveLength(2);

		entry.undo();
		expect(wb.sheets[0].cells.size).toBe(0);
		entry.redo();
		expect(getCell(wb.sheets[0], 0, 0)?.value).toBe(1234.5);
	});

	it('does not override an explicit number format', () => {
		const wb = createWorkbook();
		const tx = new EditTransaction(wb, 'fmt');
		applyStylePatch(tx, wb, 0, [r('A1')], { numFmt: '0.000' });
		setCellInput(tx, wb, 0, 0, 0, '$5');
		expect(numFmtOf(wb, 0, 0)).toBe('0.000');
	});

	it('rejects edits to part of a CSE array', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0];
		sheet.cells.set(cellKey(0, 0), {
			value: 1,
			formula: '{1,2}',
			array: { ref: r('A1:B1'), dynamic: false },
		});
		sheet.cells.set(cellKey(0, 1), { value: 2, arrayAnchor: cellKey(0, 0) });
		const tx = new EditTransaction(wb, 'x');
		expect(setCellInput(tx, wb, 0, 0, 1, '5').error).toBe('array-part');
	});
});

describe('style operations', () => {
	it('derives shared formats and supports whole columns', () => {
		const wb = createWorkbook();
		const tx = new EditTransaction(wb, 'Bold');
		applyStylePatch(tx, wb, 0, [r('A1:B2')], { font: { bold: true } });
		const a1 = getCell(wb.sheets[0], 0, 0)!.style!;
		expect(getCell(wb.sheets[0], 1, 1)!.style).toBe(a1);
		expect(wb.styles.fonts[wb.styles.cellXfs[a1].fontId].bold).toBe(true);

		applyStylePatch(tx, wb, 0, [r('C:C')], {
			fill: { type: 'pattern', patternType: 'solid', fgColor: { rgb: 'FFFFFF00' } },
		});
		expect(wb.sheets[0].cols[0]).toMatchObject({ min: 2, max: 2 });
		const entry = tx.commit();
		entry.undo();
		expect(wb.sheets[0].cols).toHaveLength(0);
		expect(wb.sheets[0].cells.size).toBe(0);
	});

	it('applies outside and inside border presets', () => {
		const wb = createWorkbook();
		const tx = new EditTransaction(wb, 'Borders');
		applyBorderPreset(tx, wb, 0, r('A1:B2'), 'outside');
		const border = (a: string) => {
			const ad = parseRange(a)!;
			const cell = getCell(wb.sheets[0], ad.top, ad.left)!;
			return wb.styles.borders[wb.styles.cellXfs[cell.style!].borderId];
		};
		expect(border('A1')).toMatchObject({ top: { style: 'thin' }, left: { style: 'thin' } });
		expect(border('A1').right).toBeUndefined();
		expect(border('B2')).toMatchObject({ bottom: { style: 'thin' }, right: { style: 'thin' } });
	});
});

describe('merge, clear, layout', () => {
	it('merges keeping the top-left value and unmerges', () => {
		const wb = createWorkbook();
		const tx = new EditTransaction(wb, 'Merge');
		setCellInput(tx, wb, 0, 0, 0, 'Title');
		setCellInput(tx, wb, 0, 0, 1, 'lost');
		mergeRange(tx, wb, 0, r('A1:C1'), { center: true });
		expect(wb.sheets[0].merges).toEqual([r('A1:C1')]);
		expect(getCell(wb.sheets[0], 0, 1)?.value ?? null).toBeNull();
		unmergeRange(tx, wb, 0, r('B1'));
		expect(wb.sheets[0].merges).toHaveLength(0);
	});

	it('clears contents but keeps formats', () => {
		const wb = createWorkbook();
		const tx = new EditTransaction(wb, 'Clear');
		setCellInput(tx, wb, 0, 0, 0, '50%');
		clearRanges(tx, wb, 0, [r('A1')], 'contents');
		expect(getCell(wb.sheets[0], 0, 0)).toMatchObject({ value: null });
		expect(numFmtOf(wb, 0, 0)).toBe('0%');
		clearRanges(tx, wb, 0, [r('A1')], 'all');
		expect(getCell(wb.sheets[0], 0, 0)).toBeUndefined();
	});

	it('sets widths and freeze panes', () => {
		const wb = createWorkbook();
		const tx = new EditTransaction(wb, 'Layout');
		setColumnWidth(tx, wb, 0, 1, 3, 20);
		expect(wb.sheets[0].cols).toEqual([{ min: 1, max: 3, width: 20, customWidth: true }]);
		setFreezePanes(tx, wb, 0, { row: 1, col: 1 });
		expect(wb.sheets[0].view.pane).toMatchObject({ xSplit: 1, ySplit: 1, state: 'frozen' });
	});
});

describe('sheet operations', () => {
	it('adds, renames (updating references), moves and deletes sheets', () => {
		const wb = createWorkbook();
		const tx = new EditTransaction(wb, 'Sheets');
		const idx = addSheet(tx, wb);
		expect(wb.sheets[idx].name).toBe('Sheet2');
		setCellInput(tx, wb, 0, 0, 0, '=Sheet2!A1+SUM(Sheet2!B1:B3)&"Sheet2!"');
		renameSheet(tx, wb, idx, 'Data 2024');
		expect(getCell(wb.sheets[0], 0, 0)?.formula).toBe(
			"'Data 2024'!A1+SUM('Data 2024'!B1:B3)&\"Sheet2!\"",
		);
		moveSheet(tx, wb, 1, 0);
		expect(wb.sheets[0].name).toBe('Data 2024');
		deleteSheet(tx, wb, 0);
		expect(wb.sheets.map((s) => s.name)).toEqual(['Sheet1']);
		const entry = tx.commit();
		entry.undo();
		expect(wb.sheets.map((s) => s.name)).toEqual(['Sheet1']);
		expect(wb.sheets).toHaveLength(1);
	});

	it('rewrites quoted references', () => {
		expect(renameSheetReferences("'My Sheet'!A1+mysheet2!A1", 'My Sheet', 'Data')).toBe(
			'Data!A1+mysheet2!A1',
		);
		expect(renameSheetReferences('Sheet1!A1', 'Sheet1', 'Q1')).toBe("'Q1'!A1");
	});
});

describe('fillRange (AutoFill)', () => {
	const setup = (values: Array<string>) => {
		const wb = createWorkbook();
		const tx = new EditTransaction(wb, 'fill');
		values.forEach((v, i) => setCellInput(tx, wb, 0, i, 0, v));
		return { wb, tx };
	};
	const col = (wb: Workbook, n: number) =>
		Array.from({ length: n }, (_, i) => getCell(wb.sheets[0], i, 0)?.value ?? null);

	it('continues linear number series and repeats single numbers', () => {
		const { wb, tx } = setup(['1', '3']);
		fillRange(tx, wb, 0, r('A1:A2'), r('A1:A5'));
		expect(col(wb, 5)).toEqual([1, 3, 5, 7, 9]);
		const s2 = setup(['7']);
		fillRange(s2.tx, s2.wb, 0, r('A1'), r('A1:A3'));
		expect(col(s2.wb, 3)).toEqual([7, 7, 7]);
	});

	it('increments text with numbers and cycles day names', () => {
		const { wb, tx } = setup(['Item 9']);
		fillRange(tx, wb, 0, r('A1'), r('A1:A3'));
		expect(col(wb, 3)).toEqual(['Item 9', 'Item 10', 'Item 11']);
		const d = setup(['Fri']);
		fillRange(d.tx, d.wb, 0, r('A1'), r('A1:A4'));
		expect(col(d.wb, 4)).toEqual(['Fri', 'Sat', 'Sun', 'Mon']);
	});

	it('steps dates by day or by month', () => {
		const { wb, tx } = setup(['1/31/2024']);
		fillRange(tx, wb, 0, r('A1'), r('A1:A2'));
		expect(col(wb, 2)).toEqual([45322, 45323]);
		const m = setup(['1/15/2024', '2/15/2024']);
		fillRange(m.tx, m.wb, 0, r('A1:A2'), r('A1:A4'));
		expect(col(m.wb, 4)).toEqual([45306, 45337, 45366, 45397]);
	});

	it('copies formulas with relative references', () => {
		const { wb, tx } = setup(['=B1*2']);
		fillRange(tx, wb, 0, r('A1'), r('A1:A3'));
		expect(getCell(wb.sheets[0], 2, 0)?.formula).toBe('B3*2');
	});
});

describe('structure operations', () => {
	it('inserts rows, shifting cells, formulas (all sheets), merges and names', () => {
		const wb = createWorkbook({ sheetNames: ['Data', 'Calc'] });
		const tx = new EditTransaction(wb, 'setup');
		setCellInput(tx, wb, 0, 0, 0, '1');
		setCellInput(tx, wb, 0, 1, 0, '2');
		setCellInput(tx, wb, 0, 2, 0, '=SUM(A1:A2)');
		setCellInput(tx, wb, 1, 0, 0, '=Data!A2*10');
		mergeRange(tx, wb, 0, r('B2:C3'));
		wb.definedNames.push({ name: 'Items', formula: 'Data!$A$1:$A$2' });
		insertRows(tx, wb, 0, 1, 2);
		expect(getCell(wb.sheets[0], 3, 0)?.value).toBe(2);
		expect(getCell(wb.sheets[0], 4, 0)?.formula).toBe('SUM(A1:A4)');
		expect(getCell(wb.sheets[1], 0, 0)?.formula).toBe('Data!A4*10');
		expect(wb.sheets[0].merges).toEqual([r('B4:C5')]);
		expect(wb.definedNames[0].formula).toBe('Data!$A$1:$A$4');
		const entry = tx.commit();
		entry.undo();
		expect(wb.sheets[0].cells.size).toBe(0);
	});

	it('deletes columns producing #REF! for removed references', () => {
		const wb = createWorkbook();
		const tx = new EditTransaction(wb, 'x');
		setCellInput(tx, wb, 0, 0, 0, '5');
		setCellInput(tx, wb, 0, 0, 2, '=A1+B1');
		deleteColumns(tx, wb, 0, 0, 1);
		expect(getCell(wb.sheets[0], 0, 1)?.formula).toBe('#REF!+A1');
	});
});

describe('sortRange', () => {
	it('sorts rows with Excel ordering, keeping the header and adjusting formulas', () => {
		const wb = createWorkbook();
		const tx = new EditTransaction(wb, 'sort');
		const rows = [
			['Name', 'Qty'],
			['b', '3'],
			['A', '1'],
			['', '7'],
			['c', '2'],
		];
		rows.forEach((row, i) => row.forEach((v, j) => setCellInput(tx, wb, 0, i, j, v)));
		setCellInput(tx, wb, 0, 1, 2, '=B2*2');
		sortRange(tx, wb, 0, r('A1:C5'), [{ index: 0 }], { hasHeader: true });
		const names = [1, 2, 3, 4].map((i) => getCell(wb.sheets[0], i, 0)?.value ?? null);
		expect(names).toEqual(['A', 'b', 'c', null]);
		expect(getCell(wb.sheets[0], 2, 2)?.formula).toBe('B3*2');
	});
});

describe('filters', () => {
	it('hides rows failing value-list criteria and shows them again when cleared', () => {
		const wb = createWorkbook();
		const tx = new EditTransaction(wb, 'filter');
		[
			['Fruit', 'Qty'],
			['Apple', '3'],
			['Pear', '5'],
			['apple', '7'],
			['', '1'],
		].forEach((row, i) => row.forEach((v, j) => setCellInput(tx, wb, 0, i, j, v)));
		wb.sheets[0].autoFilter = { ref: r('A1:B5'), columns: [] };
		setFilterColumn(tx, wb, 0, { kind: 'sheet' }, 0, { filters: { values: ['Apple'] } });
		const hidden = [1, 2, 3, 4].map((row) => wb.sheets[0].rows.get(row)?.hidden ?? false);
		expect(hidden).toEqual([false, true, false, true]);
		expect(filterColumnValues(wb, wb.sheets[0], r('A1:B5'), 0)).toEqual({
			values: ['Apple', 'Pear'],
			hasBlank: true,
		});
		setFilterColumn(tx, wb, 0, { kind: 'sheet' }, 0, null);
		expect([1, 2, 3, 4].some((row) => wb.sheets[0].rows.get(row)?.hidden)).toBe(false);
	});

	it('supports custom numeric filters', () => {
		const wb = createWorkbook();
		const tx = new EditTransaction(wb, 'filter');
		['Qty', '3', '5', '7'].forEach((v, i) => setCellInput(tx, wb, 0, i, 0, v));
		wb.sheets[0].autoFilter = { ref: r('A1:A4'), columns: [] };
		setFilterColumn(tx, wb, 0, { kind: 'sheet' }, 0, {
			customFilters: { filters: [{ operator: 'greaterThan', val: '4' }] },
		});
		expect([1, 2, 3].map((row) => wb.sheets[0].rows.get(row)?.hidden ?? false)).toEqual([
			true,
			false,
			false,
		]);
	});
});
