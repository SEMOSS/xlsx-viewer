import { describe, expect, it } from 'vitest';

import { cellKey } from '../address/cell-address';
import { createWorkbook } from '../sheet/workbook';
import { putCell } from '../sheet/worksheet';
import { writeWorkbook } from './index';
import { defaultCfFormula, fileFormula } from './formulas';
import { openPackage, partRoot, pick, pickAll } from './test-utils/package-check';

describe('fileFormula', () => {
	it.each([
		['SUM(A1:A3)', 'SUM(A1:A3)'],
		['XLOOKUP(1,A:A,B:B)', '_xlfn.XLOOKUP(1,A:A,B:B)'],
		['FILTER(A1:A9,B1:B9>0)', '_xlfn._xlws.FILTER(A1:A9,B1:B9>0)'],
		// Display-only syntax without a function call (regression: no "(" shortcut).
		['A1#', '_xlfn.ANCHORARRAY(A1)'],
		['@A1', '_xlfn.SINGLE(A1)'],
		['[@Qty]*2', '[[#This Row],[Qty]]*2'],
		['"a#b"&"@x"', '"a#b"&"@x"'],
		['Sheet1!B2:B4', 'Sheet1!B2:B4'],
	])('%s -> %s', (display, file) => {
		expect(fileFormula(display)).toBe(file);
	});

	it('builds the formulas Excel stores for text and period rules', () => {
		expect(defaultCfFormula({ type: 'endsWith', priority: 1, text: 'z' }, 'C3')).toBe(
			'RIGHT(C3,LEN("z"))="z"',
		);
		expect(defaultCfFormula({ type: 'notContainsErrors', priority: 1 }, 'A1')).toBe(
			'NOT(ISERROR(A1))',
		);
		expect(defaultCfFormula({ type: 'timePeriod', priority: 1, timePeriod: 'today' }, 'A1')).toBe(
			'FLOOR(A1,1)=TODAY()',
		);
		expect(defaultCfFormula({ type: 'cellIs', priority: 1 }, 'A1')).toBeUndefined();
	});
});

describe('every formula the writer emits is in file form', () => {
	it('rewrites spill refs, implicit intersection and this-row references everywhere', async () => {
		const wb = createWorkbook();
		const s = wb.sheets[0];
		putCell(s, 0, 0, {
			value: 1,
			formula: 'SEQUENCE(3)',
			array: { ref: { top: 0, left: 0, bottom: 2, right: 0 }, dynamic: true },
		});
		putCell(s, 0, 1, { value: 6, formula: 'A1#' });
		putCell(s, 1, 1, { value: 1, formula: '@A1' });
		putCell(s, 4, 0, { value: 'Qty' });
		putCell(s, 4, 1, { value: 'Double' });
		putCell(s, 5, 0, { value: 2 });
		putCell(s, 5, 1, { value: 4, formula: 'Items[@Qty]*2' });
		s.tables = [
			{
				id: 1,
				name: 'Items',
				displayName: 'Items',
				ref: { top: 4, left: 0, bottom: 5, right: 1 },
				headerRowCount: 1,
				totalsRowCount: 0,
				columns: [
					{ id: 1, name: 'Qty' },
					{ id: 2, name: 'Double', calculatedColumnFormula: '[@Qty]*2' },
				],
			},
		];
		s.conditionalFormats = [
			{
				ranges: [{ top: 0, left: 1, bottom: 2, right: 1 }],
				rules: [{ type: 'expression', priority: 1, formulas: ['@A1>1'] }],
			},
		];
		s.dataValidations = [
			{ ranges: [{ top: 8, left: 0, bottom: 8, right: 0 }], type: 'list', formula1: 'A1#' },
		];
		wb.definedNames = [{ name: 'Spill', formula: 'Sheet1!$A$1#' }];
		const pkg = await openPackage(await writeWorkbook(wb));
		const sheet = await partRoot(pkg, 'xl/worksheets/sheet1.xml');
		const formulas = pickAll(pick(sheet, 'sheetData'), 'row')
			.flatMap((r) => pickAll(r, 'c'))
			.map((c) => pick(c, 'f')?.children[0])
			.filter(Boolean);
		expect(formulas).toEqual([
			'_xlfn.SEQUENCE(3)',
			'_xlfn.ANCHORARRAY(A1)',
			'_xlfn.SINGLE(A1)',
			'Items[[#This Row],[Qty]]*2',
		]);
		expect(pick(sheet, 'conditionalFormatting', 'cfRule', 'formula')?.children[0]).toBe(
			'_xlfn.SINGLE(A1)>1',
		);
		expect(pick(sheet, 'dataValidations', 'dataValidation', 'formula1')?.children[0]).toBe(
			'_xlfn.ANCHORARRAY(A1)',
		);
		const table = await partRoot(pkg, 'xl/tables/table1.xml');
		const cols = pickAll(pick(table, 'tableColumns'), 'tableColumn');
		expect(pick(cols[1], 'calculatedColumnFormula')?.children[0]).toBe('[[#This Row],[Qty]]*2');
		const name = pick(await partRoot(pkg, 'xl/workbook.xml'), 'definedNames', 'definedName');
		expect(name?.children[0]).toBe('_xlfn.ANCHORARRAY(Sheet1!$A$1)');
		expect(s.cells.get(cellKey(0, 1))?.formula).toBe('A1#');
	});
});
