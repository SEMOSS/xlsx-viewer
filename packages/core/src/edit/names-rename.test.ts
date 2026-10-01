import { describe, expect, it } from 'vitest';

import { parseRange } from '../address/range';
import { createWorkbook } from '../sheet/workbook';
import { getCell } from '../sheet/worksheet';
import type { Workbook } from '../types/workbook';
import {
	DEFINED_NAME_DUPLICATE_MESSAGE,
	EditTransaction,
	renameDefinedNameInFormula,
	setCellInput,
	updateDefinedName,
} from './index';

function seed(wb: Workbook, entries: Record<string, string>, sheet = 0): void {
	const tx = new EditTransaction(wb, 'seed');
	for (const [ref, text] of Object.entries(entries)) {
		const r = parseRange(ref)!;
		setCellInput(tx, wb, sheet, r.top, r.left, text);
	}
	tx.commit();
}

function formulaAt(wb: Workbook, sheet: number, ref: string): string | undefined {
	const r = parseRange(ref)!;
	return getCell(wb.sheets[sheet], r.top, r.left)?.formula;
}

function names(wb: Workbook): string[] {
	return wb.definedNames.map((n) => `${n.name}@${n.localSheet ?? 'wb'}=${n.formula}`);
}

describe('renameDefinedNameInFormula', () => {
	it('renames name tokens only', () => {
		expect(renameDefinedNameInFormula('Rate*2+SUM(rate,A1)', 'Rate', 'Tax')).toBe(
			'Tax*2+SUM(Tax,A1)',
		);
		expect(renameDefinedNameInFormula('"Rate"&Rate&RateX', 'Rate', 'Tax')).toBe('"Rate"&Tax&RateX');
		expect(renameDefinedNameInFormula('RATE(10,-100,1000)+Rate', 'Rate', 'Tax')).toBe(
			'RATE(10,-100,1000)+Tax',
		);
		expect(renameDefinedNameInFormula('[1]!Rate+Rate', 'Rate', 'Tax')).toBe('[1]!Rate+Tax');
	});

	it('renames LAMBDA calls but not LET / LAMBDA parameters of the same spelling', () => {
		expect(renameDefinedNameInFormula('MyFn(2)*MYFN(3)', 'MyFn', 'Twice')).toBe(
			'Twice(2)*Twice(3)',
		);
		expect(renameDefinedNameInFormula('LET(Rate,5,Rate*2)+Rate', 'Rate', 'Tax')).toBe(
			'LET(Rate,5,Rate*2)+Tax',
		);
		expect(renameDefinedNameInFormula('LAMBDA(rate,rate+1)(Rate)', 'Rate', 'Tax')).toBe(
			'LAMBDA(rate,rate+1)(Tax)',
		);
	});

	it('lets the matcher decide by qualifier', () => {
		const onlySheet1 = (q: string | undefined) => q === 'Sheet1';
		expect(
			renameDefinedNameInFormula("Sheet1!Rate+'Sheet 2'!Rate+Rate", 'Rate', 'Tax', onlySheet1),
		).toBe("Sheet1!Tax+'Sheet 2'!Rate+Rate");
	});
});

function scopedWorkbook(): Workbook {
	const wb = createWorkbook({ sheetNames: ['Sheet1', 'Sheet2', 'Sheet3'] });
	wb.definedNames.push(
		{ name: 'Rate', formula: 'Sheet1!$A$1' },
		{ name: 'Rate', formula: 'Sheet2!$A$1', localSheet: 1 },
		{ name: 'Total', formula: 'Rate*2' },
		{ name: 'LocalTotal', formula: 'Rate*3', localSheet: 1 },
	);
	seed(wb, { B1: '=Rate*2', B2: '="Rate"&Rate', B3: '=Sheet2!Rate' }, 0);
	seed(wb, { B1: '=Rate+1', B2: '=Sheet1!Rate' }, 1);
	seed(wb, { B1: '=SUM(Rate,1)' }, 2);
	return wb;
}

describe('updateDefinedName', () => {
	it('renames a workbook name in formulas where it is visible and undoes exactly', () => {
		const wb = scopedWorkbook();
		const sheet3 = wb.sheets[2];
		sheet3.conditionalFormats.push({
			ranges: [parseRange('A1:A5')!],
			rules: [
				{ type: 'expression', priority: 1, formulas: ['A1>Rate'] },
				{
					type: 'colorScale',
					priority: 2,
					colorScale: {
						cfvos: [{ type: 'formula', value: 'Rate' }, { type: 'max' }],
						colors: [{ rgb: 'FF000000' }, { rgb: 'FFFFFFFF' }],
					},
				},
			],
		});
		sheet3.dataValidations.push({
			ranges: [parseRange('C1')!],
			type: 'decimal',
			operator: 'lessThan',
			formula1: 'Rate',
		});
		const cfBefore = sheet3.conditionalFormats[0];

		const tx = new EditTransaction(wb, 'Edit Name');
		const stored = updateDefinedName(
			tx,
			wb,
			{ name: 'rate', formula: '' },
			{ name: 'Tax', formula: '=Sheet1!$A$2' },
		);
		const entry = tx.commit();

		expect(stored).toEqual({ name: 'Tax', formula: 'Sheet1!$A$2' });
		expect(formulaAt(wb, 0, 'B1')).toBe('Tax*2');
		expect(formulaAt(wb, 0, 'B2')).toBe('"Rate"&Tax');
		// Sheet2 has its own Rate: plain and Sheet2-qualified uses keep pointing at it.
		expect(formulaAt(wb, 0, 'B3')).toBe('Sheet2!Rate');
		expect(formulaAt(wb, 1, 'B1')).toBe('Rate+1');
		expect(formulaAt(wb, 1, 'B2')).toBe('Sheet1!Tax');
		expect(formulaAt(wb, 2, 'B1')).toBe('SUM(Tax,1)');
		expect(sheet3.conditionalFormats[0].rules[0].formulas).toEqual(['A1>Tax']);
		expect(sheet3.conditionalFormats[0].rules[1].colorScale?.cfvos[0].value).toBe('Tax');
		expect(sheet3.dataValidations[0].formula1).toBe('Tax');
		expect(cfBefore.rules[0].formulas).toEqual(['A1>Rate']);
		// Existing entries keep their order; the renamed one moves to its sorted position.
		expect(names(wb)).toEqual([
			'Rate@1=Sheet2!$A$1',
			'Tax@wb=Sheet1!$A$2',
			'Total@wb=Tax*2',
			'LocalTotal@1=Rate*3',
		]);

		entry.undo();
		expect(formulaAt(wb, 0, 'B1')).toBe('Rate*2');
		expect(formulaAt(wb, 2, 'B1')).toBe('SUM(Rate,1)');
		expect(sheet3.conditionalFormats[0].rules[0].formulas).toEqual(['A1>Rate']);
		expect(sheet3.dataValidations[0].formula1).toBe('Rate');
		expect(names(wb)).toContain('Total@wb=Rate*2');
		expect(names(wb)).toContain('Rate@wb=Sheet1!$A$1');
		entry.redo();
		expect(formulaAt(wb, 1, 'B2')).toBe('Sheet1!Tax');
		expect(names(wb)).toContain('Total@wb=Tax*2');
	});

	it('renames uses in table column formulas and marks the table dirty', () => {
		const wb = scopedWorkbook();
		const table = {
			id: 1,
			name: 'Table1',
			displayName: 'Table1',
			ref: parseRange('D1:E4')!,
			headerRowCount: 1,
			totalsRowCount: 1,
			columns: [
				{ id: 1, name: 'Qty' },
				{
					id: 2,
					name: 'Due',
					calculatedColumnFormula: '[@Qty]*Rate',
					totalsRowFormula: 'SUM(Rate)',
				},
			],
			dirty: false,
		};
		wb.sheets[2].tables.push(table);
		const tx = new EditTransaction(wb, 'Edit Name');
		updateDefinedName(
			tx,
			wb,
			{ name: 'Rate', formula: '' },
			{ name: 'Tax', formula: 'Sheet1!$A$1' },
		);
		const entry = tx.commit();
		const renamed = wb.sheets[2].tables[0];
		expect(renamed.columns[1]).toMatchObject({
			calculatedColumnFormula: '[@Qty]*Tax',
			totalsRowFormula: 'SUM(Tax)',
		});
		expect(renamed.columns[0]).toBe(table.columns[0]);
		expect(renamed.dirty).toBe(true);
		expect(table.columns[1].calculatedColumnFormula).toBe('[@Qty]*Rate');
		entry.undo();
		expect(wb.sheets[2].tables[0]).toBe(table);
	});

	it('renames a sheet-level name only where it resolves', () => {
		const wb = scopedWorkbook();
		const tx = new EditTransaction(wb, 'Edit Name');
		updateDefinedName(
			tx,
			wb,
			{ name: 'Rate', formula: '', localSheet: 1 },
			{ name: 'Fee', formula: 'Sheet2!$A$1', localSheet: 1 },
		);
		tx.commit();
		expect(formulaAt(wb, 0, 'B1')).toBe('Rate*2');
		expect(formulaAt(wb, 0, 'B3')).toBe('Sheet2!Fee');
		expect(formulaAt(wb, 1, 'B1')).toBe('Fee+1');
		expect(formulaAt(wb, 1, 'B2')).toBe('Sheet1!Rate');
		expect(formulaAt(wb, 2, 'B1')).toBe('SUM(Rate,1)');
		expect(names(wb)).toContain('LocalTotal@1=Fee*3');
		expect(names(wb)).toContain('Total@wb=Rate*2');
	});

	it('changes only the formula and comment when the name stays', () => {
		const wb = scopedWorkbook();
		const tx = new EditTransaction(wb, 'Edit Name');
		updateDefinedName(
			tx,
			wb,
			{ name: 'Total', formula: '' },
			{ name: 'Total', formula: 'Rate*4', comment: 'Doubled' },
		);
		tx.commit();
		expect(wb.definedNames.find((n) => n.name === 'Total')).toEqual({
			name: 'Total',
			formula: 'Rate*4',
			comment: 'Doubled',
		});
		expect(formulaAt(wb, 0, 'B1')).toBe('Rate*2');
	});

	it('rewrites the casing of every use on a case-only rename', () => {
		const wb = scopedWorkbook();
		const tx = new EditTransaction(wb, 'Edit Name');
		updateDefinedName(
			tx,
			wb,
			{ name: 'Rate', formula: '' },
			{ name: 'RATE', formula: 'Sheet1!$A$1' },
		);
		tx.commit();
		expect(formulaAt(wb, 0, 'B1')).toBe('RATE*2');
	});

	it('keeps the built-in prefix and hidden flag', () => {
		const wb = createWorkbook();
		wb.definedNames.push({
			name: '_xlnm.Print_Area',
			formula: 'Sheet1!$A$1',
			localSheet: 0,
			builtin: true,
			hidden: true,
		});
		const tx = new EditTransaction(wb, 'Edit Name');
		updateDefinedName(tx, wb, wb.definedNames[0], {
			name: 'Print_Area',
			formula: 'Sheet1!$A$1:$D$9',
			localSheet: 0,
		});
		expect(wb.definedNames[0]).toEqual({
			name: '_xlnm.Print_Area',
			formula: 'Sheet1!$A$1:$D$9',
			localSheet: 0,
			builtin: true,
			hidden: true,
		});
	});

	it('rejects invalid or duplicate new names and ignores unknown targets', () => {
		const wb = scopedWorkbook();
		const tx = new EditTransaction(wb, 'Edit Name');
		expect(() =>
			updateDefinedName(tx, wb, { name: 'Total', formula: '' }, { name: 'rate', formula: '1' }),
		).toThrow(DEFINED_NAME_DUPLICATE_MESSAGE);
		expect(() =>
			updateDefinedName(tx, wb, { name: 'Total', formula: '' }, { name: 'R2C2', formula: '1' }),
		).toThrow(/not valid/);
		expect(
			updateDefinedName(tx, wb, { name: 'Missing', formula: '' }, { name: 'X', formula: '1' }),
		).toBeUndefined();
		expect(formulaAt(wb, 0, 'B1')).toBe('Rate*2');
		expect(wb.definedNames).toHaveLength(4);
	});
});
