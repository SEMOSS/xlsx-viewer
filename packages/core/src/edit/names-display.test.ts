import { describe, expect, it } from 'vitest';

import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../address/constants';
import { parseRange } from '../address/range';
import { createWorkbook } from '../sheet/workbook';
import type { Workbook } from '../types/workbook';
import { ERRORS } from '../values/errors';
import {
	definedNameDisplayName,
	definedNameDisplayValue,
	EditTransaction,
	parseRefersTo,
	qualifyFormulaReferences,
	REFERS_TO_ERROR_MESSAGE,
	sanitizeDefinedName,
	selectionReference,
	setCellInput,
	suggestDefinedName,
} from './index';

const r = (a1: string) => parseRange(a1)!;

function seed(wb: Workbook, entries: Record<string, string>, sheet = 0): void {
	const tx = new EditTransaction(wb, 'seed');
	for (const [ref, text] of Object.entries(entries)) {
		setCellInput(tx, wb, sheet, r(ref).top, r(ref).left, text);
	}
	tx.commit();
}

describe('selectionReference', () => {
	it('builds absolute, sheet-qualified references', () => {
		expect(selectionReference('Sheet1', [r('A1:B5')])).toBe('Sheet1!$A$1:$B$5');
		expect(selectionReference('My Sheet', [r('C3')])).toBe("'My Sheet'!$C$3");
		expect(selectionReference("Bob's", [r('A1')])).toBe("'Bob''s'!$A$1");
		expect(selectionReference('Sheet1', [r('A1:B2'), r('D4')])).toBe(
			'Sheet1!$A$1:$B$2,Sheet1!$D$4',
		);
	});

	it('writes whole columns, rows and the whole sheet like Excel', () => {
		expect(selectionReference('Sheet1', [r('A:A')])).toBe('Sheet1!$A:$A');
		expect(selectionReference('Sheet1', [r('B:D')])).toBe('Sheet1!$B:$D');
		expect(selectionReference('Sheet1', [r('2:2')])).toBe('Sheet1!$2:$2');
		const all = { top: 0, left: 0, bottom: MAX_ROW_INDEX, right: MAX_COL_INDEX };
		expect(selectionReference('Sheet1', [all])).toBe('Sheet1!$1:$1048576');
	});
});

describe('sanitizeDefinedName / definedNameDisplayName', () => {
	it('turns labels into valid names', () => {
		expect(sanitizeDefinedName('Unit Price')).toBe('Unit_Price');
		expect(sanitizeDefinedName('  Sales ($) ')).toBe('Sales____');
		expect(sanitizeDefinedName('2024 Total')).toBe('_2024_Total');
		expect(sanitizeDefinedName('.5')).toBe('_.5');
		expect(sanitizeDefinedName('Q1')).toBe('Q1_');
		expect(sanitizeDefinedName('r')).toBe('r_');
		expect(sanitizeDefinedName('True')).toBe('True_');
		expect(sanitizeDefinedName('Größe')).toBe('Größe');
		expect(sanitizeDefinedName('   ')).toBe('');
		expect(sanitizeDefinedName('x'.repeat(300))).toHaveLength(255);
	});

	it('hides the built-in prefix', () => {
		expect(definedNameDisplayName({ name: '_xlnm.Print_Area' })).toBe('Print_Area');
		expect(definedNameDisplayName({ name: 'Rate' })).toBe('Rate');
	});
});

describe('suggestDefinedName', () => {
	const wb = createWorkbook();
	seed(wb, {
		A1: 'Region',
		B1: 'Unit Price',
		A2: 'North',
		B2: '12',
		D5: 'Total due',
		E5: '99',
		G1: '7',
	});

	it('uses the top-left label of the selection', () => {
		expect(suggestDefinedName(wb, 0, r('A1:A9'))).toBe('Region');
		expect(suggestDefinedName(wb, 0, r('A2:B3'))).toBe('North');
	});

	it('falls back to the label above, then to the left', () => {
		expect(suggestDefinedName(wb, 0, r('B2:B9'))).toBe('Unit_Price');
		expect(suggestDefinedName(wb, 0, r('E5'))).toBe('Total_due');
	});

	it('returns an empty string without a text label', () => {
		expect(suggestDefinedName(wb, 0, r('G1'))).toBe('');
		expect(suggestDefinedName(wb, 0, r('K10'))).toBe('');
		expect(suggestDefinedName(wb, 3, r('A1'))).toBe('');
	});
});

describe('definedNameDisplayValue', () => {
	function book(): Workbook {
		const wb = createWorkbook({ sheetNames: ['Data', 'My Sheet'] });
		seed(wb, { A1: 'Apple', B1: '1', A2: 'Banana', B2: '2', C1: '7%', C2: '1/2/2024' });
		seed(wb, { A1: 'x"y' }, 1);
		return wb;
	}

	it('shows the displayed value of a one-cell reference', () => {
		const wb = book();
		expect(definedNameDisplayValue(wb, { name: 'R', formula: 'Data!$C$1' })).toBe('7%');
		expect(definedNameDisplayValue(wb, { name: 'R', formula: '=Data!$C$2' })).toBe('1/2/2024');
		expect(definedNameDisplayValue(wb, { name: 'R', formula: "'My Sheet'!$A$1" })).toBe('x"y');
		expect(definedNameDisplayValue(wb, { name: 'R', formula: 'Data!$D$9' })).toBe('');
	});

	it('resolves unqualified references on the name sheet', () => {
		const wb = book();
		expect(definedNameDisplayValue(wb, { name: 'R', formula: '$A$1', localSheet: 1 })).toBe('x"y');
		expect(definedNameDisplayValue(wb, { name: 'R', formula: '$A$1' })).toBe('Apple');
		const ctx = { sheet: 1, row: 0, col: 0 };
		expect(definedNameDisplayValue(wb, { name: 'R', formula: '$A$1' }, undefined, ctx)).toBe('x"y');
	});

	it("shows multi-cell references as Excel's quoted array", () => {
		const wb = book();
		expect(definedNameDisplayValue(wb, { name: 'R', formula: 'Data!$A$1:$B$2' })).toBe(
			'{"Apple","1";"Banana","2"}',
		);
		expect(definedNameDisplayValue(wb, { name: 'R', formula: 'Data!$A$2:$A$3' })).toBe(
			'{"Banana";""}',
		);
		expect(definedNameDisplayValue(wb, { name: 'R', formula: "'My Sheet'!$A$1:$B$1" })).toBe(
			'{"x""y",""}',
		);
	});

	it('cuts long ranges off', () => {
		const wb = book();
		const text = definedNameDisplayValue(wb, { name: 'R', formula: 'Data!$A:$A' });
		expect(text.startsWith('{"Apple";"Banana";"";')).toBe(true);
		expect(text.endsWith('...}')).toBe(true);
		expect(text.split(';')).toHaveLength(50);
	});

	it('reports broken references as #REF!', () => {
		const wb = book();
		expect(definedNameDisplayValue(wb, { name: 'R', formula: 'Gone!$A$1' })).toBe('#REF!');
		expect(definedNameDisplayValue(wb, { name: 'R', formula: 'Data!#REF!' })).toBe('#REF!');
		expect(definedNameDisplayValue(wb, { name: 'R', formula: 'SUM(#REF!)' })).toBe('#REF!');
	});

	it('shows constants and evaluated formulas', () => {
		const wb = book();
		expect(definedNameDisplayValue(wb, { name: 'R', formula: '0.07' })).toBe('0.07');
		expect(definedNameDisplayValue(wb, { name: 'R', formula: '"abc"' })).toBe('abc');
		expect(definedNameDisplayValue(wb, { name: 'R', formula: 'SUM(Data!B1:B2)' })).toBe('');
		const calls: unknown[][] = [];
		const evaluate = (...args: [string, number, number, number]) => {
			calls.push(args);
			return 3;
		};
		expect(definedNameDisplayValue(wb, { name: 'R', formula: 'SUM(Data!B1:B2)' }, evaluate)).toBe(
			'3',
		);
		expect(calls).toEqual([['SUM(Data!B1:B2)', 0, 0, 0]]);
		expect(definedNameDisplayValue(wb, { name: 'R', formula: 'X' }, () => ERRORS.NAME)).toBe(
			'#NAME?',
		);
		expect(definedNameDisplayValue(wb, { name: 'R', formula: 'X' }, () => true)).toBe('TRUE');
		const ref = { kind: 'ref' as const, sheet: 0, range: r('A1:A2') };
		expect(definedNameDisplayValue(wb, { name: 'R', formula: 'X' }, () => ref)).toBe(
			'{"Apple";"Banana"}',
		);
		const arr = { kind: 'array' as const, rows: 1, cols: 3, values: [[1, 'b', true]] };
		expect(definedNameDisplayValue(wb, { name: 'R', formula: 'X' }, () => arr)).toBe(
			'{1,"b",TRUE}',
		);
	});
});

describe('parseRefersTo', () => {
	it('accepts formulas and qualifies plain references with the sheet', () => {
		expect(parseRefersTo('=Sheet1!$A$1:$B$2')).toEqual({ formula: 'Sheet1!$A$1:$B$2' });
		expect(parseRefersTo(' =$A$1 ', 'Sheet1')).toEqual({ formula: 'Sheet1!$A$1' });
		expect(parseRefersTo('=SUM($A$1:$A$5)*Rate', 'My Sheet')).toEqual({
			formula: "SUM('My Sheet'!$A$1:$A$5)*Rate",
		});
		expect(parseRefersTo('=Other!$A$1+$B$2', 'Sheet1')).toEqual({
			formula: 'Other!$A$1+Sheet1!$B$2',
		});
	});

	it('turns entries without = into constants', () => {
		expect(parseRefersTo('5')).toEqual({ formula: '5' });
		expect(parseRefersTo('-1.5')).toEqual({ formula: '-1.5' });
		expect(parseRefersTo('true')).toEqual({ formula: 'TRUE' });
		expect(parseRefersTo('abc')).toEqual({ formula: '"abc"' });
		expect(parseRefersTo('say "hi"')).toEqual({ formula: '"say ""hi"""' });
	});

	it("reports Excel's formula error", () => {
		expect(parseRefersTo('')).toEqual({ error: REFERS_TO_ERROR_MESSAGE });
		expect(parseRefersTo('=')).toEqual({ error: REFERS_TO_ERROR_MESSAGE });
		expect(parseRefersTo('=1+')).toEqual({ error: REFERS_TO_ERROR_MESSAGE });
		expect(parseRefersTo('=SUM(1')).toEqual({ error: REFERS_TO_ERROR_MESSAGE });
		expect(REFERS_TO_ERROR_MESSAGE).toBe("There's a problem with this formula.");
	});

	it('qualifies references without touching names or strings', () => {
		expect(qualifyFormulaReferences('A1+"B2"+Rate+Data!C3', 'Sheet 1')).toBe(
			'\'Sheet 1\'!A1+"B2"+Rate+Data!C3',
		);
	});
});
