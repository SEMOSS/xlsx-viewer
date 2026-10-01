import { describe, expect, it } from 'vitest';

import {
	formulaReferencesSheet,
	invalidateSheetInFormula,
	renameSheetInFormula,
} from './sheet-refs';

describe('renameSheetInFormula', () => {
	it.each([
		['Sheet1!A1+1', 'Sheet1', 'Data', 'Data!A1+1'],
		["Sheet1!A1+'Sheet1'!B2+sheet1!C3", 'Sheet1', 'Data', 'Data!A1+Data!B2+Data!C3'],
		['Sheet1!A1', 'Sheet1', 'My Data', "'My Data'!A1"],
		['Sheet1!A1', 'Sheet1', 'Q1', "'Q1'!A1"],
		['Sheet1!A1', 'Sheet1', "Bob's", "'Bob''s'!A1"],
		["'My Sheet'!A1:B2*2", 'My Sheet', 'Totals', 'Totals!A1:B2*2'],
		['SUM(Sheet1:Sheet3!A1)', 'Sheet3', 'Data', 'SUM(Sheet1:Data!A1)'],
		['SUM(Sheet1:Sheet3!A1)', 'Sheet1', 'Year 1', "SUM('Year 1:Sheet3'!A1)"],
		['Sheet1!Rate*Sheet1!#REF!', 'Sheet1', 'Data', 'Data!Rate*Data!#REF!'],
		['Sheet1!B2#', 'Sheet1', 'Data', 'Data!B2#'],
		['Sheet10!A1+Sheet1x!A1', 'Sheet1', 'Data', 'Sheet10!A1+Sheet1x!A1'],
		['INDIRECT("Sheet1!A1")&Sheet1', 'Sheet1', 'Data', 'INDIRECT("Sheet1!A1")&Sheet1'],
		['[1]Sheet1!A1', 'Sheet1', 'Data', '[1]Sheet1!A1'],
		['A1+B1', 'Sheet1', 'Data', 'A1+B1'],
	])('%s: %s -> %s', (formula, from, to, expected) => {
		expect(renameSheetInFormula(formula, from, to)).toBe(expected);
	});

	it('keeps formatting outside the renamed prefixes', () => {
		expect(renameSheetInFormula('SUM( Sheet1!A1 ,\n Sheet2!B1 )', 'Sheet2', 'Z')).toBe(
			'SUM( Sheet1!A1 ,\n Z!B1 )',
		);
	});
});

describe('formulaReferencesSheet', () => {
	it.each([
		['Sheet1!A1', 'Sheet1', true],
		['sheet1!A1', 'SHEET1', true],
		["'My Sheet'!A1", 'My Sheet', true],
		['Sheet1:Sheet3!A1', 'Sheet3', true],
		['Sheet1!Rate', 'Sheet1', true],
		['A1+B1', 'Sheet1', false],
		['"Sheet1!A1"', 'Sheet1', false],
		['[1]Sheet1!A1', 'Sheet1', false],
		['Sheet10!A1', 'Sheet1', false],
	])('%s mentions %s: %s', (formula, name, expected) => {
		expect(formulaReferencesSheet(formula, name)).toBe(expected);
	});
});

describe('invalidateSheetInFormula', () => {
	it.each([
		['Gone!A1+1', '#REF!+1'],
		['SUM(Gone!A1:B5,A1)', 'SUM(#REF!,A1)'],
		["'gone'!A1*Other!A1", '#REF!*Other!A1'],
		['Gone!Rate+Gone!#REF!', '#REF!+#REF!'],
		['SUM(Gone!B2#)', 'SUM(#REF!)'],
		['Gone:Gone!A1', '#REF!'],
		['SUM(Gone:Other!A1)', 'SUM(Gone:Other!A1)'],
		['"Gone!A1"', '"Gone!A1"'],
	])('%s -> %s', (formula, expected) => {
		expect(invalidateSheetInFormula(formula, 'Gone')).toBe(expected);
	});
});
