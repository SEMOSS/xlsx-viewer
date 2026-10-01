import { describe, expect, it } from 'vitest';

import { translateFormula } from '../translate';

describe('translateFormula', () => {
	it.each([
		['A1+$B$2+C$3+$D4', 1, 1, 'B2+$B$2+D$3+$D5'],
		['SUM(A1:A3)&"A1"', 2, 0, 'SUM(A3:A5)&"A1"'],
		["'Sheet A1'!B2", 1, 0, "'Sheet A1'!B3"],
		['SUM(A1:B2)', 1, 1, 'SUM(B2:C3)'],
		['$A1:B$2', 3, 3, '$A4:E$2'],
		['$A$1:A1', 4, 0, '$A$1:A5'],
		['SUM(A:B)', 5, 1, 'SUM(B:C)'],
		['SUM($A:B)', 0, 2, 'SUM($A:D)'],
		['SUM(1:2)', 1, 5, 'SUM(2:3)'],
		['SUM($1:2)', 3, 0, 'SUM($1:5)'],
		['Sheet2!A1*2', 0, 1, 'Sheet2!B1*2'],
		['Sheet1:Sheet3!A1', 1, 0, 'Sheet1:Sheet3!A2'],
		['[1]Sheet1!A1', 0, 1, '[1]Sheet1!B1'],
		['SUM(A1#)', 1, 1, 'SUM(B2#)'],
		['B2:A1', 1, 0, 'B3:A2'],
		['SUM( A1 ,\n B1 )', 1, 0, 'SUM( A2 ,\n B2 )'],
		['=A1+1', 1, 0, '=A2+1'],
	])('%s by (%i, %i)', (formula, dr, dc, expected) => {
		expect(translateFormula(formula, dr, dc)).toBe(expected);
	});

	it('leaves strings, names, functions and table references alone', () => {
		expect(translateFormula('"A1"&Rate*LOG10(A1)+Table1[A1]+R1C1', 1, 0)).toBe(
			'"A1"&Rate*LOG10(A2)+Table1[A1]+R1C1',
		);
		expect(translateFormula('Sheet1!MyName+INDIRECT("B2")', 5, 5)).toBe(
			'Sheet1!MyName+INDIRECT("B2")',
		);
	});

	it('turns references moved off the grid into #REF!', () => {
		expect(translateFormula('A1', -1, 0)).toBe('#REF!');
		expect(translateFormula('A1+B2', 0, -1)).toBe('#REF!+A2');
		expect(translateFormula('SUM(A1:B5)', -1, 0)).toBe('SUM(#REF!)');
		expect(translateFormula('Sheet2!A1', -1, 0)).toBe('Sheet2!#REF!');
		expect(translateFormula("'My Sheet'!A1:B2", 0, -1)).toBe("'My Sheet'!#REF!");
		expect(translateFormula('XFD1', 0, 1)).toBe('#REF!');
		expect(translateFormula('A1048576', 1, 0)).toBe('#REF!');
		expect(translateFormula('A:A', 0, -1)).toBe('#REF!');
		expect(translateFormula('SUM(A1#)', -1, 0)).toBe('SUM(#REF!)');
	});

	it('keeps absolute references that would leave the grid', () => {
		expect(translateFormula('$A$1+$A1', -1, 0)).toBe('$A$1+#REF!');
		expect(translateFormula('$A:$A', 0, -1)).toBe('$A:$A');
	});

	it('returns the input when nothing moves', () => {
		const formula = 'SUM( $A$1 , "x" )';
		expect(translateFormula(formula, 0, 0)).toBe(formula);
		expect(translateFormula(formula, 3, 3)).toBe(formula);
		expect(translateFormula('', 1, 1)).toBe('');
	});

	it('copes with formulas that do not parse', () => {
		expect(translateFormula('SUM(A1', 1, 0)).toBe('SUM(A2');
		expect(translateFormula('"open A1', 1, 0)).toBe('"open A1');
	});
});
