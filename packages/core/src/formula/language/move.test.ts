import { describe, expect, it } from 'vitest';

import { parseRange } from '../../address/range';
import type { RangeAddress } from '../../types/address';
import { adjustFormulaForMove, type FormulaMoveOptions } from './move';

const block = (text: string) => parseRange(text) as RangeAddress;

/** Move Sheet1!A1:B2 to D5 (row 4, col 3) unless told otherwise. */
function move(
	formula: string,
	opts: {
		formulaSheet?: string;
		from?: string;
		to?: string;
		range?: string;
		row?: number;
		col?: number;
	} & FormulaMoveOptions = {},
): string {
	return adjustFormulaForMove(
		formula,
		opts.formulaSheet ?? 'Sheet1',
		{ sheet: opts.from ?? 'Sheet1', range: block(opts.range ?? 'A1:B2') },
		{ sheet: opts.to ?? 'Sheet1', row: opts.row ?? 4, col: opts.col ?? 3 },
		{ formulaTargetSheet: opts.formulaTargetSheet },
	).formula;
}

describe('adjustFormulaForMove: same sheet', () => {
	it.each([
		['A1', 'D5'],
		['$A$1', '$D$5'],
		['A$1+$B2', 'D$5+$E6'],
		['SUM(A1:B2)', 'SUM(D5:E6)'],
		['SUM(B2:A1)', 'SUM(E6:D5)'],
		['SUM(A1:C3)', 'SUM(A1:C3)'],
		['C3', 'C3'],
		['A1#', 'D5#'],
		['Sheet1!B2', 'Sheet1!E6'],
		['Sheet2!A1', 'Sheet2!A1'],
		['SUM(A:A)', 'SUM(A:A)'],
		['"A1"&Rate', '"A1"&Rate'],
	])('%s -> %s', (formula, expected) => {
		expect(move(formula)).toBe(expected);
	});

	it('invalidates references to the overwritten destination cells', () => {
		expect(move('D5+1')).toBe('#REF!+1');
		expect(move('SUM(D5:E6)')).toBe('SUM(#REF!)');
		expect(move('Sheet1!E6')).toBe('Sheet1!#REF!');
		expect(move('SUM(D5:F7)')).toBe('SUM(D5:F7)');
		expect(move('D5', { formulaSheet: 'Sheet2' })).toBe('D5');
	});

	it('moves references inside an overlapping block before invalidating', () => {
		// A1:A5 moves down two rows to A3:A7.
		const opts = { range: 'A1:A5', row: 2, col: 0 };
		expect(move('A3', opts)).toBe('A5');
		expect(move('A6', opts)).toBe('#REF!');
		expect(move('SUM(A1:A5)', opts)).toBe('SUM(A3:A7)');
	});

	it('moves whole columns when whole columns are cut', () => {
		const opts = { range: 'B:C', row: 0, col: 4 };
		expect(move('SUM(B:B)+C1', opts)).toBe('SUM(E:E)+F1');
	});

	it('reports whether anything changed', () => {
		const source = { sheet: 'Sheet1', range: block('A1:B2') };
		const target = { sheet: 'Sheet1', row: 4, col: 3 };
		expect(adjustFormulaForMove('C9', 'Sheet1', source, target)).toEqual({
			formula: 'C9',
			changed: false,
		});
		expect(adjustFormulaForMove('A1', 'Sheet1', source, target)).toEqual({
			formula: 'D5',
			changed: true,
		});
	});
});

describe('adjustFormulaForMove: across sheets', () => {
	const toSheet2 = { to: 'Sheet2' };

	it('qualifies references that follow the block to another sheet', () => {
		expect(move('A1+C3', toSheet2)).toBe('Sheet2!D5+C3');
		expect(move('Sheet1!A1', { ...toSheet2, formulaSheet: 'Sheet3' })).toBe('Sheet2!D5');
		expect(move('Sheet1!A1', { ...toSheet2, formulaSheet: 'Sheet2' })).toBe('Sheet2!D5');
		expect(move('A1', { to: 'My Data' })).toBe("'My Data'!D5");
	});

	it('invalidates overwritten cells on the target sheet only', () => {
		expect(move('Sheet2!D5+D5', toSheet2)).toBe('Sheet2!#REF!+D5');
		expect(move('D5', { ...toSheet2, formulaSheet: 'Sheet2' })).toBe('#REF!');
	});

	it('keeps a moved formula pointing at the cells it used', () => {
		const opts = { ...toSheet2, formulaTargetSheet: 'Sheet2' };
		expect(move('A1+C3', opts)).toBe('D5+Sheet1!C3');
		expect(move('Sheet1!C3+Sheet3!A1', opts)).toBe('Sheet1!C3+Sheet3!A1');
	});
});
