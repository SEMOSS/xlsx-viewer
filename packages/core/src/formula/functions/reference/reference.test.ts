import { describe, expect, it } from 'vitest';

import { calculated, edit, evalIn, valueAt } from '../../../calc/test-utils/workbook';
import { ERRORS } from '../../../values/errors';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import { REFERENCE_FUNCTIONS } from './index';

const f = (name: string) => fn(REFERENCE_FUNCTIONS, name);

describe('reference functions (MockContext)', () => {
	const ctx = new MockContext({ A1: 1, B2: 2, C3: 3 }).at('D5');

	it('ROW / COLUMN of the formula cell and of references', () => {
		expect(call(f('ROW'), [], ctx)).toBe(5);
		expect(call(f('COLUMN'), [], ctx)).toBe(4);
		expect(call(f('ROW'), [ctx.ref('B7')], ctx)).toBe(7);
		expect(call(f('ROW'), [ctx.ref('A2:B3')], ctx)).toEqual(arr([2], [3]));
		expect(call(f('COLUMN'), [ctx.ref('B1:D9')], ctx)).toEqual(arr([2, 3, 4]));
		expect(call(f('ROW'), [arr([1])], ctx)).toBe(ERRORS.VALUE);
		expect(call(f('ROW'), [{ kind: 'refs', refs: [ctx.ref('A1'), ctx.ref('B1')] }], ctx)).toBe(
			ERRORS.REF,
		);
	});

	it('ROWS / COLUMNS / AREAS measure references and arrays', () => {
		expect(call(f('ROWS'), [ctx.ref('A1:C10')], ctx)).toBe(10);
		expect(call(f('COLUMNS'), [ctx.ref('A1:C10')], ctx)).toBe(3);
		expect(call(f('ROWS'), [arr([1, 2], [3, 4], [5, 6])], ctx)).toBe(3);
		expect(call(f('COLUMNS'), [arr([1, 2])], ctx)).toBe(2);
		expect(call(f('ROWS'), [7], ctx)).toBe(1);
		expect(call(f('ROWS'), [ERRORS.NA], ctx)).toBe(ERRORS.NA);
		expect(
			call(
				f('AREAS'),
				[{ kind: 'refs', refs: [ctx.ref('A1'), ctx.ref('B1'), ctx.ref('C1')] }],
				ctx,
			),
		).toBe(3);
		expect(call(f('AREAS'), [ctx.ref('A1:B2')], ctx)).toBe(1);
	});

	it('OFFSET shifts and resizes references', () => {
		expect(call(f('OFFSET'), [ctx.ref('A1'), 1, 1], ctx)).toEqual(ctx.ref('B2'));
		expect(call(f('OFFSET'), [ctx.ref('A1:B2'), 1, 0], ctx)).toEqual(ctx.ref('A2:B3'));
		expect(call(f('OFFSET'), [ctx.ref('A1'), 0, 0, 3, 2], ctx)).toEqual(ctx.ref('A1:B3'));
		expect(call(f('OFFSET'), [ctx.ref('C3'), 0, 0, -2, -2], ctx)).toEqual(ctx.ref('B2:C3'));
		expect(call(f('OFFSET'), [ctx.ref('A1'), -1, 0], ctx)).toBe(ERRORS.REF);
		expect(call(f('OFFSET'), [ctx.ref('A1'), 0, 0, 0], ctx)).toBe(ERRORS.REF);
		expect(call(f('OFFSET'), [5, 0, 0], ctx)).toBe(ERRORS.VALUE);
		expect(call(f('OFFSET'), [ctx.ref('A1'), 1.9, 0], ctx)).toEqual(ctx.ref('A2'));
		expect(f('OFFSET').volatile).toBe(true);
	});

	it('INDEX in reference and array form', () => {
		const grid = ctx.ref('A1:C3');
		expect(call(f('INDEX'), [grid, 2, 3], ctx)).toEqual(ctx.ref('C2'));
		expect(call(f('INDEX'), [grid, 2, 0], ctx)).toEqual(ctx.ref('A2:C2'));
		expect(call(f('INDEX'), [grid, 0, 2], ctx)).toEqual(ctx.ref('B1:B3'));
		expect(call(f('INDEX'), [grid, 0, 0], ctx)).toEqual(grid);
		expect(call(f('INDEX'), [grid, 2], ctx)).toEqual(ctx.ref('A2:C2'));
		expect(call(f('INDEX'), [ctx.ref('A1:E1'), 3], ctx)).toEqual(ctx.ref('C1'));
		expect(call(f('INDEX'), [ctx.ref('A1:A5'), 4], ctx)).toEqual(ctx.ref('A4'));
		expect(call(f('INDEX'), [grid, 4, 1], ctx)).toBe(ERRORS.REF);
		expect(call(f('INDEX'), [grid, -1, 1], ctx)).toBe(ERRORS.VALUE);
		const areas = { kind: 'refs' as const, refs: [ctx.ref('A1:B2'), ctx.ref('D1:E2')] };
		expect(call(f('INDEX'), [areas, 1, 2, 2], ctx)).toEqual(ctx.ref('E1'));
		expect(call(f('INDEX'), [areas, 1, 1, 3], ctx)).toBe(ERRORS.REF);
		const a = arr([1, 2, 3], [4, 5, 6]);
		expect(call(f('INDEX'), [a, 2, 3], ctx)).toBe(6);
		expect(call(f('INDEX'), [a, 2], ctx)).toEqual(arr([4, 5, 6]));
		expect(call(f('INDEX'), [a, 0, 2], ctx)).toEqual(arr([2], [5]));
		expect(call(f('INDEX'), [arr([1, 2, 3]), 2], ctx)).toBe(2);
		expect(call(f('INDEX'), [a, 1, 1, 2], ctx)).toBe(ERRORS.REF);
		expect(call(f('INDEX'), [9, 1, 1], ctx)).toBe(9);
	});

	it('ADDRESS in every style', () => {
		const address = f('ADDRESS');
		expect(call(address, [2, 3], ctx)).toBe('$C$2');
		expect(call(address, [2, 3, 2], ctx)).toBe('C$2');
		expect(call(address, [2, 3, 3], ctx)).toBe('$C2');
		expect(call(address, [2, 3, 4], ctx)).toBe('C2');
		expect(call(address, [2, 3, 1, false], ctx)).toBe('R2C3');
		expect(call(address, [2, 3, 2, false], ctx)).toBe('R2C[3]');
		expect(call(address, [2, 3, 3, false], ctx)).toBe('R[2]C3');
		expect(call(address, [2, 3, 4, false], ctx)).toBe('R[2]C[3]');
		expect(call(address, [1, 28, 1, true, 'Data'], ctx)).toBe('Data!$AB$1');
		expect(call(address, [1, 1, 1, true, 'My Sheet'], ctx)).toBe("'My Sheet'!$A$1");
		expect(call(address, [1, 16384], ctx)).toBe('$XFD$1');
		expect(call(address, [0, 1], ctx)).toBe(ERRORS.VALUE);
		expect(call(address, [1, 16385], ctx)).toBe(ERRORS.VALUE);
		expect(call(address, [1, 1, 5], ctx)).toBe(ERRORS.VALUE);
		expect(call(address, [arr([1, 2]), 1], ctx)).toEqual(arr(['$A$1', '$A$2']));
	});

	it('ISREF, SHEET, SHEETS, FORMULATEXT', () => {
		const multi = new MockContext({}, { Data: {}, Other: {} });
		multi.setFormula('A1', 'SUM(B1:B2)');
		expect(call(f('ISREF'), [multi.ref('A1')], multi)).toBe(true);
		expect(call(f('ISREF'), [5], multi)).toBe(false);
		expect(call(f('ISREF'), [ERRORS.REF], multi)).toBe(false);
		expect(call(f('SHEET'), [], multi)).toBe(1);
		expect(call(f('SHEET'), [multi.ref('A1', 2)], multi)).toBe(3);
		expect(call(f('SHEET'), ['data'], multi)).toBe(2);
		expect(call(f('SHEET'), ['nope'], multi)).toBe(ERRORS.NA);
		expect(call(f('SHEETS'), [], multi)).toBe(3);
		expect(
			call(
				f('SHEETS'),
				[{ kind: 'ref', sheet: 0, sheetEnd: 2, range: multi.ref('A1').range }],
				multi,
			),
		).toBe(3);
		expect(call(f('FORMULATEXT'), [multi.ref('A1')], multi)).toBe('=SUM(B1:B2)');
		expect(call(f('FORMULATEXT'), [multi.ref('B1')], multi)).toBe(ERRORS.NA);
		expect(call(f('FORMULATEXT'), [1], multi)).toBe(ERRORS.VALUE);
		expect(f('FORMULATEXT').filePrefix).toBe('_xlfn.');
	});
});

describe('reference functions in a workbook', () => {
	const { engine, wb } = calculated(
		{
			Sheet1: { A1: 1, A2: 2, A3: 3, B1: 'B', B2: 'Sheet2!A1', C1: '=SUM(A1:A3)' },
			Sheet2: { A1: 'other', A2: 5 },
		},
		{},
		(w) => w.definedNames.push({ name: 'Nums', formula: 'Sheet1!$A$1:$A$3' }),
	);
	const ev = (formula: string, at = 'Sheet1!E2') => evalIn(engine, formula, at);

	it('INDIRECT reads A1 and R1C1 text, sheet-qualified references and names', () => {
		expect(ev('INDIRECT("A2")')).toBe(2);
		expect(ev('INDIRECT("a1:a3")')).toEqual([[1], [2], [3]]);
		expect(ev('INDIRECT(B2)')).toBe('other');
		expect(ev('INDIRECT("Sheet2!A"&2)')).toBe(5);
		expect(ev('INDIRECT("R1C1",FALSE)')).toBe(1);
		expect(ev('INDIRECT("R[1]C[-4]",FALSE)')).toBe(3);
		expect(ev('SUM(INDIRECT("Nums"))')).toBe(6);
		expect(ev('INDIRECT("nonsense name")')).toEqual(ERRORS.REF);
		expect(ev('INDIRECT("")')).toEqual(ERRORS.REF);
		expect(ev('INDIRECT("1+1")')).toEqual(ERRORS.REF);
		expect(ev('INDIRECT("Nope!A1")')).toEqual(ERRORS.REF);
		expect(ev('ROWS(INDIRECT("A1:A10"))')).toBe(10);
	});

	it('OFFSET and INDEX build references for other functions and operators', () => {
		expect(ev('SUM(OFFSET(A1,0,0,3,1))')).toBe(6);
		expect(ev('OFFSET(A1,2,0)')).toBe(3);
		expect(ev('SUM(A1:INDEX(A1:A3,2))')).toBe(3);
		expect(ev('INDEX(A1:A3,0)*10')).toEqual([[10], [20], [30]]);
		expect(ev('INDEX(A1:A3,{1,3})')).toEqual([[1, 3]]);
		expect(ev('ROW(INDEX(A1:A3,2))')).toBe(2);
		expect(ev('CELL("address",INDEX(A1:A3,3))')).toBe('$A$3');
	});

	it('ADDRESS round-trips through INDIRECT', () => {
		expect(ev('INDIRECT(ADDRESS(3,1))')).toBe(3);
		expect(ev('INDIRECT(ADDRESS(2,1,1,TRUE,"Sheet2"))')).toBe(5);
	});

	it('FORMULATEXT follows formula edits', () => {
		edit(engine, 'Sheet1!D1', '=FORMULATEXT(C1)');
		expect(valueAt(wb, 'D1')).toBe('=SUM(A1:A3)');
		edit(engine, 'Sheet1!C1', '=MAX(A1:A3)');
		expect(valueAt(wb, 'D1')).toBe('=MAX(A1:A3)');
	});

	it('ROW / COLUMN return arrays for spills', () => {
		expect(ev('ROW(A2:A4)')).toEqual([[2], [3], [4]]);
		expect(ev('COLUMN(B1:C1)')).toEqual([[2, 3]]);
		expect(ev('SUM(ROW(A1:A3))')).toBe(6);
		expect(ev('COLUMNS(SEQUENCE(2,5))')).toBe(5);
		expect(ev('ROWS(A:A)')).toBe(1048576);
		expect(ev('ROW(A:A)')).toEqual([[1], [2], [3]]);
	});
});

describe('reference functions: more workbook cases', () => {
	const { engine } = calculated(
		{
			Data: { A1: 1, A2: 2, B1: 3, B2: 4, D1: 10, D2: 20, E1: 30, E2: 40 },
			Other: { C3: 'third' },
		},
		{},
		(w) =>
			w.definedNames.push(
				{ name: 'Local', formula: 'Data!$D$1', localSheet: 0 },
				{ name: 'Areas', formula: '(Data!$A$1:$B$2,Data!$D$1:$E$2)' },
			),
	);
	const ev = (formula: string, at = 'Data!H1') => evalIn(engine, formula, at);

	it('INDEX selects areas of multi-area references', () => {
		expect(ev('INDEX((A1:B2,D1:E2),2,2,2)')).toBe(40);
		expect(ev('INDEX(Areas,1,2,1)')).toBe(3);
		expect(ev('AREAS(Areas)')).toBe(2);
		expect(ev('SUM(INDEX(Areas,0,1,2))')).toBe(30);
		expect(ev('INDEX(Areas,1,1,3)')).toEqual(ERRORS.REF);
	});

	it('OFFSET and INDIRECT respect sheets and scoped names', () => {
		expect(ev('OFFSET(Other!A1,2,2)')).toBe('third');
		expect(ev('INDIRECT("Local")')).toBe(10);
		expect(ev('INDIRECT("Local")', 'Other!H1')).toEqual(ERRORS.REF);
		expect(ev('INDIRECT("\'Other\'!C3")')).toBe('third');
		expect(ev('SUM(OFFSET(A1:B2,0,3))')).toBe(100);
		expect(ev('OFFSET(A1,{0;1},0)')).toEqual([[1], [2]]);
	});

	it('SHEET / SHEETS / ADDRESS combine with other functions', () => {
		expect(ev('SHEET(Other!A1)')).toBe(2);
		expect(ev('SHEETS(Data:Other!A1)')).toBe(2);
		expect(ev('ADDRESS(ROW(B2),COLUMN(B2),4)')).toBe('B2');
		expect(ev('ADDRESS(1,1,1,TRUE,"O\'Neil")')).toBe("'O''Neil'!$A$1");
	});

	it('ROWS / COLUMNS of computed arrays and spills', () => {
		expect(ev('ROWS(FILTER(A1:A2,A1:A2>1))')).toBe(1);
		expect(ev('COLUMNS(HSTACK(A1:A2,B1:B2,D1:D2))')).toBe(3);
		expect(ev('ROWS({1;2;3;4})')).toBe(4);
	});
});
