import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import type { AstNode } from '../../ast';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg, LambdaValue } from '../../types';
import { INFORMATION_FUNCTIONS } from './index';

const ctx = new MockContext({
	A1: 1,
	A2: 'text',
	A3: true,
	A4: ERRORS.NA,
	A5: ERRORS.DIV0,
	A6: '',
	B1: 2,
	B2: '=A1',
	C7: 'extends the used range to row 7',
});
ctx.setFormula('B1', 'A1+1');
const f = (name: string, ...args: FunctionArg[]) =>
	call(fn(INFORMATION_FUNCTIONS, name), args, ctx);

describe('IS* predicates', () => {
	it.each([
		['ISBLANK', [false, false, false, false, false, false, true]],
		['ISERR', [false, false, false, false, true, false, false]],
		['ISERROR', [false, false, false, true, true, false, false]],
		['ISNA', [false, false, false, true, false, false, false]],
		['ISLOGICAL', [false, false, true, false, false, false, false]],
		['ISNUMBER', [true, false, false, false, false, false, false]],
		['ISTEXT', [false, true, false, false, false, true, false]],
		['ISNONTEXT', [true, false, true, true, true, false, true]],
	] as Array<[string, boolean[]]>)('%s over A1:A7', (name, expected) => {
		expect(f(name, ctx.ref('A1:A7'))).toEqual(arr(...expected.map((v) => [v])));
	});

	it('inspects direct values without propagating errors', () => {
		expect(f('ISERROR', ERRORS.REF)).toBe(true);
		expect(f('ISNA', ERRORS.NA)).toBe(true);
		expect(f('ISERR', ERRORS.NA)).toBe(false);
		expect(f('ISNUMBER', '1')).toBe(false);
		expect(f('ISTEXT', 1)).toBe(false);
		expect(f('ISBLANK', '')).toBe(false);
		expect(f('ISBLANK', null)).toBe(true);
		expect(f('ISLOGICAL', 'TRUE')).toBe(false);
		expect(f('ISNUMBER', arr([1, 'a', ERRORS.NA]))).toEqual(arr([true, false, false]));
	});

	it('sees the value of single-cell references', () => {
		expect(f('ISBLANK', ctx.ref('C9'))).toBe(true);
		expect(f('ISTEXT', ctx.ref('A2'))).toBe(true);
		expect(f('ISERR', ctx.ref('A5'))).toBe(true);
	});
});

describe('ISEVEN / ISODD', () => {
	it('tests the parity of the truncated number', () => {
		expect(f('ISEVEN', -1)).toBe(false);
		expect(f('ISEVEN', 2.5)).toBe(true);
		expect(f('ISEVEN', 5)).toBe(false);
		expect(f('ISEVEN', 0)).toBe(true);
		expect(f('ISEVEN', null)).toBe(true);
		expect(f('ISEVEN', '4')).toBe(true);
		expect(f('ISODD', -1)).toBe(true);
		expect(f('ISODD', 2.5)).toBe(false);
		expect(f('ISODD', 5)).toBe(true);
		expect(f('ISODD', -3.9)).toBe(true);
		expect(f('ISODD', 1e20)).toBe(false);
	});

	it('returns errors and rejects non-numbers', () => {
		expect(f('ISEVEN', ERRORS.NA)).toBe(ERRORS.NA);
		expect(f('ISODD', ctx.ref('A5'))).toBe(ERRORS.DIV0);
		expect(f('ISEVEN', true)).toBe(ERRORS.VALUE);
		expect(f('ISODD', 'abc')).toBe(ERRORS.VALUE);
		expect(f('ISEVEN', arr([1, 2]))).toEqual(arr([false, true]));
	});
});

describe('ISFORMULA', () => {
	it('checks references for formulas', () => {
		expect(f('ISFORMULA', ctx.ref('B1'))).toBe(true);
		expect(f('ISFORMULA', ctx.ref('B2'))).toBe(false);
		expect(f('ISFORMULA', ctx.ref('Z99'))).toBe(false);
		expect(f('ISFORMULA', ctx.ref('B1:B2'))).toEqual(arr([true], [false]));
		expect(fn(INFORMATION_FUNCTIONS, 'ISFORMULA').filePrefix).toBe('_xlfn.');
	});

	it('rejects non-references', () => {
		expect(f('ISFORMULA', 5)).toBe(ERRORS.VALUE);
		expect(f('ISFORMULA', ERRORS.REF)).toBe(ERRORS.REF);
		expect(f('ISFORMULA', { kind: 'refs', refs: [ctx.ref('A1'), ctx.ref('B1')] })).toBe(
			ERRORS.VALUE,
		);
	});
});

describe('TYPE', () => {
	it('classifies values', () => {
		expect(f('TYPE', 1)).toBe(1);
		expect(f('TYPE', 'Smith')).toBe(2);
		expect(f('TYPE', true)).toBe(4);
		expect(f('TYPE', ERRORS.DIV0)).toBe(16);
		expect(f('TYPE', arr([1, 2], [3, 4]))).toBe(64);
		expect(f('TYPE', arr([1]))).toBe(64);
		expect(f('TYPE', null)).toBe(1);
		expect(f('TYPE', { type: 'missing' } as AstNode)).toBe(1);
	});

	it('looks through single-cell references only', () => {
		expect(f('TYPE', ctx.ref('A2'))).toBe(2);
		expect(f('TYPE', ctx.ref('A4'))).toBe(16);
		expect(f('TYPE', ctx.ref('Z1'))).toBe(1);
		expect(f('TYPE', ctx.ref('A1:A3'))).toBe(64);
		const lambda: LambdaValue = {
			kind: 'lambda',
			params: [],
			body: { type: 'number', value: 1 },
			scope: new Map(),
		};
		expect(f('TYPE', lambda)).toBe(128);
	});
});

describe('ERROR.TYPE / N / NA', () => {
	it.each([
		[ERRORS.NULL, 1],
		[ERRORS.DIV0, 2],
		[ERRORS.VALUE, 3],
		[ERRORS.REF, 4],
		[ERRORS.NAME, 5],
		[ERRORS.NUM, 6],
		[ERRORS.NA, 7],
		[ERRORS.GETTING_DATA, 8],
		[ERRORS.SPILL, 9],
		[ERRORS.CALC, 14],
	])('ERROR.TYPE(%o) = %i', (error, code) => {
		expect(f('ERROR.TYPE', error)).toBe(code);
	});

	it('reports #N/A for non-errors and lifts over ranges', () => {
		expect(f('ERROR.TYPE', 1)).toBe(ERRORS.NA);
		expect(f('ERROR.TYPE', null)).toBe(ERRORS.NA);
		expect(f('ERROR.TYPE', ctx.ref('A4:A5'))).toEqual(arr([7], [2]));
	});

	it('converts values with N', () => {
		expect(f('N', 7)).toBe(7);
		expect(f('N', 'Even')).toBe(0);
		expect(f('N', true)).toBe(1);
		expect(f('N', false)).toBe(0);
		expect(f('N', null)).toBe(0);
		expect(f('N', 45292)).toBe(45292);
		expect(f('N', ERRORS.REF)).toBe(ERRORS.REF);
		expect(f('N', ctx.ref('A1:A3'))).toEqual(arr([1], [0], [1]));
	});

	it('returns #N/A from NA()', () => {
		expect(f('NA')).toBe(ERRORS.NA);
		expect(f('NA', 1)).toBe(ERRORS.VALUE);
	});
});

describe('definitions', () => {
	it('do not include ISREF and never propagate errors from IS* functions', () => {
		const names = INFORMATION_FUNCTIONS.map((d) => d.name);
		expect(names).not.toContain('ISREF');
		expect(new Set(names).size).toBe(names.length);
		for (const def of INFORMATION_FUNCTIONS.filter((d) => d.name.startsWith('IS'))) {
			expect([def.name, def.propagateErrors]).toEqual([def.name, false]);
		}
	});
});
