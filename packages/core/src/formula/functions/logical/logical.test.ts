import { describe, expect, it } from 'vitest';

import { calculated, evalIn, valueAt } from '../../../calc/test-utils/workbook';
import { ERRORS } from '../../../values/errors';
import type { AstNode } from '../../ast';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { Scalar } from '../../types';
import { LOGICAL_FUNCTIONS } from './index';

const lit = (v: Scalar): AstNode => {
	if (typeof v === 'number') {
		return { type: 'number', value: v };
	}
	if (typeof v === 'string') {
		return { type: 'string', value: v };
	}
	if (typeof v === 'boolean') {
		return { type: 'boolean', value: v };
	}
	return { type: 'missing' };
};
/** MockContext only evaluates literals: evaluating this node throws, proving laziness. */
const bomb: AstNode = { type: 'cell', ref: { row: 0, col: 0, rowAbs: false, colAbs: false } };
const f = (name: string) => fn(LOGICAL_FUNCTIONS, name);

describe('lazy conditionals (MockContext)', () => {
	const ctx = new MockContext();

	it('IF evaluates only the taken branch', () => {
		expect(call(f('IF'), [lit(true), lit(1), bomb], ctx)).toBe(1);
		expect(call(f('IF'), [lit(0), bomb, lit('no')], ctx)).toBe('no');
		expect(call(f('IF'), [lit(false), lit(1)], ctx)).toBe(false);
		expect(call(f('IF'), [lit(true), { type: 'missing' }, lit(1)], ctx)).toBe(0);
		expect(call(f('IF'), [lit('TRUE'), lit(1), lit(2)], ctx)).toBe(1);
		expect(call(f('IF'), [lit('maybe'), lit(1), lit(2)], ctx)).toBe(ERRORS.VALUE);
		expect(call(f('IF'), [{ type: 'error', code: '#N/A' }, bomb, bomb], ctx)).toBe(ERRORS.NA);
	});

	it('IFERROR / IFNA evaluate the fallback only for (matching) errors', () => {
		expect(call(f('IFERROR'), [lit(5), bomb], ctx)).toBe(5);
		expect(call(f('IFERROR'), [{ type: 'error', code: '#DIV/0!' }, lit('x')], ctx)).toBe('x');
		expect(call(f('IFNA'), [{ type: 'error', code: '#DIV/0!' }, bomb], ctx)).toBe(ERRORS.DIV0);
		expect(call(f('IFNA'), [{ type: 'error', code: '#N/A' }, lit(0)], ctx)).toBe(0);
		expect(call(f('IFERROR'), [{ type: 'error', code: '#N/A' }, { type: 'missing' }], ctx)).toBe(0);
	});

	it('IFS, SWITCH and CHOOSE stop at the first match', () => {
		expect(call(f('IFS'), [lit(false), bomb, lit(true), lit('b'), bomb, bomb], ctx)).toBe('b');
		expect(call(f('IFS'), [lit(false), lit(1)], ctx)).toBe(ERRORS.NA);
		expect(call(f('IFS'), [lit(false), lit(1), lit(true)], ctx)).toBe(ERRORS.VALUE);
		expect(call(f('SWITCH'), [lit(2), lit(1), bomb, lit(2), lit('two'), bomb], ctx)).toBe('two');
		expect(call(f('SWITCH'), [lit('B'), lit('a'), lit(1), lit('b'), lit(2)], ctx)).toBe(2);
		expect(call(f('SWITCH'), [lit(9), lit(1), lit('a'), lit('default')], ctx)).toBe('default');
		expect(call(f('SWITCH'), [lit(9), lit(1), lit('a')], ctx)).toBe(ERRORS.NA);
		expect(call(f('SWITCH'), [lit(1), lit('1'), lit('text'), lit('none')], ctx)).toBe('none');
		expect(call(f('CHOOSE'), [lit(2), bomb, lit('second'), bomb], ctx)).toBe('second');
		expect(call(f('CHOOSE'), [lit(2.9), lit('a'), lit('b')], ctx)).toBe('b');
		expect(call(f('CHOOSE'), [lit(3), lit('a'), lit('b')], ctx)).toBe(ERRORS.VALUE);
		expect(call(f('CHOOSE'), [lit(0), lit('a')], ctx)).toBe(ERRORS.VALUE);
	});
});

describe('AND / OR / XOR / NOT', () => {
	const ctx = new MockContext({
		A1: true,
		A2: 1,
		A3: 'text',
		A4: null,
		B1: 0,
		B2: ERRORS.NA,
		C1: 'x',
	});

	it('aggregate logical values from references, ignoring text and blanks', () => {
		expect(call(f('AND'), [ctx.ref('A1:A4')], ctx)).toBe(true);
		expect(call(f('AND'), [ctx.ref('A1:A4'), ctx.ref('B1')], ctx)).toBe(false);
		expect(call(f('OR'), [ctx.ref('B1'), false], ctx)).toBe(false);
		expect(call(f('OR'), [ctx.ref('A1:B1')], ctx)).toBe(true);
		expect(call(f('XOR'), [true, true, true], ctx)).toBe(true);
		expect(call(f('XOR'), [ctx.ref('A1:A2')], ctx)).toBe(false);
	});

	it('require at least one logical value and propagate errors', () => {
		expect(call(f('AND'), [ctx.ref('C1')], ctx)).toBe(ERRORS.VALUE);
		expect(call(f('AND'), ['abc'], ctx)).toBe(ERRORS.VALUE);
		expect(call(f('AND'), ['true', 1], ctx)).toBe(true);
		expect(call(f('OR'), [ctx.ref('B1:B2')], ctx)).toBe(ERRORS.NA);
		expect(call(f('AND'), [arr([true, 'x', 1])], ctx)).toBe(true);
		expect(call(f('AND'), [true, { type: 'missing' }], ctx)).toBe(false);
	});

	it('NOT coerces and lifts', () => {
		expect(call(f('NOT'), [0], ctx)).toBe(true);
		expect(call(f('NOT'), [null], ctx)).toBe(true);
		expect(call(f('NOT'), ['FALSE'], ctx)).toBe(true);
		expect(call(f('NOT'), ['nope'], ctx)).toBe(ERRORS.VALUE);
		expect(call(f('NOT'), [arr([true, false])], ctx)).toEqual(arr([false, true]));
		expect(call(f('TRUE'), [], ctx)).toBe(true);
		expect(call(f('FALSE'), [], ctx)).toBe(false);
	});

	it('declare file prefixes for newer functions', () => {
		expect(f('XOR').filePrefix).toBe('_xlfn.');
		expect(f('IFS').filePrefix).toBe('_xlfn.');
		expect(f('SWITCH').filePrefix).toBe('_xlfn.');
		expect(f('IFNA').filePrefix).toBe('_xlfn.');
		expect(f('IF').filePrefix).toBeUndefined();
	});
});

describe('conditionals in a workbook', () => {
	const { engine } = calculated({ Sheet1: { A1: 1, A2: -2, A3: 3, B1: 10, B2: 20, B3: 30 } });
	const ev = (formula: string, at = 'Sheet1!E2', dynamic = true) =>
		evalIn(engine, formula, at, dynamic);

	it('lift over array conditions in dynamic mode', () => {
		expect(ev('IF(A1:A3>0,"pos","neg")')).toEqual([['pos'], ['neg'], ['pos']]);
		expect(ev('IF(A1:A3>0,B1:B3,-B1:B3)')).toEqual([[10], [-20], [30]]);
		expect(ev('IF({TRUE,FALSE},1)')).toEqual([[1, false]]);
		expect(ev('IFERROR(1/(A1:A3-1),"div")')).toEqual([['div'], [-1 / 3], [0.5]]);
		expect(ev('IFNA({1,#N/A},0)')).toEqual([[1, 0]]);
		expect(ev('IFS(A1:A3>2,"big",A1:A3>0,"small",TRUE,"neg")')).toEqual([
			['small'],
			['neg'],
			['big'],
		]);
		expect(ev('SWITCH(A1:A3,1,"one",3,"three","other")')).toEqual([['one'], ['other'], ['three']]);
		expect(ev('CHOOSE({1,2},"a","b")')).toEqual([['a', 'b']]);
	});

	it('implicitly intersect conditions in legacy mode', () => {
		expect(ev('IF(A1:A3>0,"pos","neg")', 'Sheet1!E2', false)).toBe('neg');
		expect(ev('IF(A1:A3>0,"pos","neg")', 'Sheet1!E3', false)).toBe('pos');
	});

	it('return references from the chosen branch', () => {
		expect(ev('SUM(IF(TRUE,B1:B3,A1:A3))')).toBe(60);
		expect(ev('ROWS(CHOOSE(2,A1,B1:B3))')).toBe(3);
		expect(ev('SUM(CHOOSE(2,A1:A3,B1:B3))')).toBe(60);
		expect(ev('ISREF(IF(TRUE,A1))')).toBe(true);
	});

	it('AND / OR over computed arrays', () => {
		expect(ev('AND(A1:A3>-5)')).toBe(true);
		expect(ev('OR(A1:A3>5)')).toBe(false);
		expect(ev('XOR(A1:A3>0)')).toBe(false);
	});
});

describe('conditionals: more cases', () => {
	const { engine, wb } = calculated({
		Sheet1: { A1: 5, A2: 'x', A3: '=1/0', B1: '=IF(A1>3,A2:A3,0)' },
	});
	const ev = (formula: string) => evalIn(engine, formula, 'Sheet1!E5');

	it('spills a reference chosen by IF and keeps errors inside', () => {
		expect(valueAt(wb, 'B1')).toBe('x');
		expect(valueAt(wb, 'B2')).toEqual(ERRORS.DIV0);
		expect(ev('IFERROR(A1:A3,"bad")')).toEqual([[5], ['x'], ['bad']]);
		expect(ev('ISERROR(A3)')).toBe(true);
	});

	it('coerces numeric and text conditions', () => {
		expect(ev('IF(0.5,"yes","no")')).toBe('yes');
		expect(ev('IF("false","yes","no")')).toBe('no');
		expect(ev('IF(A2,"yes","no")')).toEqual(ERRORS.VALUE);
		expect(ev('IF(Z99,"yes","no")')).toBe('no');
		expect(ev('NOT(A2)')).toEqual(ERRORS.VALUE);
	});

	it('propagates the first error in AND / OR arguments', () => {
		expect(ev('AND(TRUE,A3)')).toEqual(ERRORS.DIV0);
		expect(ev('OR(A1:A3)')).toEqual(ERRORS.DIV0);
		expect(ev('OR(FALSE,"TRUE")')).toBe(true);
	});

	it('SWITCH and CHOOSE pass references through', () => {
		expect(ev('ROWS(SWITCH(1,1,A1:A3,0))')).toBe(3);
		expect(ev('CHOOSE(1,A1:A2)')).toEqual([[5], ['x']]);
		expect(ev('SWITCH("X","x",1,2)')).toBe(1);
		expect(ev('SWITCH(NA(),1,2)')).toEqual(ERRORS.NA);
	});
});
