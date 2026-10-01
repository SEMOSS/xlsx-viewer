import { describe, expect, it } from 'vitest';

import { calculated, evalIn, valuesAt } from '../../../calc/test-utils/workbook';
import { ERRORS } from '../../../values/errors';
import { fn } from '../../test-utils/mock-context';
import { LAMBDA_FUNCTIONS } from './index';

describe('LAMBDA helper functions', () => {
	const { engine } = calculated({
		Sheet1: { A1: 1, A2: 2, A3: 3, B1: 10, B2: 20, B3: 30, C2: 'x' },
	});
	const ev = (formula: string) => evalIn(engine, formula, 'Sheet1!H2');

	it('MAP applies a LAMBDA element-wise over one or more arrays', () => {
		expect(ev('MAP(A1:A3,LAMBDA(v,v*v))')).toEqual([[1], [4], [9]]);
		expect(ev('MAP(A1:A3,B1:B3,LAMBDA(a,b,a+b))')).toEqual([[11], [22], [33]]);
		expect(ev('MAP({1,2},{10;20},LAMBDA(a,b,a*b))')).toEqual([
			[10, 20],
			[20, 40],
		]);
		expect(ev('MAP({1,2,3},{1,2},LAMBDA(a,b,a+b))')).toEqual([[2, 4, ERRORS.NA]]);
		expect(ev('MAP(C1:C3,LAMBDA(c,ISBLANK(c)))')).toEqual([[true], [false], [true]]);
		expect(ev('MAP(A1:A3,LAMBDA(c,ROW(c)))')).toEqual([[1], [2], [3]]);
		expect(ev('MAP(A1:A2,LAMBDA(v,{1,2}))')).toEqual([[ERRORS.CALC], [ERRORS.CALC]]);
		expect(ev('MAP(A1:A3,LAMBDA(a,b,a))')).toEqual(ERRORS.VALUE);
		expect(ev('MAP(A1:A3,5)')).toEqual(ERRORS.VALUE);
	});

	it('REDUCE and SCAN accumulate', () => {
		expect(ev('REDUCE(0,A1:A3,LAMBDA(acc,v,acc+v))')).toBe(6);
		expect(ev('REDUCE(1,A1:B3,LAMBDA(acc,v,acc*v))')).toBe(36000);
		expect(ev('REDUCE(,A1:A3,LAMBDA(acc,v,acc+v))')).toBe(6);
		expect(ev('REDUCE("",{"a","b"},LAMBDA(acc,v,acc&v))')).toBe('ab');
		expect(ev('REDUCE(0,A1:A3,LAMBDA(acc,v,IF(v>acc,v,acc)))')).toBe(3);
		expect(ev('REDUCE(0,A1:A3,LAMBDA(acc,v,VSTACK(acc,v)))')).toEqual([[0], [1], [2], [3]]);
		expect(ev('SCAN(0,A1:A3,LAMBDA(acc,v,acc+v))')).toEqual([[1], [3], [6]]);
		expect(ev('SCAN(1,{1,2;3,4},LAMBDA(acc,v,acc*v))')).toEqual([
			[1, 2],
			[6, 24],
		]);
		expect(ev('SCAN("",{"a","b","c"},LAMBDA(acc,v,acc&v))')).toEqual([['a', 'ab', 'abc']]);
		expect(ev('REDUCE(0,A1:A3,LAMBDA(acc,acc))')).toEqual(ERRORS.VALUE);
	});

	it('MAKEARRAY builds arrays from row / column numbers', () => {
		expect(ev('MAKEARRAY(2,3,LAMBDA(r,c,r*10+c))')).toEqual([
			[11, 12, 13],
			[21, 22, 23],
		]);
		expect(ev('MAKEARRAY(1,1,LAMBDA(r,c,"x"))')).toEqual([['x']]);
		expect(ev('MAKEARRAY(0,2,LAMBDA(r,c,1))')).toEqual(ERRORS.VALUE);
		expect(ev('MAKEARRAY(2,2,LAMBDA(r,1))')).toEqual(ERRORS.VALUE);
	});

	it('BYROW and BYCOL reduce rows and columns', () => {
		expect(ev('BYROW(A1:B3,LAMBDA(r,SUM(r)))')).toEqual([[11], [22], [33]]);
		expect(ev('BYCOL(A1:B3,LAMBDA(c,MAX(c)))')).toEqual([[3, 30]]);
		expect(ev('BYROW({1,2;3,4},LAMBDA(r,COLUMNS(r)))')).toEqual([[2], [2]]);
		expect(ev('BYROW(A1:B2,LAMBDA(r,r))')).toEqual([[ERRORS.CALC], [ERRORS.CALC]]);
		expect(ev('BYCOL(A1:B3,COUNT)')).toEqual([[3, 3]]);
	});

	it('declare the _xlfn. file prefix', () => {
		for (const name of [
			'LET',
			'LAMBDA',
			'MAP',
			'REDUCE',
			'SCAN',
			'MAKEARRAY',
			'BYROW',
			'BYCOL',
			'ISOMITTED',
		]) {
			expect(fn(LAMBDA_FUNCTIONS, name).filePrefix).toBe('_xlfn.');
		}
	});
});

describe('LAMBDA helpers in cells', () => {
	it('spill their results and recalculate', () => {
		const { wb } = calculated({
			Sheet1: { A1: 2, A2: 4, B1: '=MAP(A1:A2,LAMBDA(v,v/2))', C1: '=LET(s,SUM(A1:A2),s*s)' },
		});
		expect(valuesAt(wb, 'B1:B2')).toEqual([[1], [2]]);
		expect(valuesAt(wb, 'C1:C1')).toEqual([[36]]);
	});
});

describe('LAMBDA helpers: more cases', () => {
	const { engine } = calculated({ Sheet1: { A1: 1, A2: 2, A3: 3 } }, {}, (w) =>
		w.definedNames.push({ name: 'Square', formula: 'LAMBDA(x,x*x)' }),
	);
	const ev = (formula: string) => evalIn(engine, formula, 'Sheet1!H2');

	it('accepts named LAMBDAs and LAMBDAs bound with LET', () => {
		expect(ev('MAP(A1:A3,Square)')).toEqual([[1], [4], [9]]);
		expect(ev('LET(f,LAMBDA(v,v+1),MAP(A1:A3,f))')).toEqual([[2], [3], [4]]);
		expect(ev('BYCOL(A1:A3,LAMBDA(c,SUM(MAP(c,Square))))')).toEqual([[14]]);
	});

	it('captures LET bindings in closures', () => {
		expect(ev('LET(k,10,MAP(A1:A3,LAMBDA(v,v*k)))')).toEqual([[10], [20], [30]]);
		expect(ev('LET(add,LAMBDA(n,LAMBDA(v,v+n)),MAP(A1:A3,add(100)))')).toEqual([
			[101],
			[102],
			[103],
		]);
	});

	it('propagates errors from the LAMBDA per element', () => {
		expect(ev('MAP({1,0},LAMBDA(v,1/v))')).toEqual([[1, ERRORS.DIV0]]);
		expect(ev('SCAN(0,{1,"a",2},LAMBDA(acc,v,acc+v))')).toEqual([[1, ERRORS.VALUE, ERRORS.VALUE]]);
		expect(ev('MAKEARRAY(2,1,LAMBDA(r,c,IF(r=1,NA(),r)))')).toEqual([[ERRORS.NA], [2]]);
	});
});
