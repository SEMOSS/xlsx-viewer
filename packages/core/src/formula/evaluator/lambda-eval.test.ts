import { describe, expect, it } from 'vitest';

import { calculated, evalIn } from '../../calc/test-utils/workbook';
import { ERRORS } from '../../values/errors';

describe('LET / LAMBDA evaluation', () => {
	const { engine } = calculated({ Sheet1: { A1: 1, A2: 2, A3: 3, B1: 'b' } }, {}, (wb) => {
		wb.definedNames.push(
			{ name: 'x', formula: '1000' },
			{ name: 'AddN', formula: 'LAMBDA(a,b,a+b)' },
		);
	});
	const ev = (f: string) => evalIn(engine, f, 'Sheet1!E2');

	it('binds LET names in order, shadowing defined names', () => {
		expect(ev('LET(a,2,b,a*3,a+b)')).toBe(8);
		expect(ev('LET(x,5,x*2)')).toBe(10);
		expect(ev('x')).toBe(1000);
		expect(ev('LET(r,A1:A3,SUM(r)*ROWS(r))')).toBe(18);
		expect(ev('LET(r,A1:A3,r*10)')).toEqual([[10], [20], [30]]);
		expect(ev('LET(a,1,LET(a,2,a)+a)')).toBe(3);
		expect(ev('_xlfn.LET(_xlpm.n,4,_xlpm.n^2)')).toBe(16);
	});

	it('rejects malformed LET', () => {
		expect(ev('LET(1,2,3)')).toEqual(ERRORS.VALUE);
		expect(ev('LET(a,1,b,2)')).toEqual(ERRORS.VALUE);
	});

	it('calls LAMBDA values directly and through LET', () => {
		expect(ev('LAMBDA(x,x+1)(2)')).toBe(3);
		expect(ev('LAMBDA(1+1)()')).toBe(2);
		expect(ev('LET(f,LAMBDA(v,v*v),f(4)+f(2))')).toBe(20);
		expect(ev('LET(k,10,f,LAMBDA(v,v+k),k+f(1))')).toBe(21);
		expect(ev('LAMBDA(a,LAMBDA(b,a+b))(1)(2)')).toBe(3);
		expect(ev('AddN(2,3)')).toBe(5);
		expect(ev('LET(AddN,LAMBDA(a,b,a*b),AddN(2,3))')).toBe(6);
	});

	it('reports wrong arities and non-callables', () => {
		expect(ev('LAMBDA(x,x)(1,2)')).toEqual(ERRORS.VALUE);
		expect(ev('LAMBDA(x,y,x+y)(1)')).toEqual(ERRORS.VALUE);
		expect(ev('LET(v,5,v(1))')).toEqual(ERRORS.VALUE);
		expect(ev('LAMBDA(x,x,x)')).toEqual(ERRORS.VALUE);
		expect(ev('LAMBDA(x,x+1)')).toBe('lambda');
	});

	it('supports optional parameters and ISOMITTED', () => {
		expect(ev('LAMBDA(a,[b],IF(ISOMITTED(b),a,a+b))(1)')).toBe(1);
		expect(ev('LAMBDA(a,[b],IF(ISOMITTED(b),a,a+b))(1,2)')).toBe(3);
		expect(ev('LAMBDA(a,b,IF(ISOMITTED(b),"none",b))(1,)')).toBe('none');
		expect(ev('LAMBDA(a,[b],b)(1)')).toBe(null);
		expect(ev('ISOMITTED(A1)')).toBe(false);
	});

	it('passes eta-reduced functions to LAMBDA helpers', () => {
		expect(ev('MAP({-1,2},ABS)')).toEqual([[1, 2]]);
		expect(ev('BYROW({1,2;3,4},SUM)')).toEqual([[3], [7]]);
		expect(ev('REDUCE(0,{1,2,3},SUM)')).toBe(6);
		expect(ev('LET(f,ABS,f(-3))')).toBe(3);
		expect(ev('SUM')).toEqual(ERRORS.NAME);
	});

	it('limits recursion depth', () => {
		const deep = calculated({ Sheet1: {} }, {}, (wb) => {
			wb.definedNames.push({ name: 'Down', formula: 'LAMBDA(n,IF(n=0,0,Down(n-1)))' });
		});
		expect(evalIn(deep.engine, 'Down(50)')).toBe(0);
		expect(evalIn(deep.engine, 'Down(100000)')).toEqual(ERRORS.NUM);
	});
});
