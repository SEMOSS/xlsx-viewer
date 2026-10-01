import { describe, expect, it } from 'vitest';

import { calculated, evalIn } from '../../calc/test-utils/workbook';
import { ERRORS } from '../../values/errors';

/**
 * Evaluator semantics through a real workbook: Sheet1 holds A1:A3 = 1, 2, 3,
 * B1:B3 = "x", "Y", TRUE, C1 = "5" (text), D1 an empty (formatted) cell.
 */
function setup() {
	return calculated({
		Sheet1: { A1: 1, A2: 2, A3: 3, B1: 'x', B2: 'Y', B3: true, C1: '5', D1: null },
		Sheet2: { A1: 10, A2: 20, B5: 'two' },
		'My Sheet': { A1: 100 },
	});
}

describe('operators and precedence', () => {
	const { engine } = setup();
	const ev = (f: string) => evalIn(engine, f, 'Sheet1!E2');

	it.each([
		['1+2*3', 7],
		['(1+2)*3', 9],
		['-2^2', 4],
		['-(2^2)', -4],
		['2^3^2', 64],
		['10-2-3', 5],
		['2*3%', 0.06],
		['1&2+3', '15'],
		['1+2=3', true],
		['"a"&"b"="AB"', true],
		['+A1', 1],
		['--"3"', 3],
	])('%s', (formula, expected) => {
		expect(ev(formula)).toEqual(expected);
	});

	it('coerces operands like Excel', () => {
		expect(ev('C1+1')).toBe(6);
		expect(ev('B3*2')).toBe(2);
		expect(ev('D1+1')).toBe(1);
		expect(ev('D1&"z"')).toBe('z');
		expect(ev('Z99=0')).toBe(true);
		expect(ev('Z99=""')).toBe(true);
		expect(ev('B1+1')).toEqual(ERRORS.VALUE);
		expect(ev('"1/2/2024"+0')).toBe(45293);
	});

	it('propagates errors', () => {
		expect(ev('1/0+NA()')).toEqual(ERRORS.DIV0);
		expect(ev('NA()&"x"')).toEqual(ERRORS.NA);
		expect(ev('-#REF!')).toEqual(ERRORS.REF);
		expect(ev('#N/A%')).toEqual(ERRORS.NA);
	});

	it('compares with type ordering and case-insensitive text', () => {
		expect(ev('B1="X"')).toBe(true);
		expect(ev('A1<B1')).toBe(true);
		expect(ev('B1<B3')).toBe(true);
		expect(ev('C1=5')).toBe(false);
	});
});

describe('array arithmetic (dynamic mode)', () => {
	const { engine } = setup();
	const ev = (f: string) => evalIn(engine, f, 'Sheet1!E2');

	it('broadcasts arrays and references element-wise', () => {
		expect(ev('{1,2;3,4}*10')).toEqual([
			[10, 20],
			[30, 40],
		]);
		expect(ev('{1,2,3}+{10;20}')).toEqual([
			[11, 12, 13],
			[21, 22, 23],
		]);
		expect(ev('{1,2}+{1,2,3}')).toEqual([[2, 4, ERRORS.NA]]);
		expect(ev('A1:A3*2')).toEqual([[2], [4], [6]]);
		expect(ev('A1:A3&"!"')).toEqual([['1!'], ['2!'], ['3!']]);
		expect(ev('-A1:A3')).toEqual([[-1], [-2], [-3]]);
		expect(ev('A1:A3>1')).toEqual([[false], [true], [true]]);
		expect(ev('A1:A3*{1,10}')).toEqual([
			[1, 10],
			[2, 20],
			[3, 30],
		]);
		expect(ev('SUM(A1:A3*A1:A3)')).toBe(14);
	});

	it('keeps element errors local', () => {
		expect(ev('{1,0}/{1,0}')).toEqual([[1, ERRORS.DIV0]]);
		expect(ev('B1:B3*1')).toEqual([[ERRORS.VALUE], [ERRORS.VALUE], [1]]);
	});
});

describe('references', () => {
	const { engine } = setup();
	const ev = (f: string, at = 'Sheet1!E2') => evalIn(engine, f, at);

	it('reads cells, ranges and other sheets', () => {
		expect(ev('A2')).toBe(2);
		expect(ev('Sheet2!A2')).toBe(20);
		expect(ev("'My Sheet'!A1")).toBe(100);
		expect(ev('sheet2!a1')).toBe(10);
		expect(ev('A1:A3')).toEqual([[1], [2], [3]]);
		expect(ev('Nope!A1')).toEqual(ERRORS.REF);
		expect(ev('[1]Sheet1!A1')).toEqual(ERRORS.REF);
	});

	it('supports the range, union and intersection operators', () => {
		expect(ev('SUM(A1:A2:A3)')).toBe(6);
		expect(ev('SUM(A1:INDEX(A1:A3,3))')).toBe(6);
		expect(ev('INDEX(A1:A3,1):INDEX(A1:A3,2)')).toEqual([[1], [2]]);
		expect(ev('SUM((A1,A3))')).toBe(4);
		expect(ev('SUM((A1:A3,Sheet1!A1))')).toBe(7);
		expect(ev('A1:B3 B2:C3')).toEqual([['Y'], [true]]);
		expect(ev('A1:A3 C1:C3')).toEqual(ERRORS.NULL);
		expect(ev('SUM(A1:A3 A2:B2)')).toBe(2);
		expect(ev('(A1,Sheet2!A1)')).toEqual(ERRORS.VALUE);
		expect(ev('A1:Sheet2!A1')).toEqual(ERRORS.REF);
	});

	it('sums 3-D references across sheets', () => {
		expect(ev("SUM('Sheet1:My Sheet'!A1)")).toBe(111);
		expect(ev("COUNT('Sheet1:My Sheet'!A1:A2)")).toBe(5);
		expect(ev("MAX('Sheet1:My Sheet'!A1)")).toBe(100);
		expect(ev('Sheet1:Sheet2!A1')).toBe('3D 0:1');
		expect(ev('Sheet1:Sheet2!A1+1')).toEqual(ERRORS.VALUE);
		expect(ev('VLOOKUP(1,Sheet1:Sheet2!A1:B2,2)')).toEqual(ERRORS.VALUE);
	});

	it('clamps whole columns and rows to the used range', () => {
		expect(ev('SUM(A:A)')).toBe(6);
		expect(ev('COUNTA(1:1)')).toBe(3);
		expect(ev('ROWS(A:A)')).toBe(1048576);
		expect(ev('A:A*1')).toEqual([[1], [2], [3]]);
		expect(ev('SUM(Sheet2!A:B)')).toBe(30);
	});

	it('implicitly intersects with @', () => {
		expect(ev('@A1:A3')).toBe(2);
		expect(ev('@A1:A3', 'Sheet1!E9')).toEqual(ERRORS.VALUE);
		expect(ev('@A1:C1', 'Sheet1!B9')).toBe('x');
		expect(ev('ISREF(@A1:A3)')).toBe(true);
		expect(ev('@{5,6}')).toBe(5);
		expect(ev('ROW(@A1:A3)', 'Sheet1!E3')).toBe(3);
	});
});

describe('functions', () => {
	const { engine } = setup();
	const ev = (f: string) => evalIn(engine, f, 'Sheet1!E2');

	it('reports unknown functions as #NAME?', () => {
		expect(ev('NOSUCHFUNCTION(1)')).toEqual(ERRORS.NAME);
		expect(ev('_xludf.MYMACRO(1)')).toEqual(ERRORS.NAME);
		expect(ev('nosuchname')).toEqual(ERRORS.NAME);
	});

	it('passes empty arguments and checks the argument count', () => {
		expect(ev('SUM(1,,2)')).toBe(3);
		expect(ev('IF(TRUE,,1)')).toBe(0);
		expect(ev('ABS()')).toEqual(ERRORS.VALUE);
		expect(ev('ABS(1,2)')).toEqual(ERRORS.VALUE);
	});

	it('lifts scalar functions over arrays in dynamic mode', () => {
		expect(ev('ABS({-1,2,-3})')).toEqual([[1, 2, 3]]);
		expect(ev('LEN(B1:B2)')).toEqual([[1], [1]]);
	});

	it('returns the raw result of evaluateFormula (leading = ignored)', () => {
		expect(ev('=1+1')).toBe(2);
		expect(engine.evaluateFormula('A1:A3', 0, 0, 4)).toMatchObject({
			kind: 'ref',
			range: { top: 0, bottom: 2 },
		});
		expect(engine.evaluateFormula('1+', 0, 0, 0)).toEqual(ERRORS.NAME);
	});
});
