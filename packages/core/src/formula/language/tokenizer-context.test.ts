import { describe, expect, it } from 'vitest';

import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../../address/constants';
import { tokenizeFormula } from './tokenizer';
import type { FormulaToken } from './tokens';

/** `type:text` of every token except whitespace and the leading `=`. */
function sig(formula: string): string[] {
	return tokenizeFormula(formula)
		.filter((t) => t.type !== 'whitespace' && t.type !== 'equals')
		.map((t) => `${t.type}:${t.text}`);
}

function only(formula: string): FormulaToken {
	const tokens = tokenizeFormula(formula).filter((t) => t.type !== 'equals');
	expect(tokens).toHaveLength(1);
	return tokens[0];
}

describe('tokenizeFormula: names, functions and tables', () => {
	it('reads dotted, underscored, backslash and unicode names', () => {
		expect(sig('=\\name+_x.y+Größe+日本語+x?')).toEqual([
			'name:\\name',
			'operator:+',
			'name:_x.y',
			'operator:+',
			'name:Größe',
			'operator:+',
			'name:日本語',
			'operator:+',
			'name:x?',
		]);
	});

	it('reads prefixed and dotted function names', () => {
		expect(sig('=_xlfn._xlws.FILTER(A1:A5,B1:B5>0)')[0]).toBe('function:_xlfn._xlws.FILTER');
		expect(sig('=NORM.S.DIST(1,TRUE)')[0]).toBe('function:NORM.S.DIST');
		expect(sig('=SUM (A1)')).toEqual([
			'name:SUM',
			'intersect: ',
			'paren-open:(',
			'ref:A1',
			'paren-close:)',
		]);
	});

	it.each([
		'Table1[Col]',
		'Table1[]',
		'Table1[#All]',
		'Table1[[#Headers],[A]:[C]]',
		'Table1[@Qty]',
		'Table1[@[Unit Price]]',
		'[@Qty]',
		'[Sales]',
		'[[#This Row],[Qty]]',
		"Table1[Price'[USD']]",
		"Table1[['#Items]]",
		'Table1[ [Col1] , [Col2] ]',
	])('structured reference %s is one token', (text) => {
		expect(only(text)).toMatchObject({ type: 'structured', text });
	});

	it('distinguishes structured references from external prefixes', () => {
		expect(sig('=[@Qty]*2')).toEqual(['structured:[@Qty]', 'operator:*', 'number:2']);
		expect(sig('=[1]Sheet1!A1')).toEqual(['ref:[1]Sheet1!A1']);
	});

	it('flags unterminated structured references', () => {
		expect(only('Table1[[#Headers],[A')).toMatchObject({ type: 'structured', incomplete: true });
	});
});

describe('tokenizeFormula: operators and context', () => {
	it('reads every operator', () => {
		expect(sig('=1+2-3*4/5^6&7=8<>9<1>2<=3>=4')).toEqual([
			'number:1',
			'operator:+',
			'number:2',
			'operator:-',
			'number:3',
			'operator:*',
			'number:4',
			'operator:/',
			'number:5',
			'operator:^',
			'number:6',
			'operator:&',
			'number:7',
			'operator:=',
			'number:8',
			'operator:<>',
			'number:9',
			'operator:<',
			'number:1',
			'operator:>',
			'number:2',
			'operator:<=',
			'number:3',
			'operator:>=',
			'number:4',
		]);
	});

	it('tells unary from binary + and -', () => {
		expect(sig('=-2^2')).toEqual(['prefix:-', 'number:2', 'operator:^', 'number:2']);
		expect(sig('=+-A1')).toEqual(['prefix:+', 'prefix:-', 'ref:A1']);
		expect(sig('=A1 -1')).toEqual(['ref:A1', 'operator:-', 'number:1']);
		expect(sig('=2*-3')).toEqual(['number:2', 'operator:*', 'prefix:-', 'number:3']);
		expect(sig('=(1)-2')[3]).toBe('operator:-');
		expect(sig('=1%-2')[2]).toBe('operator:-');
	});

	it('reads postfix percent and the @ operator', () => {
		expect(sig('=1%%')).toEqual(['number:1', 'percent:%', 'percent:%']);
		expect(sig('=@A1:A10')).toEqual(['prefix:@', 'ref:A1:A10']);
	});

	it('distinguishes argument commas from union commas', () => {
		expect(sig('=(A1,B1)')).toEqual([
			'paren-open:(',
			'ref:A1',
			'union:,',
			'ref:B1',
			'paren-close:)',
		]);
		expect(sig('=SUM(A1,B1)')).toEqual([
			'function:SUM',
			'call-open:(',
			'ref:A1',
			'arg-sep:,',
			'ref:B1',
			'call-close:)',
		]);
		expect(sig('=SUM((A1,B1),C1)')).toEqual([
			'function:SUM',
			'call-open:(',
			'paren-open:(',
			'ref:A1',
			'union:,',
			'ref:B1',
			'paren-close:)',
			'arg-sep:,',
			'ref:C1',
			'call-close:)',
		]);
		expect(sig('=Sheet1!$A$1:$B$2,Sheet1!$D$1')[1]).toBe('union:,');
	});

	it('reads array constants', () => {
		expect(sig('={1,-2;"x",#N/A}')).toEqual([
			'array-open:{',
			'number:1',
			'array-col-sep:,',
			'prefix:-',
			'number:2',
			'array-row-sep:;',
			'string:"x"',
			'array-col-sep:,',
			'error:#N/A',
			'array-close:}',
		]);
		expect(sig('=SUM({1,2},3)')).toEqual([
			'function:SUM',
			'call-open:(',
			'array-open:{',
			'number:1',
			'array-col-sep:,',
			'number:2',
			'array-close:}',
			'arg-sep:,',
			'number:3',
			'call-close:)',
		]);
		expect(sig('=A1;B1')[1]).toBe('unknown:;');
	});

	it('reads whitespace between references as intersection', () => {
		expect(sig('=SUM(A1:A3 A2:B2)')).toEqual([
			'function:SUM',
			'call-open:(',
			'ref:A1:A3',
			'intersect: ',
			'ref:A2:B2',
			'call-close:)',
		]);
		expect(sig('=A1  \n B1')[1]).toBe('intersect:  \n ');
		expect(sig('=(A1) (B1)')[3]).toBe('intersect: ');
		expect(sig('=Sales Q1')).toEqual(['name:Sales', 'intersect: ', 'ref:Q1']);
		expect(sig('=A1 INDEX(B:B,1)')[1]).toBe('intersect: ');
		expect(sig('=#REF! A1')[1]).toBe('intersect: ');
	});

	it('keeps other whitespace as plain whitespace', () => {
		const types = (f: string) =>
			tokenizeFormula(f)
				.filter((t) => t.text.trim() === '')
				.map((t) => t.type);
		expect(types('=SUM( A1 , B1 )')).toEqual([
			'whitespace',
			'whitespace',
			'whitespace',
			'whitespace',
		]);
		expect(types('=A1 + B1')).toEqual(['whitespace', 'whitespace']);
		expect(types('=1 2')).toEqual(['whitespace']);
		expect(types('={1, 2}')).toEqual(['whitespace']);
		expect(types('=A1 ')).toEqual(['whitespace']);
	});

	it('calls the result of a call or group', () => {
		expect(sig('=LAMBDA(x,x+1)(2)')).toEqual([
			'function:LAMBDA',
			'call-open:(',
			'name:x',
			'arg-sep:,',
			'name:x',
			'operator:+',
			'number:1',
			'call-close:)',
			'call-open:(',
			'number:2',
			'call-close:)',
		]);
		expect(sig('=(f)(1)')[3]).toBe('call-open:(');
	});

	it('marks stray characters as unknown', () => {
		expect(sig('=A1 ! 2')).toContain('unknown:!');
		expect(sig('=1~2')).toContain('unknown:~');
		expect(sig('=$A')).toEqual(['unknown:$A']);
	});
});

describe('tokenizeFormula: R1C1 mode', () => {
	const r1c1 = (f: string, row = 4, col = 4) => tokenizeFormula(f, { r1c1: { row, col } });

	it('resolves relative and absolute parts against the base cell', () => {
		const [t] = r1c1('R[-1]C[2]');
		expect(t.area?.start).toEqual({ row: 3, col: 6, rowAbs: false, colAbs: false });
		expect(r1c1('R1C1')[0].area?.start).toEqual({ row: 0, col: 0, rowAbs: true, colAbs: true });
		expect(r1c1('RC')[0].area?.start).toEqual({ row: 4, col: 4, rowAbs: false, colAbs: false });
		expect(r1c1('r2c')[0].area?.start).toEqual({ row: 1, col: 4, rowAbs: true, colAbs: false });
	});

	it('reads whole rows, whole columns and ranges', () => {
		expect(r1c1('R2')[0].area?.kind).toBe('rows');
		expect(r1c1('C[1]:C[3]')[0].area).toMatchObject({
			kind: 'columns',
			start: { col: 5 },
			end: { col: 7 },
		});
		expect(r1c1('R1C1:R2C2')[0].area?.kind).toBe('cells');
		expect(r1c1('Sheet1!R1C1')[0].sheet).toEqual({ sheet: 'Sheet1' });
	});

	it('leaves names and functions alone and treats A1 text as names', () => {
		const types = r1c1('ROUND(RC,0)+Rate+A1').map((t) => t.type);
		expect(types).toEqual([
			'function',
			'call-open',
			'ref',
			'arg-sep',
			'number',
			'call-close',
			'operator',
			'name',
			'operator',
			'name',
		]);
	});

	it('wraps relative offsets around the grid', () => {
		expect(r1c1('R[-1]C', 0, 0)[0].area?.start.row).toBe(MAX_ROW_INDEX);
		expect(r1c1('RC[-1]', 0, 0)[0].area?.start.col).toBe(MAX_COL_INDEX);
	});
});
