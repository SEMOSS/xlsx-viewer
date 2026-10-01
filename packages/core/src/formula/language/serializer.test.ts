import { describe, expect, it } from 'vitest';

import type { AstNode, CellRefNode, RefCorner } from '../ast';
import { parseFormula } from './parser';
import { formatFormulaNumber, serializeFormula } from './serializer';

const corner = (row: number, col: number, abs = false): RefCorner => ({
	row,
	col,
	rowAbs: abs,
	colAbs: abs,
});
const cell = (row: number, col: number, sheet?: string): CellRefNode =>
	sheet === undefined
		? { type: 'cell', ref: corner(row, col) }
		: { type: 'cell', ref: corner(row, col), sheet: { sheet } };
const num = (value: number): AstNode => ({ type: 'number', value });
const bin = (
	op: '+' | '-' | '*' | '/' | '^' | '&' | '=' | ':' | ',' | ' ',
	left: AstNode,
	right: AstNode,
): AstNode => ({
	type: 'binary',
	op,
	left,
	right,
});
const neg = (operand: AstNode): AstNode => ({ type: 'unary', op: '-', operand });
const fn = (name: string, ...args: AstNode[]): AstNode => ({ type: 'function', name, args });

const display = (formula: string) => serializeFormula(parseFormula(formula));
const file = (formula: string) => serializeFormula(parseFormula(formula), { forFile: true });

describe('serializeFormula: round trips', () => {
	it.each([
		'SUM(A1:A3 A2:B2)',
		'(A1,B1)',
		"'It''s'!A1",
		'Table1[[#This Row],[Qty]]*2',
		'{1,-2;"x",#N/A}',
		'-2^2',
		'1%%',
		'A1#',
		'LET(x,1,x+1)',
		'Größe*2',
		'(1+2)*3',
		'((A1))',
		'1-(2-3)',
		'2^-2',
		'-(2^2)',
		'(-2)^2',
		'IF(A1>0,"yes","no")',
		'IF(A1,,1)',
		'SUM()',
		'SUM((A1,B1),C1)',
		'INDEX((A1:B2,C1:D2),1,1,2)',
		'LAMBDA(x,x+1)(2)',
		'Sheet1:Sheet3!$A$1:B2',
		"'Sheet 1:Sheet 3'!A1",
		'[1]Sheet1!A1',
		"'[Book 1.xlsx]Data'!A1",
		'[1]!Total',
		'Sheet1!#REF!',
		'Sheet1!Rate*2',
		'$A:$C',
		'1:3',
		'A1:INDEX(B:B,5)',
		'@A1:A10',
		'[@Qty]*Table1[@[Unit Price]]',
		'"a""b"&""',
		'TRUE<>FALSE',
		'A1<=B1',
		'-A1%',
		'---1',
		'1--1',
		'NORM.S.DIST(1,TRUE)',
		'Totals#',
		'(LAMBDA(x,x))(1)',
		'A1:B2 B1:C3,D4',
	])('%s', (formula) => {
		expect(display(formula)).toBe(formula);
		// Printing the parse of the printed text is stable.
		expect(display(display(formula))).toBe(display(formula));
	});

	it('normalizes names, spacing, numbers and dead references', () => {
		expect(display('= sum( a1 , b1 )')).toBe('SUM(A1,B1)');
		expect(display('1E+10+.5+1e-3')).toBe('10000000000+0.5+0.001');
		expect(display('_xlfn.XLOOKUP(A1,B:B,C:C)')).toBe('XLOOKUP(A1,B:B,C:C)');
		expect(display('#REF!A1+1')).toBe('#REF!+1');
		expect(display("'Sheet1'!A1")).toBe('Sheet1!A1');
		expect(display('_xlfn.SINGLE(A1+B1)')).toBe('@(A1+B1)');
	});
});

describe('serializeFormula: minimal parentheses for built trees', () => {
	it.each<[string, AstNode, string]>([
		['sum then product', bin('*', bin('+', num(1), num(2)), num(3)), '(1+2)*3'],
		['product then sum', bin('+', bin('*', num(1), num(2)), num(3)), '1*2+3'],
		['right-nested minus', bin('-', num(1), bin('-', num(2), num(3))), '1-(2-3)'],
		['left-nested minus', bin('-', bin('-', num(1), num(2)), num(3)), '1-2-3'],
		['right-nested power', bin('^', num(2), bin('^', num(3), num(2))), '2^(3^2)'],
		['negated power', neg(bin('^', num(2), num(2))), '-(2^2)'],
		['power of negation', bin('^', neg(num(2)), num(2)), '-2^2'],
		['negative literal base', bin('^', num(-2), num(2)), '-2^2'],
		['percent of sum', { type: 'percent', operand: bin('+', num(1), num(2)) }, '(1+2)%'],
		['concat of sum', bin('&', num(1), bin('+', num(2), num(3))), '1&2+3'],
		['sum of concat', bin('+', num(1), bin('&', num(2), num(3))), '1+(2&3)'],
		['comparison of comparisons', bin('=', num(1), bin('=', num(2), num(3))), '1=(2=3)'],
		['root union', bin(',', cell(0, 0), cell(0, 1)), 'A1,B1'],
		['union argument', fn('SUM', bin(',', cell(0, 0), cell(0, 1))), 'SUM((A1,B1))'],
		['negated union argument', fn('SUM', neg(bin(',', cell(0, 0), cell(0, 1)))), 'SUM(-(A1,B1))'],
		['range of intersection', bin(':', cell(0, 0), bin(' ', cell(1, 1), cell(2, 2))), 'A1:(B2 C3)'],
		['intersection of ranges', bin(' ', bin(':', cell(0, 0), cell(1, 1)), cell(2, 2)), 'A1:B2 C3'],
		['negated range', neg(bin(':', cell(0, 0), cell(1, 1))), '-A1:B2'],
		['@ of sum', { type: 'unary', op: '@', operand: bin('+', cell(0, 0), num(1)) }, '@(A1+1)'],
		[
			'range with @',
			bin(':', { type: 'unary', op: '@', operand: cell(0, 0) }, cell(4, 1)),
			'(@A1):B5',
		],
		['call of sum', { type: 'call', callee: bin('+', num(1), num(2)), args: [] }, '(1+2)()'],
	])('%s', (_label, node, expected) => {
		expect(serializeFormula(node)).toBe(expected);
	});
});

describe('serializeFormula: literals and sheets', () => {
	it.each([
		[0, '0'],
		[-0, '0'],
		[1.5, '1.5'],
		[-3, '-3'],
		[1e21, '1E+21'],
		[1.5e-7, '1.5E-7'],
		[123456789012345680000, '123456789012345680000'],
		[0.1 + 0.2, '0.30000000000000004'],
		[Infinity, '#NUM!'],
	])('number %d', (value, text) => {
		expect(formatFormulaNumber(value)).toBe(text);
	});

	it('escapes strings', () => {
		expect(serializeFormula({ type: 'string', value: 'say "hi"' })).toBe('"say ""hi"""');
	});

	it('quotes sheet names that need it', () => {
		expect(serializeFormula(cell(0, 0, 'My Sheet'))).toBe("'My Sheet'!A1");
		expect(serializeFormula(cell(0, 0, "Bob's"))).toBe("'Bob''s'!A1");
		expect(serializeFormula(cell(1, 1, 'A1'))).toBe("'A1'!B2");
		expect(serializeFormula(cell(1, 1, 'R1C1'))).toBe("'R1C1'!B2");
		expect(serializeFormula(cell(0, 0, '2024'))).toBe("'2024'!A1");
		expect(serializeFormula(cell(0, 0, 'Data'))).toBe('Data!A1');
		expect(
			serializeFormula({
				type: 'cell',
				ref: corner(0, 0, true),
				sheet: { sheet: 'Jan', sheetEnd: 'Dec 2' },
			}),
		).toBe("'Jan:Dec 2'!$A$1");
		expect(
			serializeFormula({
				type: 'cell',
				ref: corner(0, 0),
				sheet: { sheet: 'Sheet 1', workbook: 'C:\\Docs\\Book.xlsx' },
			}),
		).toBe("'C:\\Docs\\[Book.xlsx]Sheet 1'!A1");
		expect(
			serializeFormula({ type: 'name', name: 'Total', sheet: { sheet: '', workbook: '1' } }),
		).toBe('[1]!Total');
	});

	it('regenerates structured references without source text', () => {
		const node: AstNode = {
			type: 'structured',
			table: 'T',
			specifiers: ['#Headers'],
			columnStart: 'A',
			columnEnd: 'C',
			text: '',
		};
		expect(serializeFormula(node)).toBe('T[[#Headers],[A]:[C]]');
	});
});

describe('serializeFormula: file form', () => {
	it('prefixes newer functions', () => {
		expect(file('XLOOKUP(A1,B:B,C:C)')).toBe('_xlfn.XLOOKUP(A1,B:B,C:C)');
		expect(file('SORT(FILTER(A1:A5,B1:B5>0))')).toBe(
			'_xlfn._xlws.SORT(_xlfn._xlws.FILTER(A1:A5,B1:B5>0))',
		);
		expect(file('SUM(A1:A3)+CONCAT("a",B1)')).toBe('SUM(A1:A3)+_xlfn.CONCAT("a",B1)');
	});

	it('writes spill and implicit intersection as functions', () => {
		expect(file('SUM(A1#)')).toBe('SUM(_xlfn.ANCHORARRAY(A1))');
		expect(file('@A1:A10*2')).toBe('_xlfn.SINGLE(A1:A10)*2');
		expect(file('@(A1:A10,B1)')).toBe('_xlfn.SINGLE((A1:A10,B1))');
		expect(file('Totals#')).toBe('_xlfn.ANCHORARRAY(Totals)');
	});

	it('prefixes LET and LAMBDA parameters where they are bound', () => {
		expect(file('LET(x,1,x+1)')).toBe('_xlfn.LET(_xlpm.x,1,_xlpm.x+1)');
		expect(file('LET(x,x,x)')).toBe('_xlfn.LET(_xlpm.x,x,_xlpm.x)');
		expect(file('LET(a,1,b,a+1,a+b)')).toBe(
			'_xlfn.LET(_xlpm.a,1,_xlpm.b,_xlpm.a+1,_xlpm.a+_xlpm.b)',
		);
		expect(file('LAMBDA(x,y,x*y)(2,3)')).toBe('_xlfn.LAMBDA(_xlpm.x,_xlpm.y,_xlpm.x*_xlpm.y)(2,3)');
		expect(file('MAP(A1:A3,LAMBDA(v,v*2))+v')).toBe(
			'_xlfn.MAP(A1:A3,_xlfn.LAMBDA(_xlpm.v,_xlpm.v*2))+v',
		);
		expect(file('LET(f,LAMBDA(x,x+1),f(2))')).toBe(
			'_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1),_xlpm.f(2))',
		);
		expect(file('LET(rows,5,ROWS(A1:A3)+rows)')).toBe(
			'_xlfn.LET(_xlpm.rows,5,ROWS(A1:A3)+_xlpm.rows)',
		);
		expect(file('LET(x,1,LET(y,x,y*2))')).toBe(
			'_xlfn.LET(_xlpm.x,1,_xlfn.LET(_xlpm.y,_xlpm.x,_xlpm.y*2))',
		);
		expect(file('LET(x,1,Sheet1!x)')).toBe('_xlfn.LET(_xlpm.x,1,Sheet1!x)');
	});

	it('uses the file syntax for this-row table references', () => {
		expect(file('[@Qty]*2')).toBe('[[#This Row],[Qty]]*2');
		expect(file('Table1[@[Unit Price]]')).toBe('Table1[[#This Row],[Unit Price]]');
		expect(file('Table1[@]')).toBe('Table1[#This Row]');
		expect(file('Table1[[#Headers],[A]]')).toBe('Table1[[#Headers],[A]]');
	});

	it('is the inverse of parsing file formulas', () => {
		for (const f of [
			'_xlfn.LET(_xlpm.x,1,_xlpm.x+1)',
			'_xlfn._xlws.FILTER(A1:A5,B1:B5>0)',
			'SUM(_xlfn.ANCHORARRAY(B2))',
			'_xlfn.SINGLE(A1:A10)+1',
			'Table1[[#This Row],[Qty]]*2',
			'_xlfn.LAMBDA(_xlpm.a,_xlpm.b,_xlpm.a+_xlpm.b)(1,2)',
		]) {
			expect(file(f)).toBe(f);
		}
	});
});
