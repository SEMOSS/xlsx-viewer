import { describe, expect, it } from 'vitest';

import type { AstNode } from '../ast';
import { FormulaParseError } from './parse-error';
import { MAX_FUNCTION_ARGS, parseFormula } from './parser';
import { tokenizeFormula } from './tokenizer';
import { parseFormulaTokens } from './parser';

/** Compact S-expression view of an AST, for readable precedence assertions. */
function show(node: AstNode): string {
	const p = node.parens ? `(${node.parens})` : '';
	switch (node.type) {
		case 'number':
			return `${node.value}${p}`;
		case 'string':
			return `"${node.value}"${p}`;
		case 'boolean':
			return `${node.value}${p}`;
		case 'error':
			return `${node.sheet ? `${node.sheet.sheet}!` : ''}${node.code}${p}`;
		case 'cell':
			return `cell[${node.sheet ? `${node.sheet.sheet}!` : ''}${node.ref.row},${node.ref.col}]${p}`;
		case 'range':
			return `${node.kind}[${node.start.row},${node.start.col}:${node.end.row},${node.end.col}]${p}`;
		case 'name':
			return `name:${node.name}${p}`;
		case 'structured':
			return `table:${node.text}${p}`;
		case 'function':
			return `${node.name}(${node.args.map(show).join(' ')})${p}`;
		case 'call':
			return `call[${show(node.callee)}](${node.args.map(show).join(' ')})${p}`;
		case 'unary':
			return `(${node.op} ${show(node.operand)})${p}`;
		case 'percent':
			return `(% ${show(node.operand)})${p}`;
		case 'binary':
			return `(${node.op === ' ' ? '_' : node.op} ${show(node.left)} ${show(node.right)})${p}`;
		case 'array':
			return `{${node.rows.map((r) => r.map(show).join(',')).join(';')}}${p}`;
		case 'missing':
			return '<missing>';
		case 'spill':
			return `spill[${show(node.ref)}]${p}`;
	}
}

const s = (formula: string) => show(parseFormula(formula));

function parseError(formula: string): FormulaParseError {
	try {
		parseFormula(formula);
	} catch (e) {
		expect(e).toBeInstanceOf(FormulaParseError);
		return e as FormulaParseError;
	}
	throw new Error(`expected ${formula} to fail`);
}

describe('parseFormula: precedence', () => {
	it.each([
		['1+2*3', '(+ 1 (* 2 3))'],
		['(1+2)*3', '(* (+ 1 2)(1) 3)'],
		['2^3^2', '(^ (^ 2 3) 2)'],
		['-2^2', '(^ (- 2) 2)'],
		['2^-2', '(^ 2 (- 2))'],
		['-2^-2', '(^ (- 2) (- 2))'],
		['-5%', '(% (- 5))'],
		['1+2%', '(+ 1 (% 2))'],
		['2^3%', '(^ 2 (% 3))'],
		['1%%', '(% (% 1))'],
		['--A1', '(- (- cell[0,0]))'],
		['2*-3^2', '(* 2 (^ (- 3) 2))'],
		['1-2-3', '(- (- 1 2) 3)'],
		['8/4/2', '(/ (/ 8 4) 2)'],
		['"a"&1+2', '(& "a" (+ 1 2))'],
		['1&2=3&4', '(= (& 1 2) (& 3 4))'],
		['1=1=TRUE', '(= (= 1 1) true)'],
		['1<>2', '(<> 1 2)'],
		['1<=2', '(<= 1 2)'],
		['1>=2', '(>= 1 2)'],
		['1<2', '(< 1 2)'],
		['1>2', '(> 1 2)'],
	])('%s', (formula, expected) => {
		expect(s(formula)).toBe(expected);
	});

	it('binds reference operators tighter than everything else', () => {
		expect(s('A1:A3 A2:B2')).toBe('(_ cells[0,0:2,0] cells[1,0:1,1])');
		expect(s('(A1,B1)')).toBe('(, cell[0,0] cell[0,1])(1)');
		expect(s('A1:B2 B1:C3,D4')).toBe('(, (_ cells[0,0:1,1] cells[0,1:2,2]) cell[3,3])');
		expect(s('-A1:INDEX(B:B,5)')).toBe('(- (: cell[0,0] INDEX(columns[0,1:1048575,1] 5)))');
		expect(s('@A1:A10')).toBe('(@ cells[0,0:9,0])');
		expect(s('A1:B2%')).toBe('(% cells[0,0:1,1])');
		expect(s('A1#:B5')).toBe('(: spill[cell[0,0]] cell[4,1])');
	});

	it('accepts an optional leading = and surrounding whitespace', () => {
		expect(s('=1+2')).toBe(s('1+2'));
		expect(s('= 1 + 2 ')).toBe('(+ 1 2)');
	});
});

describe('parseFormula: operands', () => {
	it('parses literals', () => {
		expect(parseFormula('1.5E+3')).toEqual({ type: 'number', value: 1500 });
		expect(parseFormula('.5')).toEqual({ type: 'number', value: 0.5 });
		expect(parseFormula('"a""b"')).toEqual({ type: 'string', value: 'a"b' });
		expect(parseFormula('""')).toEqual({ type: 'string', value: '' });
		expect(parseFormula('true')).toEqual({ type: 'boolean', value: true });
		expect(parseFormula('#n/a')).toEqual({ type: 'error', code: '#N/A' });
		expect(parseFormula('Sheet1!#REF!')).toEqual({
			type: 'error',
			code: '#REF!',
			sheet: { sheet: 'Sheet1' },
		});
		expect(parseFormula('#REF!A1')).toEqual({ type: 'error', code: '#REF!' });
	});

	it('parses references with sheets', () => {
		expect(parseFormula("'My Sheet'!$B$2")).toEqual({
			type: 'cell',
			sheet: { sheet: 'My Sheet' },
			ref: { row: 1, col: 1, rowAbs: true, colAbs: true },
		});
		expect(parseFormula('Sheet1:Sheet3!A1:B2')).toMatchObject({
			type: 'range',
			kind: 'cells',
			sheet: { sheet: 'Sheet1', sheetEnd: 'Sheet3' },
		});
		expect(parseFormula('[1]Sheet1!C:C')).toMatchObject({
			type: 'range',
			kind: 'columns',
			sheet: { workbook: '1' },
		});
		expect(parseFormula('2:2')).toMatchObject({
			type: 'range',
			kind: 'rows',
			start: { row: 1 },
			end: { row: 1 },
		});
	});

	it('parses names, stripping _xlpm. only from unqualified names', () => {
		expect(parseFormula('_xlpm.x')).toEqual({ type: 'name', name: 'x' });
		expect(parseFormula('Sheet1!Rate')).toEqual({
			type: 'name',
			name: 'Rate',
			sheet: { sheet: 'Sheet1' },
		});
		expect(parseFormula('Größe')).toEqual({ type: 'name', name: 'Größe' });
	});

	it('parses structured references', () => {
		expect(parseFormula('Table1[[#This Row],[Qty]]')).toEqual({
			type: 'structured',
			table: 'Table1',
			specifiers: ['#This Row'],
			columnStart: 'Qty',
			text: 'Table1[[#This Row],[Qty]]',
		});
		expect(parseFormula('[@Qty]')).toMatchObject({ specifiers: ['#This Row'], columnStart: 'Qty' });
		expect(parseFormula('[@Qty]')).not.toHaveProperty('table');
	});

	it('parses array constants', () => {
		expect(s('{1,-2;"x",#N/A}')).toBe('{1,-2;"x",#N/A}');
		expect(s('{+3}')).toBe('{3}');
		expect(s('{TRUE;FALSE}')).toBe('{true;false}');
		expect(parseFormula('{-0}')).toEqual({ type: 'array', rows: [[{ type: 'number', value: 0 }]] });
	});

	it('parses spill references', () => {
		expect(s('A1#')).toBe('spill[cell[0,0]]');
		expect(s('Sheet1!B2#')).toBe('spill[cell[Sheet1!1,1]]');
		expect(s('Totals#')).toBe('ANCHORARRAY(name:Totals)');
	});
});

describe('parseFormula: functions', () => {
	it('normalizes function names', () => {
		expect(s('sum(1)')).toBe('SUM(1)');
		expect(s('_xlfn.XLOOKUP(A1,B:B,C:C)')).toBe(
			'XLOOKUP(cell[0,0] columns[0,1:1048575,1] columns[0,2:1048575,2])',
		);
		expect(s('_xlfn._xlws.filter(A1:A5,B1:B5>0)')).toBe(
			'FILTER(cells[0,0:4,0] (> cells[0,1:4,1] 0))',
		);
		expect(s('norm.s.dist(1,TRUE)')).toBe('NORM.S.DIST(1 true)');
		expect(s('LOG10(100)')).toBe('LOG10(100)');
		expect(s('_xll.MyAddin(1)')).toBe('_XLL.MYADDIN(1)');
	});

	it('parses missing arguments', () => {
		expect(s('IF(A1,,1)')).toBe('IF(cell[0,0] <missing> 1)');
		expect(s('SUM()')).toBe('SUM()');
		expect(s('SUM( )')).toBe('SUM()');
		expect(s('F(,)')).toBe('F(<missing> <missing>)');
		expect(s('IF(A1,)')).toBe('IF(cell[0,0] <missing>)');
	});

	it('keeps unions inside argument groups', () => {
		expect(s('SUM((A1,B1),C1)')).toBe('SUM((, cell[0,0] cell[0,1])(1) cell[0,2])');
		expect(s('INDEX((A1:B2,C1:D2),1,1,2)')).toBe(
			'INDEX((, cells[0,0:1,1] cells[0,2:1,3])(1) 1 1 2)',
		);
	});

	it('maps ANCHORARRAY and SINGLE to their operators', () => {
		expect(s('_xlfn.ANCHORARRAY(B2)')).toBe('spill[cell[1,1]]');
		expect(s('_xlfn.ANCHORARRAY(B2:C3)')).toBe('ANCHORARRAY(cells[1,1:2,2])');
		expect(s('_xlfn.SINGLE(A1:A10)+1')).toBe('(+ (@ cells[0,0:9,0]) 1)');
		expect(s('SINGLE(A1+B1)')).toBe('(@ (+ cell[0,0] cell[0,1]))');
	});

	it('strips _xlpm. from LET and LAMBDA parameters', () => {
		expect(s('_xlfn.LET(_xlpm.x,1,_xlpm.x+1)')).toBe('LET(name:x 1 (+ name:x 1))');
		expect(s('_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2),_xlpm.f(3))')).toBe(
			'LET(name:f LAMBDA(name:x (* name:x 2)) F(3))',
		);
	});

	it('parses calls on expressions', () => {
		expect(s('LAMBDA(x,x+1)(2)')).toBe('call[LAMBDA(name:x (+ name:x 1))](2)');
		expect(s('(LAMBDA(x,x))(1)(2)')).toBe('call[call[LAMBDA(name:x name:x)(1)](1)](2)');
	});

	it(`accepts ${MAX_FUNCTION_ARGS} arguments and rejects more`, () => {
		const args = (n: number) => Array.from({ length: n }, (_, i) => String(i)).join(',');
		const node = parseFormula(`SUM(${args(255)})`);
		expect(node.type === 'function' && node.args.length).toBe(255);
		const err = parseError(`SUM(${args(256)})`);
		expect(err.message).toMatch(/255/);
	});
});

describe('parseFormula: parentheses metadata', () => {
	it('records parentheses written in the source', () => {
		expect(parseFormula('((A1))')).toMatchObject({ type: 'cell', parens: 2 });
		expect(parseFormula('(1)+2')).toMatchObject({ type: 'binary', left: { parens: 1 } });
		expect(parseFormula('1+2')).not.toHaveProperty('parens');
	});
});

describe('parseFormula: R1C1', () => {
	it('parses R1C1 references against a base cell', () => {
		expect(show(parseFormula('SUM(R[-1]C:R1C1)', { r1c1: { row: 5, col: 2 } }))).toBe(
			'SUM(cells[4,2:0,0])',
		);
	});

	it('parses pre-tokenized input', () => {
		expect(show(parseFormulaTokens(tokenizeFormula('=1+2')))).toBe('(+ 1 2)');
	});
});

describe('parseFormula: errors', () => {
	it.each([
		['', 'The formula is empty', 0],
		['=', 'The formula is empty', 1],
		['=1+', 'Unexpected end of formula', 3],
		['=SUM(1,2', 'Missing closing parenthesis', 8],
		['=(1+2', 'Missing closing parenthesis', 5],
		['=)', "Unexpected ')'", 1],
		['=SUM(A1))', "Unexpected ')'", 8],
		['=1 2', "Unexpected '2'", 3],
		['="abc', "Unterminated text value '\"abc'", 1],
		['=A1;B1', "Unexpected ';'", 3],
		['=1~2', "Unexpected '~'", 2],
		['=Sheet1!', "Unterminated reference 'Sheet1!'", 1],
		['=Table1[[#Bogus]]', "Invalid table reference 'Table1[[#Bogus]]'", 1],
	])('%j', (formula, message, position) => {
		const err = parseError(formula);
		expect(err.message).toBe(message);
		expect(err.position).toBe(position);
	});

	it('rejects invalid reference operator operands', () => {
		expect(parseError('=(A1+1):B2').message).toMatch(/range operator/);
		expect(parseError('="a":B2').position).toBe(4);
		expect(parseError('=(1,2)').message).toMatch(/union operator/);
		expect(parseError('=#N/A:A1').message).toMatch(/range operator/);
	});

	it('rejects invalid spill references and calls', () => {
		expect(parseError('=A1:B2#').message).toMatch(/spill operator/);
		expect(parseError('=(A1)(B1)').message).toMatch(/can be called/);
	});

	it('rejects invalid array constants', () => {
		expect(parseError('={1,2;3}').message).toMatch(/same number of values/);
		expect(parseError('={A1}').message).toMatch(/Array constants can only contain/);
		expect(parseError('={1,{2}}').position).toBe(4);
		expect(parseError('={}').position).toBe(2);
		expect(parseError('={1,,2}').position).toBe(4);
		expect(parseError('={-"a"}').message).toMatch(/signed/);
		expect(parseError('={1,2').message).toMatch(/Missing closing brace/);
		expect(parseError('={Sheet1!#REF!}').message).toMatch(/Array constants/);
	});

	it('rejects out-of-range numbers', () => {
		expect(parseError('=1E400').message).toMatch(/out of range/);
	});

	it('limits nesting depth instead of overflowing the stack', () => {
		const deep = `=${'('.repeat(300)}1${')'.repeat(300)}`;
		expect(parseError(deep).message).toMatch(/nested too deeply/);
		const ok = `=${'('.repeat(100)}1${')'.repeat(100)}`;
		expect(parseFormula(ok)).toMatchObject({ type: 'number', parens: 100 });
		expect(parseFormula(`=${'-'.repeat(5000)}1`).type).toBe('unary');
	});
});
