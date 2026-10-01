import { describe, expect, it } from 'vitest';

import {
	addFormulaPrefixes,
	filePrefixFor,
	mapCodeSegments,
	stripFormulaPrefixes,
	XLFN_FUNCTIONS,
} from '../prefixes';
import { parseFormula } from './parser';
import { serializeFormula } from './serializer';

/** File form (as stored in sheet XML) -> display form. */
const PAIRS: Array<[string, string]> = [
	[
		'_xlfn.XLOOKUP(A2,Sheet2!$A:$A,Sheet2!$B:$B,"Not found")',
		'XLOOKUP(A2,Sheet2!$A:$A,Sheet2!$B:$B,"Not found")',
	],
	['_xlfn._xlws.SORT(_xlfn.UNIQUE(A2:A100))', 'SORT(UNIQUE(A2:A100))'],
	[
		'_xlfn._xlws.FILTER(Table1[Name],Table1[Score]>=90,"")',
		'FILTER(Table1[Name],Table1[Score]>=90,"")',
	],
	[
		'_xlfn.LET(_xlpm.total,SUM(A1:A10),_xlpm.n,COUNT(A1:A10),IF(_xlpm.n=0,0,_xlpm.total/_xlpm.n))',
		'LET(total,SUM(A1:A10),n,COUNT(A1:A10),IF(n=0,0,total/n))',
	],
	['_xlfn.LAMBDA(_xlpm.x,_xlpm.y,_xlpm.x*_xlpm.y)(2,3)', 'LAMBDA(x,y,x*y)(2,3)'],
	['_xlfn.MAP(A1:A3,_xlfn.LAMBDA(_xlpm.v,_xlpm.v*2))', 'MAP(A1:A3,LAMBDA(v,v*2))'],
	[
		'_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1),_xlpm.f(_xlpm.f(1)))',
		'LET(f,LAMBDA(x,x+1),f(f(1)))',
	],
	['SUM(_xlfn.ANCHORARRAY(B2))', 'SUM(B2#)'],
	['_xlfn.ANCHORARRAY(Sheet1!B2)*2', 'Sheet1!B2#*2'],
	['_xlfn.SINGLE(A1:A10)*2', '@A1:A10*2'],
	['_xlfn.SINGLE(A1:A10+1)', '@(A1:A10+1)'],
	['_xlfn.SINGLE(INDEX(A:A,1))', '@INDEX(A:A,1)'],
	['-_xlfn.SINGLE(A1:A3)^2', '-@A1:A3^2'],
	[
		'Table1[[#This Row],[Qty]]*Table1[[#This Row],[Price]]',
		'Table1[[#This Row],[Qty]]*Table1[[#This Row],[Price]]',
	],
	['_xlfn.CONCAT("_xlfn.KEEP",\'_xlfn.Sheet\'!A1)', 'CONCAT("_xlfn.KEEP",\'_xlfn.Sheet\'!A1)'],
	['_xlfn.STDEV.S(A1:A9)+_xlfn.NORM.S.DIST(1,TRUE)', 'STDEV.S(A1:A9)+NORM.S.DIST(1,TRUE)'],
	['IFERROR(_xlfn.XMATCH(5,{1,5,9}),-1)', 'IFERROR(XMATCH(5,{1,5,9}),-1)'],
	['SUM(A1:A3)', 'SUM(A1:A3)'],
];

describe('stripFormulaPrefixes', () => {
	it.each(PAIRS)('%s', (fileForm, displayForm) => {
		expect(stripFormulaPrefixes(fileForm)).toBe(displayForm);
	});

	it('preserves spacing and line breaks', () => {
		expect(stripFormulaPrefixes('_xlfn.XLOOKUP( A1 ,\n  B:B , C:C )')).toBe(
			'XLOOKUP( A1 ,\n  B:B , C:C )',
		);
		expect(stripFormulaPrefixes('_xlfn.LET(\n  _xlpm.x, 1,\n  _xlpm.x + 1\n)')).toBe(
			'LET(\n  x, 1,\n  x + 1\n)',
		);
	});

	it('keeps ANCHORARRAY when its argument is not a single cell', () => {
		expect(stripFormulaPrefixes('_xlfn.ANCHORARRAY(B2:C3)')).toBe('ANCHORARRAY(B2:C3)');
		expect(stripFormulaPrefixes('_xlfn.ANCHORARRAY( B2 )')).toBe('ANCHORARRAY( B2 )');
	});

	it('parenthesizes @ where reference operators would capture more', () => {
		expect(stripFormulaPrefixes('_xlfn.SINGLE(INDEX(A:A,1)):B5')).toBe('(@INDEX(A:A,1)):B5');
		expect(stripFormulaPrefixes('B5:_xlfn.SINGLE(A1)')).toBe('B5:(@A1)');
		expect(stripFormulaPrefixes('(_xlfn.SINGLE(A1),B1)')).toBe('((@A1),B1)');
	});

	it('keeps the meaning of SINGLE', () => {
		const tree = (f: string) =>
			JSON.stringify(parseFormula(f), (k, v: unknown) => (k === 'parens' ? undefined : v));
		for (const f of [
			'_xlfn.SINGLE(INDEX(A:A,1)):B5',
			'B5:_xlfn.SINGLE(A1)',
			'(_xlfn.SINGLE(A1),B1)',
			'_xlfn.SINGLE(A1:A10+1)',
			'-_xlfn.SINGLE(A1:A3)^2',
			'_xlfn.SINGLE(A1)%',
		]) {
			expect(tree(stripFormulaPrefixes(f))).toBe(tree(f));
		}
		expect(
			serializeFormula(parseFormula(stripFormulaPrefixes('_xlfn.SINGLE(A1:A10+1)')), {
				forFile: true,
			}),
		).toBe('_xlfn.SINGLE(A1:A10+1)');
	});

	it('handles empty and prefix-free input fast', () => {
		expect(stripFormulaPrefixes('')).toBe('');
		const plain = 'A1+B1';
		expect(stripFormulaPrefixes(plain)).toBe(plain);
	});

	it('falls back to text replacement for formulas it cannot tokenize', () => {
		expect(stripFormulaPrefixes('_xlfn.XLOOKUP(A1;B1)~"_xlfn.x"')).toBe('XLOOKUP(A1;B1)~"_xlfn.x"');
		expect(stripFormulaPrefixes('_xlfn.CONCAT("unterminated')).toBe('CONCAT("unterminated');
	});
});

describe('addFormulaPrefixes', () => {
	it.each(PAIRS)('%s', (fileForm, displayForm) => {
		expect(addFormulaPrefixes(displayForm)).toBe(fileForm);
	});

	it('round-trips file formulas', () => {
		for (const [fileForm] of PAIRS) {
			expect(addFormulaPrefixes(stripFormulaPrefixes(fileForm))).toBe(fileForm);
		}
	});

	it('is idempotent', () => {
		for (const [fileForm] of PAIRS) {
			expect(addFormulaPrefixes(fileForm)).toBe(fileForm);
		}
	});

	it('writes the this-row table shorthand in file syntax', () => {
		expect(addFormulaPrefixes('[@Qty]*2')).toBe('[[#This Row],[Qty]]*2');
		expect(addFormulaPrefixes('Table1[@[Unit Price]]+Table1[@]')).toBe(
			'Table1[[#This Row],[Unit Price]]+Table1[#This Row]',
		);
	});

	it('writes @ and spill references as functions', () => {
		expect(addFormulaPrefixes('@(A1+B1)')).toBe('_xlfn.SINGLE(A1+B1)');
		expect(addFormulaPrefixes('@ A1:A3')).toBe('_xlfn.SINGLE( A1:A3)');
		expect(addFormulaPrefixes('@@A1')).toBe('_xlfn.SINGLE(_xlfn.SINGLE(A1))');
		expect(addFormulaPrefixes('@A1#')).toBe('_xlfn.SINGLE(_xlfn.ANCHORARRAY(A1))');
		// The union binds tighter than @, as in `@A1:B2`.
		expect(addFormulaPrefixes('(@A1,B1)')).toBe('(_xlfn.SINGLE((A1,B1)))');
		expect(addFormulaPrefixes('((@A1),B1)')).toBe('((_xlfn.SINGLE(A1)),B1)');
		expect(addFormulaPrefixes('@A1,B1')).toBe('_xlfn.SINGLE((A1,B1))');
		expect(addFormulaPrefixes('Totals#')).toBe('_xlfn.ANCHORARRAY(Totals)');
	});

	it('binds LET names only after their value and calls only to LAMBDA values', () => {
		expect(addFormulaPrefixes('LET(x,x,x)')).toBe('_xlfn.LET(_xlpm.x,x,_xlpm.x)');
		expect(addFormulaPrefixes('LET(x,1,y)')).toBe('_xlfn.LET(_xlpm.x,1,y)');
		expect(addFormulaPrefixes('LET(rows,5,ROWS(A1:A3)+rows)')).toBe(
			'_xlfn.LET(_xlpm.rows,5,ROWS(A1:A3)+_xlpm.rows)',
		);
		expect(addFormulaPrefixes('x+LET(x,1,x)+x')).toBe('x+_xlfn.LET(_xlpm.x,1,_xlpm.x)+x');
		expect(addFormulaPrefixes('LAMBDA(f,v,f(v))(SQRT,4)')).toBe(
			'_xlfn.LAMBDA(_xlpm.f,_xlpm.v,_xlpm.f(_xlpm.v))(SQRT,4)',
		);
	});

	it('normalizes partial prefixes', () => {
		expect(addFormulaPrefixes('_xlfn.FILTER(A1:A3,B1:B3)')).toBe('_xlfn._xlws.FILTER(A1:A3,B1:B3)');
		expect(addFormulaPrefixes('_xlfn.SUM(1)')).toBe('SUM(1)');
	});

	it('leaves strings, sheet names and defined names alone', () => {
		expect(addFormulaPrefixes('"CONCAT(1)"&\'CONCAT Sheet\'!A1&CONCAT.X')).toBe(
			'"CONCAT(1)"&\'CONCAT Sheet\'!A1&CONCAT.X',
		);
	});

	it('falls back to text replacement for formulas it cannot tokenize', () => {
		expect(addFormulaPrefixes('CONCAT(A1;B1)')).toBe('_xlfn.CONCAT(A1;B1)');
	});

	it('converts formulas without any function call', () => {
		expect(addFormulaPrefixes('A1#*2')).toBe('_xlfn.ANCHORARRAY(A1)*2');
		expect(addFormulaPrefixes('@A1:A9')).toBe('_xlfn.SINGLE(A1:A9)');
		expect(addFormulaPrefixes('[@Qty]')).toBe('[[#This Row],[Qty]]');
		const plain = 'A1+Sheet2!B2*Rate';
		expect(addFormulaPrefixes(plain)).toBe(plain);
	});
});

describe('prefix tables and legacy helper', () => {
	it('knows the file prefix of newer functions', () => {
		expect(filePrefixFor('xlookup')).toBe('_xlfn.');
		expect(filePrefixFor('FILTER')).toBe('_xlfn._xlws.');
		expect(filePrefixFor('SUM')).toBe('');
		expect(XLFN_FUNCTIONS.has('ACOT')).toBe(true);
	});

	it('maps only code outside strings and quoted names', () => {
		expect(mapCodeSegments('a"b"c\'d\'e', (code) => code.toUpperCase())).toBe('A"b"C\'d\'E');
		expect(mapCodeSegments('"it""s"x', (code) => code.toUpperCase())).toBe('"it""s"X');
	});
});
