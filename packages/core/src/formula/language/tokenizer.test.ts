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

describe('tokenizeFormula: lossless token stream', () => {
	it.each([
		'=SUM(A1:A3 A2:B2)',
		'=IF( A1 > 0 ,\n  "yes" , "no" )',
		"='It''s'!A1 + 'My Sheet'!B2:C3",
		'={1,-2;"x",#N/A}',
		'=Table1[[#This Row],[Qty]]*2',
		'=_xlfn.LET(_xlpm.x,1,_xlpm.x+1)',
		'="unterminated',
		'=Sheet1!',
		'=SUM(A1',
		'',
		'=',
		'   ',
	])('joins back to %j with contiguous offsets', (formula) => {
		const tokens = tokenizeFormula(formula);
		expect(tokens.map((t) => t.text).join('')).toBe(formula);
		let pos = 0;
		for (const t of tokens) {
			expect(t.start).toBe(pos);
			expect(t.end).toBe(pos + t.text.length);
			pos = t.end;
		}
		expect(pos).toBe(formula.length);
	});

	it('marks only an initial = as the formula start', () => {
		expect(tokenizeFormula('=A1=B1').map((t) => t.type)).toEqual([
			'equals',
			'ref',
			'operator',
			'ref',
		]);
		expect(tokenizeFormula('A1=B1')[0].type).toBe('ref');
	});
});

describe('tokenizeFormula: literals', () => {
	it.each(['1', '1.5', '.5', '1.', '1E+10', '1e-3', '1E5', '12.5e2', '0'])('number %s', (n) => {
		expect(only(n)).toMatchObject({ type: 'number', text: n });
	});

	it('stops a number before an incomplete exponent', () => {
		expect(sig('=1E')).toEqual(['number:1', 'name:E']);
		expect(sig('=1E+')).toEqual(['number:1', 'name:E', 'operator:+']);
	});

	it('reads strings with doubled quotes', () => {
		expect(only('"a""b"')).toMatchObject({ type: 'string', text: '"a""b"' });
		expect(sig('=""&""""')).toEqual(['string:""', 'operator:&', 'string:""""']);
	});

	it('flags unterminated strings', () => {
		const t = only('"abc');
		expect(t.type).toBe('string');
		expect(t.incomplete).toBe(true);
	});

	it('reads booleans case-insensitively, but TRUE() is a function', () => {
		expect(sig('=TRUE+false')).toEqual(['boolean:TRUE', 'operator:+', 'boolean:false']);
		expect(sig('=TRUE()')).toEqual(['function:TRUE', 'call-open:(', 'call-close:)']);
	});

	it.each([
		'#NULL!',
		'#DIV/0!',
		'#VALUE!',
		'#REF!',
		'#NAME?',
		'#NUM!',
		'#N/A',
		'#GETTING_DATA',
		'#SPILL!',
		'#CALC!',
		'#FIELD!',
		'#BLOCKED!',
		'#CONNECT!',
		'#BUSY!',
		'#UNKNOWN!',
	])('error literal %s', (code) => {
		expect(only(code)).toMatchObject({ type: 'error', text: code });
		expect(only(code.toLowerCase())).toMatchObject({ type: 'error' });
	});

	it('keeps dead references written after a sheet deletion in one token', () => {
		expect(sig('=#REF!A1+1')).toEqual(['error:#REF!A1', 'operator:+', 'number:1']);
		expect(sig('=SUM(#REF!$A$1:B2)')).toEqual([
			'function:SUM',
			'call-open:(',
			'error:#REF!$A$1:B2',
			'call-close:)',
		]);
	});
});

describe('tokenizeFormula: references', () => {
	it.each([
		['A1', 'cell', { row: 0, col: 0, rowAbs: false, colAbs: false }],
		['$A$1', 'cell', { row: 0, col: 0, rowAbs: true, colAbs: true }],
		['A$1', 'cell', { row: 0, col: 0, rowAbs: true, colAbs: false }],
		['$B3', 'cell', { row: 2, col: 1, rowAbs: false, colAbs: true }],
		[
			'xfd1048576',
			'cell',
			{ row: MAX_ROW_INDEX, col: MAX_COL_INDEX, rowAbs: false, colAbs: false },
		],
	])('cell %s', (text, kind, corner) => {
		const t = only(text);
		expect(t.type).toBe('ref');
		expect(t.area?.kind).toBe(kind);
		expect(t.area?.start).toEqual(corner);
	});

	it('reads areas, whole columns and whole rows as single tokens', () => {
		expect(only('A1:B2').area).toEqual({
			kind: 'cells',
			start: { row: 0, col: 0, rowAbs: false, colAbs: false },
			end: { row: 1, col: 1, rowAbs: false, colAbs: false },
		});
		expect(only('$A:C').area).toEqual({
			kind: 'columns',
			start: { row: 0, col: 0, rowAbs: false, colAbs: true },
			end: { row: MAX_ROW_INDEX, col: 2, rowAbs: false, colAbs: false },
		});
		expect(only('$1:3').area).toEqual({
			kind: 'rows',
			start: { row: 0, col: 0, rowAbs: true, colAbs: false },
			end: { row: 2, col: MAX_COL_INDEX, rowAbs: false, colAbs: false },
		});
		expect(only('B2:A1').area?.start.row).toBe(1);
	});

	it('treats out-of-grid addresses as names', () => {
		expect(only('XFE1').type).toBe('name');
		expect(only('A1048577').type).toBe('name');
		expect(only('A0').type).toBe('name');
		expect(only('ABCD1').type).toBe('name');
		expect(only('R1C1').type).toBe('name');
	});

	it('does not merge mismatched range sides', () => {
		expect(sig('=A1:B')).toEqual(['ref:A1', 'operator::', 'name:B']);
		expect(sig('=A:B2')).toEqual(['name:A', 'operator::', 'ref:B2']);
		expect(sig('=A1:B2:C3')).toEqual(['ref:A1:B2', 'operator::', 'ref:C3']);
		expect(sig('=A1:INDEX(B:B,5)')).toEqual([
			'ref:A1',
			'operator::',
			'function:INDEX',
			'call-open:(',
			'ref:B:B',
			'arg-sep:,',
			'number:5',
			'call-close:)',
		]);
	});

	it('reads cell-like function names as functions', () => {
		expect(sig('=LOG10(A1)+ATAN2(1,1)+DAYS360(A1,B1)')[0]).toBe('function:LOG10');
		expect(sig('=ATAN2(1,1)')[0]).toBe('function:ATAN2');
	});

	it('reads sheet-qualified references', () => {
		const t = only('Sheet1!$A$1:B2');
		expect(t).toMatchObject({ type: 'ref', sheet: { sheet: 'Sheet1' }, qualifierLength: 7 });
		expect(t.area?.kind).toBe('cells');
		expect(only("'My Sheet'!A1")).toMatchObject({
			sheet: { sheet: 'My Sheet' },
			qualifierLength: 11,
		});
		expect(only("'It''s'!A1").sheet).toEqual({ sheet: "It's" });
		expect(only('Données!A:A').sheet).toEqual({ sheet: 'Données' });
		expect(only('Sheet1!1:1').area?.kind).toBe('rows');
	});

	it('reads 3-D references, including column-like sheet names', () => {
		expect(only('Sheet1:Sheet3!A1').sheet).toEqual({ sheet: 'Sheet1', sheetEnd: 'Sheet3' });
		expect(only('Jan:Dec!B5').sheet).toEqual({ sheet: 'Jan', sheetEnd: 'Dec' });
		expect(only("'Sheet 1:Sheet 3'!A1:B2").sheet).toEqual({
			sheet: 'Sheet 1',
			sheetEnd: 'Sheet 3',
		});
	});

	it('reads external references', () => {
		expect(only('[1]Sheet1!A1').sheet).toEqual({ sheet: 'Sheet1', workbook: '1' });
		expect(only("'[Book.xlsx]Sheet 1'!A1").sheet).toEqual({
			sheet: 'Sheet 1',
			workbook: 'Book.xlsx',
		});
		expect(only("'C:\\Docs\\[Book.xlsx]Data'!A1").sheet).toEqual({
			sheet: 'Data',
			workbook: 'C:\\Docs\\Book.xlsx',
		});
		expect(only('[1]!Total')).toMatchObject({ type: 'name', sheet: { sheet: '', workbook: '1' } });
		expect(only('[Book1.xlsx]Sheet1!$A$1').sheet?.workbook).toBe('Book1.xlsx');
	});

	it('reads #REF! and names after a sheet prefix', () => {
		expect(only('Sheet1!#REF!')).toMatchObject({ type: 'error', sheet: { sheet: 'Sheet1' } });
		expect(only('Sheet1!MyName')).toMatchObject({
			type: 'name',
			text: 'Sheet1!MyName',
			qualifierLength: 7,
		});
	});

	it('flags incomplete sheet prefixes', () => {
		expect(only('Sheet1!')).toMatchObject({ type: 'unknown', incomplete: true });
		expect(only("'My Sheet")).toMatchObject({ type: 'unknown', incomplete: true });
		expect(sig("='abc'")).toEqual(["unknown:'abc'"]);
	});

	it('reads the spill operator only directly after a reference or name', () => {
		expect(sig('=A1#')).toEqual(['ref:A1', 'spill:#']);
		expect(sig('=Sheet1!B2#+1')).toEqual(['ref:Sheet1!B2', 'spill:#', 'operator:+', 'number:1']);
		expect(sig('=Totals#')).toEqual(['name:Totals', 'spill:#']);
		expect(sig('=A1 #N/A')).toEqual(['ref:A1', 'error:#N/A']);
	});
});
