import { describe, expect, it } from 'vitest';

import { serializeFormula } from '../formula/language/serializer';
import { parseFormula } from '../formula/language/parser';
import { analyzeFormula } from './analysis';

function deps(formula: string) {
	const a = analyzeFormula(parseFormula(formula));
	return {
		cells: a.cells.map((c) => serializeFormula(c)),
		ranges: a.ranges.map((r) => serializeFormula(r)),
		names: a.names.map((n) => n.name),
		functions: [...a.functions].sort(),
		volatile: a.volatile,
		external: a.external,
		subtotal: a.subtotal,
		structured: a.structured.map((s) => s.text),
	};
}

describe('formula analysis', () => {
	it('collects cells, ranges, names, tables and functions', () => {
		const d = deps('SUM(A1:B2,Sheet2!C3)+Rate*Table1[Qty]+MAX(D:D)');
		expect(d.cells).toEqual(['Sheet2!C3']);
		expect(d.ranges).toEqual(['A1:B2', 'D:D']);
		expect(d.names).toEqual(['Rate']);
		expect(d.structured).toEqual(['Table1[Qty]']);
		expect(d.functions).toEqual(['MAX', 'SUM']);
		expect(d.volatile).toBe(false);
	});

	it('does not count references only inspected for their position', () => {
		expect(deps('ROW(A1)+COLUMNS(B1:C9)+ROWS(Data)').cells).toEqual([]);
		expect(deps('ROWS(A1:A9)').ranges).toEqual([]);
		expect(deps('ROWS(FILTER(A1:A9,B1:B9))').ranges).toEqual(['A1:A9', 'B1:B9']);
		expect(deps('OFFSET(A1,B1,0)').cells).toEqual(['B1']);
		expect(deps('CELL("address",A1)').cells).toEqual([]);
		expect(deps('CELL("contents",A1)').cells).toEqual(['A1']);
		expect(deps('ROWS(A1#)').cells).toEqual(['A1']);
	});

	it('scopes LET and LAMBDA names', () => {
		const d = deps('LET(x,A1,y,x*2,x+y+z)');
		expect(d.names).toEqual(['z']);
		expect(d.cells).toEqual(['A1']);
		expect(deps('LAMBDA(a,[b],a+b+c)(1)').names).toEqual(['c']);
		expect(deps('LET(f,LAMBDA(v,v+1),f(B1))').functions).toEqual(['LAMBDA', 'LET']);
	});

	it('flags volatile functions, external references and subtotals', () => {
		expect(deps('NOW()+1').volatile).toBe(true);
		expect(deps('INDIRECT("A1")').volatile).toBe(true);
		expect(deps('[1]Sheet1!A1+1').external).toBe(true);
		expect(deps('SUBTOTAL(9,A1:A3)*2').subtotal).toBe(true);
		expect(deps('AGGREGATE(9,0,A1:A3)').subtotal).toBe(true);
	});
});
