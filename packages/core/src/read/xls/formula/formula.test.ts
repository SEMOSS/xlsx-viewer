import { describe, expect, it } from 'vitest';

import { parseFormula } from '../../../formula/language/parser';
import { arrayExtra, memExtra, ptg, rgce } from '../test-utils/ptg-builder';
import {
	decodeRgce,
	type FormulaEnv,
	type FormulaPos,
	renderFormula,
	XlsExternalReferenceError,
	XlsFormulaError,
} from './index';

const env: FormulaEnv = {
	sheetNames: ['Sheet1', 'My Sheet', 'Data', "Bob's"],
	xti: [
		{ supbook: 0, first: 0, last: 0 },
		{ supbook: 0, first: 1, last: 1 },
		{ supbook: 0, first: 0, last: 2 },
		{ supbook: 0, first: 0xffff, last: 0xffff },
		{ supbook: 1, first: 0xfffe, last: 0xfffe },
		{ supbook: 2, first: 0, last: 0 },
		{ supbook: 0, first: 0xfffe, last: 0xfffe },
		{ supbook: 0, first: 3, last: 3 },
	],
	supbooks: [
		{ kind: 'self', names: [] },
		{ kind: 'addin', names: ['EDATE', '_xlfn.IFERROR', 'MyAddinFn'] },
		{ kind: 'external', path: 'Other.xls', sheets: ['S1'], names: ['ExtName'] },
	],
	names: ['Rate', '_xlnm.Print_Area', '_xlfn.SUMIFS', 'MyUdf'],
};

const at = (row = 0, col = 0, mode: FormulaPos['mode'] = 'cell'): FormulaPos => ({
	row,
	col,
	mode,
});

function text(bytes: Uint8Array, extra?: number[], pos = at()): string {
	return renderFormula(decodeRgce(bytes, extra ? Uint8Array.from(extra) : undefined), env, pos);
}

const A1 = ptg.ref(0, 0);
const B1 = ptg.ref(0, 1);

describe('constants and operators', () => {
	it.each<[string, number[][], string]>([
		['integer', [ptg.int(1)], '1'],
		['number', [ptg.num(3.5)], '3.5'],
		['large number', [ptg.num(1e21)], '1E+21'],
		['small number', [ptg.num(0.1), ptg.num(0.2), ptg.add], '0.1+0.2'],
		['string with quotes', [ptg.str('He said "hi"')], '"He said ""hi"""'],
		['16-bit string', [ptg.str('Grüße €', true)], '"Grüße €"'],
		['boolean', [ptg.bool(true)], 'TRUE'],
		['error', [ptg.err(0x07)], '#DIV/0!'],
		['precedence', [ptg.int(1), ptg.int(2), ptg.int(3), ptg.mul, ptg.add], '1+2*3'],
		[
			'explicit parens',
			[ptg.int(1), ptg.int(2), ptg.add, ptg.paren, ptg.int(3), ptg.mul],
			'(1+2)*3',
		],
		['implied parens', [ptg.int(1), ptg.int(2), ptg.add, ptg.int(3), ptg.mul], '(1+2)*3'],
		['right-nested minus', [ptg.int(1), ptg.int(2), ptg.int(3), ptg.sub, ptg.sub], '1-(2-3)'],
		[
			'power is left-associative',
			[ptg.int(2), ptg.int(3), ptg.power, ptg.int(2), ptg.power],
			'2^3^2',
		],
		['unary minus', [A1, ptg.uminus], '-A1'],
		['unary plus', [A1, ptg.uplus], '+A1'],
		['percent', [ptg.int(50), ptg.percent], '50%'],
		['negation before power', [ptg.int(2), ptg.uminus, ptg.int(2), ptg.power], '-2^2'],
		['negated power', [ptg.int(2), ptg.int(2), ptg.power, ptg.paren, ptg.uminus], '-(2^2)'],
		['concatenation', [ptg.str('a'), ptg.str('b'), ptg.concat], '"a"&"b"'],
		['less than', [A1, B1, ptg.lt], 'A1<B1'],
		['less or equal', [A1, B1, ptg.le], 'A1<=B1'],
		['equal', [A1, B1, ptg.eq], 'A1=B1'],
		['greater or equal', [A1, B1, ptg.ge], 'A1>=B1'],
		['greater than', [A1, B1, ptg.gt], 'A1>B1'],
		['not equal', [A1, B1, ptg.ne], 'A1<>B1'],
		['spaces are dropped', [A1, ptg.attrSpace(2), ptg.int(1), ptg.add], 'A1+1'],
	])('%s', (_label, tokens, expected) => {
		expect(text(rgce(...tokens))).toBe(expected);
	});

	it('maps every BIFF error code', () => {
		const codes: Array<[number, string]> = [
			[0x00, '#NULL!'],
			[0x07, '#DIV/0!'],
			[0x0f, '#VALUE!'],
			[0x17, '#REF!'],
			[0x1d, '#NAME?'],
			[0x24, '#NUM!'],
			[0x2a, '#N/A'],
		];
		for (const [code, literal] of codes) {
			expect(text(rgce(ptg.err(code)))).toBe(literal);
		}
	});
});

describe('references', () => {
	it.each<[string, number[][], string]>([
		['absolute cell', [ptg.ref(0, 0, { rowAbs: true, colAbs: true })], '$A$1'],
		[
			'mixed cells',
			[ptg.ref(4, 2, { colAbs: true }), ptg.ref(4, 2, { rowAbs: true }), ptg.add],
			'$C5+C$5',
		],
		['area', [ptg.area(0, 0, 1, 1)], 'A1:B2'],
		['absolute area', [ptg.area(0, 0, 9, 1, { rowAbs: true, colAbs: true })], '$A$1:$B$10'],
		['whole column', [ptg.area(0, 0, 65535, 0)], 'A:A'],
		['absolute columns', [ptg.area(0, 1, 65535, 2, { rowAbs: true, colAbs: true })], '$B:$C'],
		['whole row', [ptg.area(0, 0, 0, 255)], '1:1'],
		['absolute rows', [ptg.area(1, 0, 2, 255, { rowAbs: true, colAbs: true })], '$2:$3'],
		['whole sheet', [ptg.area(0, 0, 65535, 255)], '1:65536'],
		['3-D cell', [ptg.ref3d(0, 0, 0)], 'Sheet1!A1'],
		['quoted sheet', [ptg.ref3d(1, 1, 1, { rowAbs: true, colAbs: true })], "'My Sheet'!$B$2"],
		['sheet with apostrophe', [ptg.ref3d(7, 0, 0)], "'Bob''s'!A1"],
		['sheet range', [ptg.area3d(2, 0, 0, 1, 1)], 'Sheet1:Data!A1:B2'],
		['3-D whole column', [ptg.area3d(0, 0, 3, 65535, 3)], 'Sheet1!D:D'],
		['deleted sheet', [ptg.ref3d(3, 0, 0)], '#REF!'],
		['ref error', [ptg.refErr()], '#REF!'],
		['area error', [ptg.areaErr()], '#REF!'],
		['3-D ref error', [ptg.refErr3d(0)], 'Sheet1!#REF!'],
		['3-D area error', [ptg.areaErr3d(1)], "'My Sheet'!#REF!"],
		['last BIFF cell', [ptg.ref(65535, 255, { rowAbs: true, colAbs: true })], '$IV$65536'],
	])('%s', (_label, tokens, expected) => {
		expect(text(rgce(...tokens))).toBe(expected);
	});

	it('intersects, unions and ranges references', () => {
		expect(text(rgce(ptg.area(0, 0, 1, 1), ptg.area(0, 1, 2, 2), ptg.isect))).toBe('A1:B2 B1:C3');
		expect(text(rgce(A1, B1, ptg.union, ptg.paren))).toBe('(A1,B1)');
		expect(text(rgce(A1, B1, ptg.union, ptg.ref(0, 2), ptg.union, ptg.paren))).toBe('(A1,B1,C1)');
		expect(text(rgce(A1, ptg.ref(2, 1), ptg.range))).toBe('A1:B3');
	});

	it('resolves ptgRefN / ptgAreaN offsets against the base cell', () => {
		expect(text(rgce(ptg.refN(-1, -2)), undefined, at(5, 3, 'shared'))).toBe('B5');
		expect(text(rgce(ptg.refN(0, 0, { colAbs: true })), undefined, at(5, 3, 'shared'))).toBe('$A6');
		expect(text(rgce(ptg.areaN(-2, 0, 0, 0)), undefined, at(10, 2, 'shared'))).toBe('C9:C11');
		expect(
			text(
				rgce(ptg.areaN(0, 0, 0, 1, { rowAbs: true, colAbs: false })),
				undefined,
				at(7, 3, 'shared'),
			),
		).toBe('D$1:E$1');
	});

	it('treats relative 3-D references in shared formulas as offsets', () => {
		expect(text(rgce(ptg.ref3d(0, 1, 0)), undefined, at(3, 1, 'shared'))).toBe('Sheet1!B5');
		// Plain ptgRef in a shared formula holds absolute coordinates.
		expect(text(rgce(ptg.ref(1, 0)), undefined, at(3, 1, 'shared'))).toBe('A2');
	});

	it('wraps negative offsets in names around the SpreadsheetML grid', () => {
		expect(text(rgce(ptg.ref3d(0, 0, -1, { rowAbs: true })), undefined, at(0, 0, 'name'))).toBe(
			'Sheet1!XFD$1',
		);
		expect(text(rgce(ptg.ref3d(0, -1, 0, { colAbs: true })), undefined, at(0, 0, 'name'))).toBe(
			'Sheet1!$A1048576',
		);
		expect(
			text(
				rgce(ptg.area3d(0, 0, 0, 9, 1, { rowAbs: true, colAbs: true })),
				undefined,
				at(0, 0, 'name'),
			),
		).toBe('Sheet1!$A$1:$B$10');
	});

	it('uses the top-left cell as the base of rule formulas', () => {
		expect(text(rgce(ptg.refN(0, -1)), undefined, at(1, 1, 'rule'))).toBe('A2');
		expect(text(rgce(ptg.ref(0, 0), ptg.int(5), ptg.gt), undefined, at(1, 1, 'rule'))).toBe('A1>5');
	});
});

describe('functions', () => {
	it.each<[string, number[][], string]>([
		['SUM shortcut', [ptg.area(0, 0, 2, 0), ptg.attrSum], 'SUM(A1:A3)'],
		['SUM with arguments', [A1, B1, ptg.int(3), ptg.funcVar(4, 3)], 'SUM(A1,B1,3)'],
		['fixed ROUND', [A1, ptg.int(2), ptg.func(27)], 'ROUND(A1,2)'],
		['no arguments', [ptg.func(74)], 'NOW()'],
		['PI', [ptg.func(19), ptg.int(2), ptg.mul], 'PI()*2'],
		['TRUE()', [ptg.func(34)], 'TRUE()'],
		['volatile marker', [ptg.attrSemi, ptg.func(63)], 'RAND()'],
		[
			'IF with jumps',
			[
				A1,
				ptg.attrIf(7),
				ptg.int(1),
				ptg.attrGoto(9),
				ptg.int(2),
				ptg.attrGoto(3),
				ptg.funcVar(1, 3),
			],
			'IF(A1,1,2)',
		],
		[
			'CHOOSE with jump table',
			[
				ptg.int(2),
				ptg.attrChoose([4, 10, 16]),
				ptg.str('a'),
				ptg.attrGoto(6),
				ptg.str('b'),
				ptg.attrGoto(0),
				ptg.funcVar(100, 3),
			],
			'CHOOSE(2,"a","b")',
		],
		['missing argument', [A1, ptg.miss, ptg.int(1), ptg.funcVar(1, 3)], 'IF(A1,,1)'],
		['trailing missing argument', [ptg.int(1), ptg.miss, ptg.funcVar(4, 2)], 'SUM(1,)'],
		[
			'nested calls',
			[ptg.area(0, 0, 2, 0), ptg.attrSum, ptg.area(0, 0, 2, 0), ptg.funcVar(0, 1), ptg.div],
			'SUM(A1:A3)/COUNT(A1:A3)',
		],
		[
			'INDEX',
			[ptg.area(0, 0, 2, 2), ptg.int(2), ptg.int(2), ptg.funcVar(29, 3)],
			'INDEX(A1:C3,2,2)',
		],
		[
			'VLOOKUP',
			[
				ptg.str('x'),
				ptg.area(0, 0, 9, 1, { rowAbs: true, colAbs: true }),
				ptg.int(2),
				ptg.bool(false),
				ptg.funcVar(102, 4),
			],
			'VLOOKUP("x",$A$1:$B$10,2,FALSE)',
		],
		['union argument', [A1, B1, ptg.union, ptg.paren, ptg.funcVar(4, 1)], 'SUM((A1,B1))'],
		['union argument without paren', [A1, B1, ptg.union, ptg.funcVar(4, 1)], 'SUM((A1,B1))'],
		['parenthesised call', [A1, ptg.attrSum, ptg.paren, ptg.int(2), ptg.mul], '(SUM(A1))*2'],
		['dotted name', [A1, ptg.func(261)], 'ERROR.TYPE(A1)'],
		['Excel 2007 function', [A1, ptg.int(0), ptg.func(480)], 'IFERROR(A1,0)'],
		['macro command', [ptg.funcVar(8, 0, { command: true })], 'PRINT()'],
	])('%s', (_label, tokens, expected) => {
		expect(text(rgce(...tokens))).toBe(expected);
	});

	it('names user-defined and add-in functions from their first operand', () => {
		expect(text(rgce(ptg.nameX(4, 1), A1, ptg.int(1), ptg.funcVar(255, 3)))).toBe('EDATE(A1,1)');
		expect(text(rgce(ptg.nameX(4, 2), A1, ptg.int(0), ptg.funcVar(255, 3)))).toBe('IFERROR(A1,0)');
		expect(
			text(
				rgce(
					ptg.name(3),
					ptg.area(0, 0, 2, 0),
					ptg.area(0, 1, 2, 1),
					ptg.str('>1'),
					ptg.funcVar(255, 4),
				),
			),
		).toBe('SUMIFS(A1:A3,B1:B3,">1")');
		expect(text(rgce(ptg.name(4), ptg.int(1), ptg.funcVar(255, 2)))).toBe('MyUdf(1)');
		expect(text(rgce(ptg.nameX(4, 3), ptg.funcVar(255, 1)))).toBe('MyAddinFn()');
	});

	it('reads ranges built with ptgMemFunc / ptgMemArea wrappers', () => {
		const index = [ptg.area(0, 1, 65535, 1), ptg.int(2), ptg.funcVar(29, 2)];
		const body = [A1, ...index, ptg.range];
		const cce = body.flat().length;
		expect(text(rgce(ptg.memFunc(cce), ...body))).toBe('A1:INDEX(B:B,2)');
		// The cached areas of ptgMemArea precede the array data in the extra bytes.
		const isect = [ptg.area(0, 0, 1, 1), ptg.area(0, 1, 2, 2), ptg.isect];
		const tokens = rgce(
			ptg.memArea(isect.flat().length),
			...isect,
			ptg.attrSum,
			ptg.array(),
			ptg.funcVar(4, 2),
		);
		const extra = [...memExtra([[0, 1, 1, 1]]), ...arrayExtra([[1, 2]])];
		expect(text(tokens, extra)).toBe('SUM(SUM(A1:B2 B1:C3),{1,2})');
		expect(text(rgce(ptg.memErr(1), ptg.refErr()))).toBe('#REF!');
		expect(text(rgce(ptg.memNoMem(1), ptg.refErr()))).toBe('#REF!');
	});

	it('reads array constants', () => {
		expect(
			text(
				rgce(ptg.array()),
				arrayExtra([
					[1, 2],
					[3, 4],
				]),
			),
		).toBe('{1,2;3,4}');
		expect(
			text(
				rgce(ptg.array()),
				arrayExtra([
					[1, 'a'],
					[true, { error: 0x2a }],
				]),
			),
		).toBe('{1,"a";TRUE,#N/A}');
		expect(text(rgce(ptg.array()), arrayExtra([[-1, 2.5]]))).toBe('{-1,2.5}');
		expect(
			text(rgce(ptg.array(), ptg.array(), ptg.funcVar(4, 2)), [
				...arrayExtra([[1]]),
				...arrayExtra([['x']]),
			]),
		).toBe('SUM({1},{"x"})');
	});
});

describe('names', () => {
	it('renders defined names and qualified names', () => {
		expect(text(rgce(ptg.name(1), ptg.int(2), ptg.mul))).toBe('Rate*2');
		expect(text(rgce(ptg.name(2)))).toBe('_xlnm.Print_Area');
		expect(text(rgce(ptg.name(99)))).toBe('#NAME?');
		expect(text(rgce(ptg.nameX(1, 1)))).toBe("'My Sheet'!Rate");
		expect(text(rgce(ptg.nameX(6, 1)))).toBe('Rate');
	});
});

describe('errors', () => {
	it('reports external references', () => {
		expect(() => text(rgce(ptg.ref3d(5, 0, 0)))).toThrow(XlsExternalReferenceError);
		expect(() => text(rgce(ptg.nameX(5, 1)))).toThrow(XlsExternalReferenceError);
	});

	it('rejects unknown, truncated and unbalanced token streams', () => {
		expect(() => text(rgce(ptg.funcVar(202, 0)))).toThrow(XlsFormulaError);
		expect(() => text(rgce([0x24, 0x00]))).toThrow(/truncated/);
		expect(() => text(rgce([0x80]))).toThrow(XlsFormulaError);
		expect(() => text(rgce(A1, B1))).toThrow(/unbalanced/);
		expect(() => text(rgce(A1, ptg.funcVar(4, 3)))).toThrow(/missing/);
		expect(() => text(rgce(ptg.add))).toThrow(XlsFormulaError);
		expect(() => text(rgce(ptg.array()), [1, 0, 0])).toThrow(XlsFormulaError);
		expect(() => text(rgce(ptg.exp(0, 0)))).toThrow(XlsFormulaError);
		expect(() => text(rgce([0x18, 0x01]))).toThrow(/Extended/);
	});

	it('produces formulas the parser accepts', () => {
		const samples = [
			rgce(ptg.ref3d(2, 0, 0), ptg.ref3d(1, 0, 0), ptg.add),
			rgce(ptg.area(0, 0, 65535, 255), ptg.funcVar(0, 1)),
			rgce(ptg.str(''), ptg.str('a"b'), ptg.concat, ptg.paren, ptg.percent),
		];
		for (const s of samples) {
			expect(() => parseFormula(text(s))).not.toThrow();
		}
	});
});
