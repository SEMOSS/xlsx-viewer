import { describe, expect, it } from 'vitest';

import { MAX_ROW_INDEX } from '../../address/constants';
import {
	getActiveFunctionCall,
	getCompletionPrefix,
	isPointModeAllowed,
	toggleReferenceAbsolute,
} from './editor-caret';
import { getFormulaReferences, tokenizeForDisplay } from './editor-display';

/** Split `text` at the `|` marker into the formula and the caret offset. */
function at(text: string): [string, number] {
	const caret = text.indexOf('|');
	return [text.slice(0, caret) + text.slice(caret + 1), caret];
}

describe('tokenizeForDisplay', () => {
	it('classifies tokens for highlighting', () => {
		const kinds = tokenizeForDisplay('=SUM(A1:B2, "x") + Rate%').map((t) => `${t.kind}:${t.text}`);
		expect(kinds).toEqual([
			'operator:=',
			'function:SUM',
			'paren:(',
			'reference:A1:B2',
			'separator:,',
			'whitespace: ',
			'string:"x"',
			'paren:)',
			'whitespace: ',
			'operator:+',
			'whitespace: ',
			'name:Rate',
			'operator:%',
		]);
	});

	it('covers every character, even in partial input', () => {
		for (const f of ['=IF(A1>0,"ye', '=Sheet1!', "='My", '=Table1[Co', '={1,2;3']) {
			expect(
				tokenizeForDisplay(f)
					.map((t) => t.text)
					.join(''),
			).toBe(f);
		}
		expect(tokenizeForDisplay('=1~2')[2]).toMatchObject({ kind: 'invalid', text: '~' });
	});

	it('numbers bracket depth for matching', () => {
		const parens = tokenizeForDisplay('=IF((A1),{1},2)').filter((t) => t.depth !== undefined);
		expect(parens.map((t) => `${t.text}${t.depth}`)).toEqual(['(0', '(1', ')1', '{1', '}1', ')0']);
	});

	it('colors references like getFormulaReferences', () => {
		const tokens = tokenizeForDisplay('=A1+B2+a1+Sheet2!A1', 'Sheet1').filter(
			(t) => t.kind === 'reference',
		);
		expect(tokens.map((t) => t.colorIndex)).toEqual([0, 1, 0, 2]);
	});
});

describe('getFormulaReferences', () => {
	it('lists references with ranges and stable color slots', () => {
		const refs = getFormulaReferences('=A1+B2:C3*A1+Sheet2!A1+$A$1+sheet1!A1', 'Sheet1');
		expect(refs.map((r) => [r.text, r.sheet, r.colorIndex])).toEqual([
			['A1', 'Sheet1', 0],
			['B2:C3', 'Sheet1', 1],
			['A1', 'Sheet1', 0],
			['Sheet2!A1', 'Sheet2', 2],
			['$A$1', 'Sheet1', 0],
			['sheet1!A1', 'sheet1', 0],
		]);
		expect(refs[1]).toMatchObject({
			start: 4,
			end: 9,
			range: { top: 1, left: 1, bottom: 2, right: 2 },
		});
	});

	it('resolves whole columns / rows and flags spill references', () => {
		const [col, row, spill] = getFormulaReferences('C:A+2:2+D4#');
		expect(col.range).toEqual({ top: 0, left: 0, bottom: MAX_ROW_INDEX, right: 2 });
		expect(col.sheet).toBeUndefined();
		expect(row.range.top).toBe(1);
		expect(spill).toMatchObject({ text: 'D4', spill: true });
	});

	it('skips external, 3-D and dead references, names and tables', () => {
		const refs = getFormulaReferences(
			'=[1]Sheet1!A1+Sheet1:Sheet3!A1+#REF!+Rate+Table1[A]+"B2"+C3',
		);
		expect(refs.map((r) => r.text)).toEqual(['C3']);
	});
});

describe('getActiveFunctionCall', () => {
	it.each([
		['=SUM(|', 'SUM', 0],
		['=SUM(A1, |', 'SUM', 1],
		['=SUM(A1,B1|)', 'SUM', 1],
		['=IF(A1>0, SUM(B1,|B2), 0)', 'SUM', 1],
		['=IF(SUM(B1,B2),|', 'IF', 1],
		['=IF(A1,B1,C1|)', 'IF', 2],
		['=SUM((A1,|B1))', 'SUM', 0],
		['=SUM({1,2|})', 'SUM', 0],
		['=CONCAT("a,|b")', 'CONCAT', 0],
		['=_xlfn.XLOOKUP(A1,|', 'XLOOKUP', 1],
		['=vlookup(A1,B:C,|', 'VLOOKUP', 2],
	])('%s -> %s #%i', (text, name, argIndex) => {
		const [formula, caret] = at(text);
		expect(getActiveFunctionCall(formula, caret)).toMatchObject({ name, argIndex });
	});

	it('reports where the function name starts', () => {
		const [formula, caret] = at('=1+ROUND(A1,|');
		expect(getActiveFunctionCall(formula, caret)?.nameStart).toBe(3);
	});

	it.each(['=SUM(A1)|', '=SU|M(A1)', '=|SUM(A1)', '=A1+|', '=LAMBDA(x,x)(|'])(
		'%s -> null',
		(text) => {
			const [formula, caret] = at(text);
			expect(getActiveFunctionCall(formula, caret)).toBeNull();
		},
	);
});

describe('getCompletionPrefix', () => {
	it.each([
		['=SU|', 'SU', 1],
		['=1+VLO|', 'VLO', 3],
		['=SUM(A1,ro|', 'ro', 8],
		['=A1+B|', 'B', 4],
		['=LOG1|', 'LOG1', 1],
		['=IF(A1,_xl|', '_xl', 7],
		['=-NORM.S|', 'NORM.S', 2],
		['=@IN|', 'IN', 2],
		['=Tot|', 'Tot', 1],
	])('%s -> %s', (text, prefix, start) => {
		const [formula, caret] = at(text);
		expect(getCompletionPrefix(formula, caret)).toEqual({ prefix, start });
	});

	it.each([
		'="SU|',
		'=Sheet1!A|',
		'=12|',
		'=1.5E|',
		'=Table1[Co|',
		'=$A|',
		"='My SU|",
		'=A1|B',
		'=SU|M(1)',
		'=SUM(A1)SU|',
		'|',
	])('%s -> null', (text) => {
		const [formula, caret] = at(text);
		expect(getCompletionPrefix(formula, caret)).toBeNull();
	});
});

describe('isPointModeAllowed', () => {
	it.each([
		['=|', true],
		['=A1+|', true],
		['=SUM(|', true],
		['=SUM(A1,|', true],
		['=SUM(A1, |', true],
		['=A1:|', true],
		['=-|', true],
		['=A1&|', true],
		['=(A1,|', true],
		['=A1>=|', true],
		['=A1|', false],
		['="ab|', false],
		['="a|b"', false],
		['=SUM(A1)|', false],
		['={1,|', false],
		['=A1%|', false],
		['=SU|M', false],
		['|=A1', false],
	])('%s -> %s', (text, expected) => {
		const [formula, caret] = at(text);
		expect(isPointModeAllowed(formula, caret)).toBe(expected);
	});
});

describe('toggleReferenceAbsolute', () => {
	const f4 = (text: string) => {
		const [formula, caret] = at(text);
		const r = toggleReferenceAbsolute(formula, caret);
		return `${r.formula.slice(0, r.caret)}|${r.formula.slice(r.caret)}`;
	};

	it('cycles A1 -> $A$1 -> A$1 -> $A1 -> A1', () => {
		expect(f4('=A1|')).toBe('=$A$1|');
		expect(f4('=$A$1|')).toBe('=A$1|');
		expect(f4('=A$1|')).toBe('=$A1|');
		expect(f4('=$A1|')).toBe('=A1|');
	});

	it('toggles the reference around the caret and moves the caret after it', () => {
		expect(f4('=SUM(A|1,B2)')).toBe('=SUM($A$1|,B2)');
		expect(f4('=SUM(A1,|B2)')).toBe('=SUM(A1,$B$2|)');
		expect(f4('=SUM(A1|,B2)')).toBe('=SUM($A$1|,B2)');
		expect(f4('=x+|A1')).toBe('=x+$A$1|');
	});

	it('toggles ranges as a whole, following the first corner', () => {
		expect(f4('=SUM(A1:B2|)')).toBe('=SUM($A$1:$B$2|)');
		expect(f4('=SUM($A$1:B2|)')).toBe('=SUM(A$1:B$2|)');
		expect(f4('=Sheet1!A1:B|2')).toBe('=Sheet1!$A$1:$B$2|');
		expect(f4("='My Sheet'!|C3")).toBe("='My Sheet'!$C$3|");
	});

	it('toggles whole columns and rows between relative and absolute', () => {
		expect(f4('=SUM(A:C|)')).toBe('=SUM($A:$C|)');
		expect(f4('=SUM($A:$C|)')).toBe('=SUM(A:C|)');
		expect(f4('=SUM(2:5|)')).toBe('=SUM($2:$5|)');
	});

	it('leaves other text alone', () => {
		expect(f4('=SUM(|1)')).toBe('=SUM(|1)');
		expect(f4('="A1|"')).toBe('="A1|"');
		expect(f4('=Rate|')).toBe('=Rate|');
		expect(f4('=A1+|')).toBe('=A1+|');
	});
});
