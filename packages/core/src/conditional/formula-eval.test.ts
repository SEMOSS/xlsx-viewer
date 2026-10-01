import { describe, expect, it } from 'vitest';
import type { FormulaValue } from '../formula/types';
import { makeArray } from '../formula/types';
import { createWorkbook } from '../sheet/workbook';
import { ERRORS } from '../values/errors';
import {
	formulaResultToScalar,
	isConditionTrue,
	parseLiteralFormula,
	RelativeFormula,
	scalarToNumber,
	stripLeadingEquals,
	workbookCellReader,
	type FormulaEvaluator,
} from './formula-eval';
import { setCells } from './test-utils/fixtures';

describe('parseLiteralFormula', () => {
	it('recognises constants', () => {
		expect(parseLiteralFormula('5')).toBe(5);
		expect(parseLiteralFormula(' -1.5 ')).toBe(-1.5);
		expect(parseLiteralFormula('2E3')).toBe(2000);
		expect(parseLiteralFormula('"abc"')).toBe('abc');
		expect(parseLiteralFormula('"say ""hi"""')).toBe('say "hi"');
		expect(parseLiteralFormula('""')).toBe('');
		expect(parseLiteralFormula('true')).toBe(true);
		expect(parseLiteralFormula('FALSE')).toBe(false);
		expect(parseLiteralFormula('#N/A')).toBe(ERRORS.NA);
	});

	it('rejects expressions', () => {
		expect(parseLiteralFormula('A1')).toBeUndefined();
		expect(parseLiteralFormula('"a"&"b"')).toBeUndefined();
		expect(parseLiteralFormula('1+2')).toBeUndefined();
		expect(parseLiteralFormula('')).toBeUndefined();
	});

	it('strips a leading equals sign', () => {
		expect(stripLeadingEquals('=A1>5')).toBe('A1>5');
		expect(stripLeadingEquals('A1')).toBe('A1');
	});
});

describe('formulaResultToScalar', () => {
	const wb = createWorkbook({ sheetNames: ['S1', 'S2'] });
	setCells(wb.sheets[1], { B2: 'two' });
	const read = workbookCellReader(wb);

	it('unwraps references, arrays and errors', () => {
		expect(formulaResultToScalar(7)).toBe(7);
		expect(formulaResultToScalar(undefined)).toBeNull();
		expect(formulaResultToScalar(Number.POSITIVE_INFINITY)).toBe(ERRORS.NUM);
		expect(formulaResultToScalar(makeArray([[3, 4]]))).toBe(3);
		const ref: FormulaValue = {
			kind: 'ref',
			sheet: 1,
			range: { top: 1, left: 1, bottom: 5, right: 5 },
		};
		expect(formulaResultToScalar(ref, read)).toBe('two');
		expect(formulaResultToScalar(ref)).toBe(ERRORS.VALUE);
		expect(formulaResultToScalar({ kind: 'refs', refs: [ref as never] }, read)).toBe('two');
		expect(formulaResultToScalar(ERRORS.DIV0)).toBe(ERRORS.DIV0);
	});

	it('applies condition semantics', () => {
		expect(isConditionTrue(true)).toBe(true);
		expect(isConditionTrue(2)).toBe(true);
		expect(isConditionTrue(-0.5)).toBe(true);
		expect(isConditionTrue(0)).toBe(false);
		expect(isConditionTrue(false)).toBe(false);
		expect(isConditionTrue('TRUE')).toBe(false);
		expect(isConditionTrue(null)).toBe(false);
		expect(isConditionTrue(ERRORS.NA)).toBe(false);
	});

	it('converts threshold results to numbers', () => {
		expect(scalarToNumber(4)).toBe(4);
		expect(scalarToNumber(null)).toBe(0);
		expect(scalarToNumber('25%')).toBe(0.25);
		expect(scalarToNumber('abc')).toBeUndefined();
		expect(scalarToNumber(true)).toBeUndefined();
		expect(scalarToNumber(ERRORS.NA)).toBeUndefined();
	});
});

describe('RelativeFormula', () => {
	it('translates relative references by the offset from the anchor', () => {
		const f = new RelativeFormula('$B2>C$1', 1, 0);
		expect(f.textAt(1, 0)).toBe('$B2>C$1');
		expect(f.textAt(3, 0)).toBe('$B4>C$1');
		expect(f.textAt(1, 2)).toBe('$B2>E$1');
	});

	it('does not translate absolute or constant formulas', () => {
		expect(new RelativeFormula('$A$1*2', 0, 0).textAt(10, 10)).toBe('$A$1*2');
		const literal = new RelativeFormula('=10', 0, 0);
		expect(literal.literal).toBe(10);
		let calls = 0;
		const evaluate: FormulaEvaluator = () => {
			calls++;
			return 1;
		};
		expect(literal.evaluate(evaluate, 0, 5, 5)).toBe(10);
		expect(calls).toBe(0);
	});

	it('evaluates at the target cell and turns exceptions into #VALUE!', () => {
		const seen: Array<[string, number, number]> = [];
		const evaluate: FormulaEvaluator = (formula, _s, row, col) => {
			seen.push([formula, row, col]);
			return row;
		};
		const f = new RelativeFormula('A1>0', 0, 0);
		expect(f.evaluate(evaluate, 0, 4, 1)).toBe(4);
		expect(seen).toEqual([['B5>0', 4, 1]]);
		const throwing = new RelativeFormula('A1', 0, 0);
		expect(
			throwing.evaluate(
				() => {
					throw new Error('x');
				},
				0,
				0,
				0,
			),
		).toBe(ERRORS.VALUE);
	});
});
