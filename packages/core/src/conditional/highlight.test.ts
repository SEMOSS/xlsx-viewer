import { describe, expect, it } from 'vitest';
import type { CfOperator } from '../types/conditional-format';
import type { DifferentialFormat } from '../types/style';
import { dateToSerial } from '../values/dates';
import { ERRORS } from '../values/errors';
import { addRule, cfFixture, setColumn } from './test-utils/fixtures';

const RED: DifferentialFormat = { font: { color: { rgb: 'FFFF0000' } } };

/** Addresses in `addrs` that the fixture highlights. */
function highlighted(fx: ReturnType<typeof cfFixture>, addrs: string[]): string[] {
	return addrs.filter((a) => fx.at(a)?.style !== undefined);
}

describe('cellIs rules', () => {
	const cells = { A1: 1, A2: 5, A3: 10, A4: 'text', A5: true, A6: ERRORS.DIV0 };
	const addrs = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7'];

	const cases: Array<[CfOperator, string[], string[]]> = [
		['greaterThan', ['5'], ['A3', 'A4', 'A5']],
		['greaterThanOrEqual', ['5'], ['A2', 'A3', 'A4', 'A5']],
		['lessThan', ['5'], ['A1', 'A7']],
		['lessThanOrEqual', ['5'], ['A1', 'A2', 'A7']],
		['equal', ['5'], ['A2']],
		['notEqual', ['5'], ['A1', 'A3', 'A4', 'A5', 'A7']],
		['between', ['2', '10'], ['A2', 'A3']],
		['between', ['10', '2'], ['A2', 'A3']],
		['notBetween', ['2', '10'], ['A1', 'A4', 'A5', 'A7']],
		['equal', ['"TEXT"'], ['A4']],
		['equal', ['0'], ['A7']],
	];

	it.each(cases)('%s %j', (operator, formulas, expected) => {
		const fx = cfFixture(cells);
		addRule(fx.sheet, 'A1:A7', { type: 'cellIs', operator, formulas, dxf: RED });
		expect(highlighted(fx, addrs)).toEqual(expected);
	});

	it('evaluates operand formulas relative to each cell', () => {
		const fx = cfFixture({ A1: 5, A2: 5, A3: 5, B1: 1, B2: 9, B3: 5 });
		addRule(fx.sheet, 'A1:A3', {
			type: 'cellIs',
			operator: 'greaterThan',
			formulas: ['B1'],
			dxf: RED,
		});
		expect(highlighted(fx, ['A1', 'A2', 'A3'])).toEqual(['A1']);
	});

	it('keeps absolute operands fixed and skips the evaluator for literals', () => {
		const fx = cfFixture({ A1: 5, A2: 15, D1: 10 });
		addRule(fx.sheet, 'A1:A2', {
			type: 'cellIs',
			operator: 'greaterThan',
			formulas: ['$D$1'],
			dxf: RED,
		});
		addRule(fx.sheet, 'C1:C2', { type: 'cellIs', operator: 'lessThan', formulas: ['3'], dxf: RED });
		expect(highlighted(fx, ['A1', 'A2'])).toEqual(['A2']);
		const before = fx.evaluator.calls;
		fx.at('C1');
		fx.at('C2');
		expect(fx.evaluator.calls).toBe(before);
		expect(fx.at('C1')?.style).toBe(RED);
	});

	it('treats error operands as non-matching', () => {
		const fx = cfFixture({ A1: 5 });
		addRule(fx.sheet, 'A1', { type: 'cellIs', operator: 'notEqual', formulas: ['NA()'], dxf: RED });
		expect(fx.at('A1')).toBeUndefined();
	});
});

describe('expression rules', () => {
	it('applies relative formulas per cell (row banding)', () => {
		const fx = cfFixture();
		addRule(fx.sheet, 'A1:B4', { type: 'expression', formulas: ['MOD(ROW(),2)=0'], dxf: RED });
		expect(highlighted(fx, ['A1', 'A2', 'B2', 'A3', 'B4'])).toEqual(['A2', 'B2', 'B4']);
	});

	it('anchors mixed references at the top-left of the first range', () => {
		const fx = cfFixture({ C2: 'x', C3: 'y', C4: 'x' });
		addRule(fx.sheet, 'A2:B4', { type: 'expression', formulas: ['$C2="x"'], dxf: RED });
		expect(highlighted(fx, ['A2', 'B2', 'A3', 'B3', 'A4', 'B4'])).toEqual(['A2', 'B2', 'A4', 'B4']);
	});

	it('uses the anchor of the first range for later ranges too', () => {
		const fx = cfFixture({ B5: 1, B10: 0, D10: 1 });
		addRule(fx.sheet, 'A5 C10', { type: 'expression', formulas: ['B5=1'], dxf: RED });
		expect(fx.at('A5')?.style).toBe(RED);
		expect(fx.at('C10')?.style).toBe(RED);
	});

	it('treats numbers as conditions and errors / text as false', () => {
		const fx = cfFixture({ B1: 3, B2: 0, B3: 'yes', B4: ERRORS.NA });
		addRule(fx.sheet, 'A1:A5', { type: 'expression', formulas: ['B1'], dxf: RED });
		addRule(fx.sheet, 'C1', { type: 'expression', formulas: ['THROW()'], dxf: RED });
		expect(highlighted(fx, ['A1', 'A2', 'A3', 'A4', 'A5', 'C1'])).toEqual(['A1']);
	});

	it('takes the first element of array results', () => {
		const fx = cfFixture();
		addRule(fx.sheet, 'A1', { type: 'expression', formulas: ['ARR(1,0)'], dxf: RED });
		addRule(fx.sheet, 'A2', { type: 'expression', formulas: ['ARR(0,1)'], dxf: RED });
		expect(highlighted(fx, ['A1', 'A2'])).toEqual(['A1']);
	});
});

describe('text, blank and error rules', () => {
	const cells = {
		A1: 'Apple pie',
		A2: 'pineapple',
		A3: 'Banana',
		A4: 42,
		A5: ERRORS.NA,
		A6: '   ',
		A7: '',
	};
	const addrs = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8'];

	it.each([
		['containsText', 'APPLE', ['A1', 'A2']],
		['notContainsText', 'apple', ['A3', 'A4', 'A5', 'A6', 'A7', 'A8']],
		['beginsWith', 'apple', ['A1']],
		['endsWith', 'APPLE', ['A2']],
		['containsText', '4', ['A4']],
		['containsText', 'b?n*a', ['A3']],
	] as const)('%s "%s"', (type, text, expected) => {
		const fx = cfFixture(cells);
		addRule(fx.sheet, 'A1:A8', { type, text, dxf: RED });
		expect(highlighted(fx, addrs)).toEqual(expected);
	});

	it('falls back to the stored formula when the rule has no text', () => {
		const fx = cfFixture({ A1: 'x', A2: 'y' });
		addRule(fx.sheet, 'A1:A2', { type: 'containsText', formulas: ['A1="x"'], dxf: RED });
		expect(highlighted(fx, ['A1', 'A2'])).toEqual(['A1']);
	});

	it.each([
		['containsBlanks', ['A6', 'A7', 'A8']],
		['notContainsBlanks', ['A1', 'A2', 'A3', 'A4']],
		['containsErrors', ['A5']],
		['notContainsErrors', ['A1', 'A2', 'A3', 'A4', 'A6', 'A7', 'A8']],
	] as const)('%s', (type, expected) => {
		const fx = cfFixture(cells);
		addRule(fx.sheet, 'A1:A8', { type, dxf: RED });
		expect(highlighted(fx, addrs)).toEqual(expected);
	});
});

describe('timePeriod rules', () => {
	const today = dateToSerial(2024, 3, 13);
	it('uses the injected clock', () => {
		const fx = cfFixture(
			{ A1: today + 0.5, A2: today - 1, A3: today + 1, A4: 'today' },
			{ today: () => today },
		);
		addRule(fx.sheet, 'A1:A4', { type: 'timePeriod', timePeriod: 'today', dxf: RED });
		addRule(fx.sheet, 'B1:B2', { type: 'timePeriod', timePeriod: 'yesterday', dxf: RED });
		expect(highlighted(fx, ['A1', 'A2', 'A3', 'A4'])).toEqual(['A1']);
		fx.sheet.conditionalFormats[0].rules[0].timePeriod = 'tomorrow';
		fx.engine.invalidateSheet(0);
		expect(highlighted(fx, ['A1', 'A2', 'A3', 'A4'])).toEqual(['A3']);
	});
});

describe('top / bottom rules', () => {
	const values = [10, 10, 10, 5, 1, 7, 'x', null, 3, 8];
	const addrs = values.map((_, i) => `A${i + 1}`);

	function run(rule: { rank?: number; percent?: boolean; bottom?: boolean }) {
		const fx = cfFixture();
		setColumn(fx.sheet, 'A1', values);
		addRule(fx.sheet, 'A1:A10', { type: 'top10', dxf: RED, ...rule });
		return highlighted(fx, addrs);
	}

	it('includes ties for top N', () => {
		expect(run({ rank: 1 })).toEqual(['A1', 'A2', 'A3']);
		expect(run({ rank: 4 })).toEqual(['A1', 'A2', 'A3', 'A10']);
	});

	it('selects bottom N', () => {
		expect(run({ rank: 2, bottom: true })).toEqual(['A5', 'A9']);
	});

	it('selects a percentage of the numeric cells (at least one)', () => {
		expect(run({ rank: 50, percent: true })).toEqual(['A1', 'A2', 'A3', 'A10']);
		expect(run({ rank: 10, percent: true })).toEqual(['A1', 'A2', 'A3']);
		expect(run({ rank: 10, percent: true, bottom: true })).toEqual(['A5']);
	});

	it('highlights every number when N exceeds the count', () => {
		expect(run({ rank: 100 })).toEqual(['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A9', 'A10']);
		expect(run({ rank: 0 })).toEqual([]);
	});
});

describe('average rules', () => {
	// Mean 5, sample std dev ~3.1623.
	const values = [1, 3, 5, 7, 9, 'skip'];
	const addrs = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6'];
	function run(rule: { aboveAverage?: boolean; equalAverage?: boolean; stdDev?: number }) {
		const fx = cfFixture();
		setColumn(fx.sheet, 'A1', values);
		addRule(fx.sheet, 'A1:A6', { type: 'aboveAverage', dxf: RED, ...rule });
		return highlighted(fx, addrs);
	}

	it('handles above / below and equal variants', () => {
		expect(run({})).toEqual(['A4', 'A5']);
		expect(run({ aboveAverage: false })).toEqual(['A1', 'A2']);
		expect(run({ equalAverage: true })).toEqual(['A3', 'A4', 'A5']);
		expect(run({ aboveAverage: false, equalAverage: true })).toEqual(['A1', 'A2', 'A3']);
	});

	it('handles standard deviations', () => {
		expect(run({ stdDev: 1 })).toEqual(['A5']);
		expect(run({ stdDev: 1, aboveAverage: false })).toEqual(['A1']);
		expect(run({ stdDev: 2 })).toEqual([]);
	});
});

describe('duplicate / unique rules', () => {
	const values = ['a', 'A', 'b', 1, 1, 2, null, null, ERRORS.NA, ERRORS.NA, '1'];
	const addrs = values.map((_, i) => `A${i + 1}`);
	it('compares case-insensitively and ignores blanks and errors', () => {
		const dup = cfFixture();
		setColumn(dup.sheet, 'A1', values);
		addRule(dup.sheet, 'A1:A11', { type: 'duplicateValues', dxf: RED });
		expect(highlighted(dup, addrs)).toEqual(['A1', 'A2', 'A4', 'A5']);
		const uniq = cfFixture();
		setColumn(uniq.sheet, 'A1', values);
		addRule(uniq.sheet, 'A1:A11', { type: 'uniqueValues', dxf: RED });
		expect(highlighted(uniq, addrs)).toEqual(['A3', 'A6', 'A11']);
	});
});
