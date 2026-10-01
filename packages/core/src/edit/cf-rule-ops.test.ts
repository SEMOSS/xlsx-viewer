import { describe, expect, it } from 'vitest';

import { parseRange } from '../address/range';
import { addConditionalFormat } from '../conditional/edit';
import {
	createAverageRule,
	createCellIsRule,
	createColorScaleRule,
	createDataBarRule,
	createDateOccurringRule,
	createDuplicateRule,
	createTextRule,
	createTopBottomRule,
} from '../conditional/factories';
import { createWorkbook } from '../sheet/workbook';
import type { CfRule } from '../types/conditional-format';
import { describeConditionalRule, listManagedRules, replaceConditionalRules } from './cf-rule-ops';
import { EditTransaction } from './transaction';

const r = (a1: string) => parseRange(a1)!;

function sheetWithRules() {
	const wb = createWorkbook();
	const sheet = wb.sheets[0];
	addConditionalFormat(sheet, r('A1:A10'), createDataBarRule('FF638EC6'), wb.styles);
	addConditionalFormat(
		sheet,
		r('A1:A10'),
		createCellIsRule('greaterThan', ['5'], 'redText'),
		wb.styles,
	);
	addConditionalFormat(
		sheet,
		r('C1:C5'),
		createTextRule('containsText', 'ab', 'lightRedFill'),
		wb.styles,
	);
	return { wb, sheet };
}

describe('listManagedRules', () => {
	it('lists rules highest priority first with their ranges', () => {
		const { sheet } = sheetWithRules();
		const rules = listManagedRules(sheet.conditionalFormats);
		expect(rules.map((m) => m.rule.type)).toEqual(['containsText', 'cellIs', 'dataBar']);
		expect(rules[0].ranges).toEqual([r('C1:C5')]);
		expect(rules[1].ranges).toEqual([r('A1:A10')]);
	});
});

describe('replaceConditionalRules', () => {
	it('writes rules back in the given order, grouping identical ranges', () => {
		const { wb, sheet } = sheetWithRules();
		const rules = listManagedRules(sheet.conditionalFormats);
		const tx = new EditTransaction(wb, 'Manage Rules');
		replaceConditionalRules(tx, wb, 0, [
			rules[2],
			rules[0],
			{ ...rules[1], rule: { ...rules[1].rule, stopIfTrue: true } },
		]);
		const entry = tx.commit();
		const after = listManagedRules(sheet.conditionalFormats);
		expect(after.map((m) => [m.rule.type, m.rule.priority])).toEqual([
			['dataBar', 1],
			['containsText', 2],
			['cellIs', 3],
		]);
		expect(after[2].rule.stopIfTrue).toBe(true);
		expect(sheet.conditionalFormats).toHaveLength(2);
		entry.undo();
		expect(listManagedRules(sheet.conditionalFormats).map((m) => m.rule.type)).toEqual([
			'containsText',
			'cellIs',
			'dataBar',
		]);
	});

	it('drops deleted rules and rebuilds derived formulas for new ranges', () => {
		const { wb, sheet } = sheetWithRules();
		const [text] = listManagedRules(sheet.conditionalFormats);
		const tx = new EditTransaction(wb, 'Manage Rules');
		replaceConditionalRules(tx, wb, 0, [{ ...text, ranges: [r('E3:E9')] }]);
		tx.commit();
		expect(sheet.conditionalFormats).toEqual([
			{
				ranges: [r('E3:E9')],
				rules: [
					expect.objectContaining({
						type: 'containsText',
						priority: 1,
						formulas: ['NOT(ISERROR(SEARCH("ab",E3)))'],
					}),
				],
			},
		]);
	});

	it('keeps PivotTable blocks separate', () => {
		const wb = createWorkbook();
		const rule = createDataBarRule('FF638EC6');
		const tx = new EditTransaction(wb, 'x');
		replaceConditionalRules(tx, wb, 0, [
			{ rule, ranges: [r('A1:A3')], pivot: true },
			{ rule, ranges: [r('A1:A3')] },
		]);
		expect(wb.sheets[0].conditionalFormats.map((f) => Boolean(f.pivot))).toEqual([true, false]);
	});
});

describe('describeConditionalRule', () => {
	const d = (rule: CfRule) => describeConditionalRule(rule);
	it('uses the Rules Manager wording', () => {
		expect(d(createCellIsRule('greaterThan', ['5'], 'redText'))).toBe('Cell Value > 5');
		expect(d(createCellIsRule('between', ['1', '$B$1'], 'redText'))).toBe(
			'Cell Value between 1 and =$B$1',
		);
		expect(d(createCellIsRule('equal', ['"Done"'], 'redText'))).toBe('Cell Value = "Done"');
		expect(d({ type: 'expression', priority: 1, formulas: ['A1>0'] })).toBe('Formula: =A1>0');
		expect(d(createColorScaleRule(0))).toBe('Graded Color Scale');
		expect(d(createDataBarRule('FF638EC6'))).toBe('Data Bar');
		expect(d(createTopBottomRule({ rank: 5, bottom: true, percent: true }, 'redText'))).toBe(
			'Bottom 5%',
		);
		expect(d(createDuplicateRule(false, 'redText'))).toBe('Duplicate Values');
		expect(d(createTextRule('beginsWith', 'x', 'redText'))).toBe("Cell value begins with 'x'");
		expect(d(createDateOccurringRule('lastWeek', 'redText'))).toBe('Last week');
		expect(d(createAverageRule(false, 'redText', { equalAverage: true }))).toBe(
			'Equal to or Below Average',
		);
	});
});
