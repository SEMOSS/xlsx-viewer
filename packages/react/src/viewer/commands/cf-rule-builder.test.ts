import { describe, expect, it } from 'vitest';
import type { HighlightPresetName } from 'xlsx-viewer-core';
import { createWorkbook, HIGHLIGHT_PRESET_LABELS, parseRange } from 'xlsx-viewer-core';

import { typeCells } from '../components/dialogs/__tests__/render-dialog';
import { createTranslator } from '../../i18n';
import {
	buildConditionalRule,
	CONDITIONAL_RULES,
	HIGHLIGHT_PRESET_KEYS,
	suggestRuleValues,
	type ConditionalRuleInput,
} from './cf-rule-builder';

const input = (over: Partial<ConditionalRuleInput> = {}): ConditionalRuleInput => ({
	values: ['', ''],
	text: '',
	period: 'yesterday',
	unique: false,
	rank: '10',
	format: 'lightRedFillDarkRedText',
	...over,
});

function rule(result: ReturnType<typeof buildConditionalRule>) {
	if ('error' in result) {
		throw new Error(result.error);
	}
	return result.rule;
}

describe('buildConditionalRule', () => {
	it('builds cell value rules from typed values, dates and formulas', () => {
		expect(rule(buildConditionalRule('greaterThan', input({ values: ['50', ''] })))).toMatchObject({
			type: 'cellIs',
			operator: 'greaterThan',
			formulas: ['50'],
		});
		expect(
			rule(buildConditionalRule('between', input({ values: ['1/1/2024', '=$B$1'] }))),
		).toMatchObject({
			operator: 'between',
			formulas: ['45292', '$B$1'],
		});
		expect(rule(buildConditionalRule('equalTo', input({ values: ['Done', ''] })))).toMatchObject({
			operator: 'equal',
			formulas: ['"Done"'],
		});
		expect(rule(buildConditionalRule('lessThan', input({ values: ['10%', ''] }))).formulas).toEqual(
			['0.1'],
		);
	});

	it('uses the chosen highlight preset as the format', () => {
		const r = rule(
			buildConditionalRule(
				'greaterThan',
				input({ values: ['1', ''], format: 'greenFillDarkGreenText' }),
			),
		);
		expect(r.dxf?.font?.color).toEqual({ rgb: 'FF006100' });
	});

	it('builds text, date, duplicate, top / bottom and average rules', () => {
		expect(rule(buildConditionalRule('textContains', input({ text: 'abc' })))).toMatchObject({
			type: 'containsText',
			text: 'abc',
		});
		expect(
			rule(buildConditionalRule('dateOccurring', input({ period: 'last7Days' }))),
		).toMatchObject({ type: 'timePeriod', timePeriod: 'last7Days' });
		expect(rule(buildConditionalRule('duplicateValues', input({ unique: true }))).type).toBe(
			'uniqueValues',
		);
		expect(rule(buildConditionalRule('top10', input({ rank: '5' })))).toMatchObject({
			type: 'top10',
			rank: 5,
		});
		expect(rule(buildConditionalRule('bottomPercent', input({ rank: '20' })))).toMatchObject({
			rank: 20,
			percent: true,
			bottom: true,
		});
		expect(rule(buildConditionalRule('aboveAverage', input())).aboveAverage).toBeUndefined();
		expect(rule(buildConditionalRule('belowAverage', input())).aboveAverage).toBe(false);
	});

	it('rejects missing values and out-of-range ranks', () => {
		expect(buildConditionalRule('greaterThan', input())).toEqual({
			error: 'Enter a value or a formula.',
		});
		expect(buildConditionalRule('between', input({ values: ['1', ''] }))).toHaveProperty('error');
		expect(buildConditionalRule('textContains', input())).toHaveProperty('error');
		expect(buildConditionalRule('top10', input({ rank: '0' }))).toEqual({
			error: 'Enter a whole number between 1 and 1000.',
		});
		expect(buildConditionalRule('topPercent', input({ rank: '101' }))).toEqual({
			error: 'Enter a whole number between 1 and 100.',
		});
		expect(buildConditionalRule('top10', input({ rank: '2.5' }))).toHaveProperty('error');
	});

	it('describes every rule kind', () => {
		const t = createTranslator();
		for (const spec of Object.values(CONDITIONAL_RULES)) {
			expect(t(spec.titleKey)).not.toBe('');
			expect(t(spec.promptKey)).toMatch(/^Format cells/);
		}
	});

	it('captions the highlight presets as core names them in English', () => {
		const t = createTranslator();
		for (const [name, label] of Object.entries(HIGHLIGHT_PRESET_LABELS)) {
			expect(t(HIGHLIGHT_PRESET_KEYS[name as HighlightPresetName])).toBe(label);
		}
	});
});

describe('suggestRuleValues', () => {
	it('pre-fills from the selected numbers and the active cell text', () => {
		const wb = createWorkbook();
		typeCells(wb, { A1: '10', A2: '20', A3: '30', A4: '40', A5: 'label' });
		const sheet = wb.sheets[0];
		const ranges = [parseRange('A1:A5')!];
		expect(suggestRuleValues('greaterThan', sheet, ranges, { row: 0, col: 0 }).values).toEqual([
			'25',
			'',
		]);
		expect(suggestRuleValues('between', sheet, ranges, { row: 0, col: 0 }).values).toEqual([
			'17.5',
			'32.5',
		]);
		expect(suggestRuleValues('textContains', sheet, ranges, { row: 4, col: 0 }).text).toBe('label');
		expect(
			suggestRuleValues('greaterThan', sheet, [parseRange('C1:C3')!], { row: 0, col: 2 }).values,
		).toEqual(['', '']);
	});
});
