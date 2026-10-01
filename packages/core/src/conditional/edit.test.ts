import { describe, expect, it } from 'vitest';
import { createDefaultStyleSheet } from '../styles/defaults';
import type { CfRule } from '../types/conditional-format';
import {
	addConditionalFormat,
	clearConditionalFormats,
	getConditionalFormatRules,
	subtractRange,
} from './edit';
import { createCellIsRule, createTextRule } from './factories';
import { HIGHLIGHT_PRESETS } from './presets';
import { cfFixture, rng } from './test-utils/fixtures';

function priorities(rules: CfRule[]): number[] {
	return rules.map((r) => r.priority);
}

describe('subtractRange', () => {
	it('returns the uncovered parts', () => {
		expect(subtractRange(rng('A1:C3'), rng('E5'))).toEqual([rng('A1:C3')]);
		expect(subtractRange(rng('A1:C3'), rng('A1:C3'))).toEqual([]);
		expect(subtractRange(rng('A1:C3'), rng('B2'))).toEqual([
			rng('A1:C1'),
			rng('A2'),
			rng('C2'),
			rng('A3:C3'),
		]);
		expect(subtractRange(rng('A1:A10'), rng('A1:A3'))).toEqual([rng('A4:A10')]);
		expect(subtractRange(rng('A1:D1'), rng('C1:Z9'))).toEqual([rng('A1:B1')]);
	});
});

describe('addConditionalFormat', () => {
	it('gives the new rule priority 1 and shifts the others', () => {
		const { sheet } = cfFixture();
		sheet.dirty = false;
		const first = addConditionalFormat(
			sheet,
			rng('A1:A5'),
			createCellIsRule('greaterThan', ['5'], 'redText'),
		);
		const second = addConditionalFormat(
			sheet,
			[rng('B1:B5')],
			createCellIsRule('lessThan', ['0'], 'redText'),
		);
		expect(first.priority).toBe(1);
		expect(second.priority).toBe(1);
		expect(sheet.conditionalFormats).toHaveLength(2);
		expect(priorities(sheet.conditionalFormats[0].rules)).toEqual([2]);
		expect(priorities(sheet.conditionalFormats[1].rules)).toEqual([1]);
		expect(sheet.dirty).toBe(true);
		expect(getConditionalFormatRules(sheet).map((e) => e.rule.operator)).toEqual([
			'lessThan',
			'greaterThan',
		]);
	});

	it('joins a block with identical ranges', () => {
		const { sheet } = cfFixture();
		addConditionalFormat(sheet, rng('A1:A5'), createCellIsRule('greaterThan', ['5'], 'redText'));
		addConditionalFormat(sheet, rng('A1:A5'), createCellIsRule('lessThan', ['0'], 'redText'));
		expect(sheet.conditionalFormats).toHaveLength(1);
		expect(sheet.conditionalFormats[0].rules.map((r) => [r.operator, r.priority])).toEqual([
			['lessThan', 1],
			['greaterThan', 2],
		]);
	});

	it('does not mutate previous model objects', () => {
		const { sheet } = cfFixture();
		addConditionalFormat(sheet, rng('A1'), createCellIsRule('equal', ['1'], 'redText'));
		const before = sheet.conditionalFormats;
		const beforeRule = before[0].rules[0];
		addConditionalFormat(sheet, rng('B1'), createCellIsRule('equal', ['2'], 'redText'));
		expect(sheet.conditionalFormats).not.toBe(before);
		expect(beforeRule.priority).toBe(1);
	});

	it('interns inline formats into the style sheet', () => {
		const { sheet } = cfFixture();
		const styles = createDefaultStyleSheet();
		const a = addConditionalFormat(
			sheet,
			rng('A1'),
			createCellIsRule('equal', ['1'], 'lightRedFill'),
			styles,
		);
		const b = addConditionalFormat(
			sheet,
			rng('B1'),
			createCellIsRule('equal', ['2'], 'lightRedFill'),
			styles,
		);
		expect(a.dxf).toBeUndefined();
		expect(a.dxfId).toBe(0);
		expect(b.dxfId).toBe(0);
		expect(styles.dxfs).toEqual([HIGHLIGHT_PRESETS.lightRedFill]);
	});

	it('adds the formula Excel stores with text rules', () => {
		const { sheet } = cfFixture();
		const rule = addConditionalFormat(
			sheet,
			[rng('C3:C9'), rng('A1')],
			createTextRule('containsText', 'ok', 'redText'),
		);
		expect(rule.formulas).toEqual(['NOT(ISERROR(SEARCH("ok",C3)))']);
		expect(() => addConditionalFormat(sheet, [], rule)).toThrow(RangeError);
	});

	it('produces rules that render after invalidation', () => {
		const fx = cfFixture({ A1: 10 });
		addConditionalFormat(
			fx.sheet,
			rng('A1'),
			createCellIsRule('greaterThan', ['5'], 'greenFillDarkGreenText'),
		);
		fx.engine.invalidateSheet(0);
		expect(fx.at('A1')?.style).toEqual(HIGHLIGHT_PRESETS.greenFillDarkGreenText);
	});
});

describe('clearConditionalFormats', () => {
	it('clears the entire sheet', () => {
		const { sheet } = cfFixture();
		expect(clearConditionalFormats(sheet)).toBe(false);
		addConditionalFormat(sheet, rng('A1'), createCellIsRule('equal', ['1'], 'redText'));
		sheet.dirty = false;
		expect(clearConditionalFormats(sheet)).toBe(true);
		expect(sheet.conditionalFormats).toEqual([]);
		expect(sheet.dirty).toBe(true);
	});

	it('removes blocks inside the range and trims overlapping ones', () => {
		const { sheet } = cfFixture();
		addConditionalFormat(sheet, rng('A1:A3'), createCellIsRule('equal', ['1'], 'redText'));
		addConditionalFormat(sheet, rng('A1:C10'), createCellIsRule('equal', ['2'], 'redText'));
		addConditionalFormat(sheet, rng('F1:F5'), createCellIsRule('equal', ['3'], 'redText'));
		expect(clearConditionalFormats(sheet, rng('A1:A10'))).toBe(true);
		expect(sheet.conditionalFormats.map((f) => f.ranges)).toEqual([
			[rng('B1:C10')],
			[rng('F1:F5')],
		]);
		expect(sheet.conditionalFormats.map((f) => f.rules[0].priority)).toEqual([2, 1]);
	});

	it('re-anchors relative formulas of trimmed blocks', () => {
		const fx = cfFixture({ A4: 1, B4: 'x', B5: 'y' });
		fx.sheet.conditionalFormats.push({
			ranges: [rng('A1:A10')],
			rules: [
				{ type: 'expression', priority: 1, formulas: ['$B1="x"'], dxf: { font: { bold: true } } },
			],
		});
		expect(fx.at('A4')?.style).toBeDefined();
		clearConditionalFormats(fx.sheet, rng('A1:A3'));
		expect(fx.sheet.conditionalFormats[0].rules[0].formulas).toEqual(['$B4="x"']);
		fx.engine.invalidateSheet(0);
		expect(fx.at('A4')?.style).toBeDefined();
		expect(fx.at('A5')).toBeUndefined();
	});

	it('reports no change when nothing overlaps', () => {
		const { sheet } = cfFixture();
		addConditionalFormat(sheet, rng('A1'), createCellIsRule('equal', ['1'], 'redText'));
		sheet.dirty = false;
		expect(clearConditionalFormats(sheet, rng('Z1:Z9'))).toBe(false);
		expect(sheet.dirty).toBe(false);
	});
});
