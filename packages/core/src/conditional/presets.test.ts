import { describe, expect, it } from 'vitest';
import {
	createAverageRule,
	createCellIsRule,
	createColorScaleRule,
	createDataBarRule,
	createDateOccurringRule,
	createDuplicateRule,
	createIconSetRule,
	createTextRule,
	createTopBottomRule,
	toCriteriaFormula,
} from './factories';
import {
	COLOR_SCALE_PRESETS,
	DATA_BAR_PRESETS,
	HIGHLIGHT_PRESET_LABELS,
	HIGHLIGHT_PRESETS,
} from './presets';
import { buildConditionalRuleFormula as buildRuleFormula } from './rule-formulas';
import { addRule, cfFixture, setColumn } from './test-utils/fixtures';

describe('HIGHLIGHT_PRESETS', () => {
	it('matches Excel colors', () => {
		const p = HIGHLIGHT_PRESETS;
		expect(p.lightRedFillDarkRedText.font?.color?.rgb).toBe('FF9C0006');
		expect(p.lightRedFillDarkRedText.fill).toEqual({
			type: 'pattern',
			patternType: 'solid',
			fgColor: { rgb: 'FFFFC7CE' },
			bgColor: { rgb: 'FFFFC7CE' },
		});
		expect(p.yellowFillDarkYellowText.font?.color?.rgb).toBe('FF9C5700');
		expect(p.yellowFillDarkYellowText.fill).toMatchObject({ bgColor: { rgb: 'FFFFEB9C' } });
		expect(p.greenFillDarkGreenText.font?.color?.rgb).toBe('FF006100');
		expect(p.greenFillDarkGreenText.fill).toMatchObject({ bgColor: { rgb: 'FFC6EFCE' } });
		expect(p.lightRedFill.font).toBeUndefined();
		expect(p.redText.fill).toBeUndefined();
		expect(p.redBorder.border?.top).toEqual({ style: 'thin', color: { rgb: 'FF9C0006' } });
		expect(Object.isFrozen(p.lightRedFill.fill)).toBe(true);
		expect(HIGHLIGHT_PRESET_LABELS.lightRedFillDarkRedText).toBe(
			'Light Red Fill with Dark Red Text',
		);
	});
});

describe('rule factories', () => {
	it('creates cell value rules with cloned formats', () => {
		const rule = createCellIsRule('between', ['=1', '10'], 'lightRedFill');
		expect(rule).toMatchObject({
			type: 'cellIs',
			operator: 'between',
			formulas: ['1', '10'],
			priority: 1,
		});
		expect(rule.dxf).toEqual(HIGHLIGHT_PRESETS.lightRedFill);
		expect(rule.dxf).not.toBe(HIGHLIGHT_PRESETS.lightRedFill);
		const custom = createCellIsRule('equal', ['"x"'], { font: { bold: true } });
		expect(custom.dxf).toEqual({ font: { bold: true } });
	});

	it('creates text, date, duplicate, top/bottom and average rules', () => {
		expect(createTextRule('notContainsText', 'abc', 'redText')).toMatchObject({
			type: 'notContainsText',
			operator: 'notContains',
			text: 'abc',
		});
		expect(createTextRule('beginsWith', 'a', 'redText').operator).toBe('beginsWith');
		expect(createDateOccurringRule('lastWeek', 'redText')).toMatchObject({
			type: 'timePeriod',
			timePeriod: 'lastWeek',
		});
		expect(createDuplicateRule(false, 'redText').type).toBe('duplicateValues');
		expect(createDuplicateRule(true, 'redText').type).toBe('uniqueValues');
		expect(createTopBottomRule({}, 'redText')).toMatchObject({ type: 'top10', rank: 10 });
		const bottom = createTopBottomRule({ bottom: true, percent: true, rank: 5 }, 'redText');
		expect(bottom).toMatchObject({ rank: 5, percent: true, bottom: true });
		expect(createAverageRule(true, 'redText').aboveAverage).toBeUndefined();
		expect(createAverageRule(false, 'redText', { equalAverage: true, stdDev: 2 })).toMatchObject({
			aboveAverage: false,
			equalAverage: true,
			stdDev: 2,
		});
	});

	it('creates Excel 2010 data bars', () => {
		const gradient = createDataBarRule('#638EC6');
		expect(gradient.dataBar).toMatchObject({
			cfvos: [{ type: 'autoMin' }, { type: 'autoMax' }],
			color: { rgb: 'FF638EC6' },
			gradient: true,
			border: true,
			borderColor: { rgb: 'FF638EC6' },
			negativeFillColor: { rgb: 'FFFF0000' },
			axisPosition: 'automatic',
			minLength: 0,
			maxLength: 100,
		});
		const solid = createDataBarRule({ theme: 5 }, false);
		expect(solid.dataBar?.gradient).toBe(false);
		expect(solid.dataBar?.border).toBe(false);
		expect(solid.dataBar?.borderColor).toBeUndefined();
		expect(solid.dataBar?.color).toEqual({ theme: 5 });
		expect(DATA_BAR_PRESETS).toHaveLength(6);
	});

	it('creates the twelve color scale presets', () => {
		expect(COLOR_SCALE_PRESETS).toHaveLength(12);
		const first = createColorScaleRule(0);
		expect(first.colorScale?.cfvos).toEqual([
			{ type: 'min' },
			{ type: 'percentile', value: '50' },
			{ type: 'max' },
		]);
		expect(first.colorScale?.colors).toEqual([
			{ rgb: 'FFF8696B' },
			{ rgb: 'FFFFEB84' },
			{ rgb: 'FF63BE7B' },
		]);
		const whiteRed = createColorScaleRule(6);
		expect(whiteRed.colorScale?.cfvos).toEqual([{ type: 'min' }, { type: 'max' }]);
		expect(COLOR_SCALE_PRESETS[6].name).toBe('White - Red Color Scale');
		expect(() => createColorScaleRule(12)).toThrow(RangeError);
	});

	it('creates icon set rules with default thresholds', () => {
		const rule = createIconSetRule('4Arrows', { reverse: true, showValue: false });
		expect(rule.iconSet).toEqual({
			iconSet: '4Arrows',
			cfvos: [
				{ type: 'percent', value: '0' },
				{ type: 'percent', value: '25' },
				{ type: 'percent', value: '50' },
				{ type: 'percent', value: '75' },
			],
			reverse: true,
			showValue: false,
		});
		expect(() => createIconSetRule('NoIcons')).toThrow(RangeError);
	});

	it('produces rules the engine evaluates', () => {
		const fx = cfFixture();
		setColumn(fx.sheet, 'A1', [1, 2, 3, 4, 5]);
		const { type: _t, priority: _p, ...scale } = createColorScaleRule(0);
		addRule(fx.sheet, 'A1:A5', { type: 'colorScale', ...scale });
		expect(fx.at('A5')?.colorScale).toBe('#63BE7B');
		const { type: _t2, priority: _p2, ...bar } = createDataBarRule('FF0000');
		addRule(fx.sheet, 'A1:A5', { type: 'dataBar', ...bar });
		fx.engine.invalidate();
		expect(fx.at('A5')?.dataBar).toMatchObject({
			start: 0,
			end: 1,
			color: '#FF0000',
			borderColor: '#FF0000',
		});
	});
});

describe('toCriteriaFormula', () => {
	it('converts dialog input like Excel', () => {
		expect(toCriteriaFormula('10')).toBe('10');
		expect(toCriteriaFormula(' 1,000 ')).toBe('1000');
		expect(toCriteriaFormula('50%')).toBe('0.5');
		expect(toCriteriaFormula('1/1/2024')).toBe('45292');
		expect(toCriteriaFormula('=$A$1')).toBe('$A$1');
		expect(toCriteriaFormula('true')).toBe('TRUE');
		expect(toCriteriaFormula('say "hi"')).toBe('"say ""hi"""');
		expect(toCriteriaFormula(2.5)).toBe('2.5');
		expect(toCriteriaFormula(false)).toBe('FALSE');
	});
});

describe('buildRuleFormula', () => {
	it('writes the formulas Excel stores with text, blank, error and date rules', () => {
		expect(buildRuleFormula({ type: 'containsText', priority: 1, text: 'a"b' }, 1, 2)).toBe(
			'NOT(ISERROR(SEARCH("a""b",C2)))',
		);
		expect(buildRuleFormula({ type: 'notContainsText', priority: 1, text: 'x' }, 0, 0)).toBe(
			'ISERROR(SEARCH("x",A1))',
		);
		expect(buildRuleFormula({ type: 'beginsWith', priority: 1, text: 'x' }, 0, 0)).toBe(
			'LEFT(A1,LEN("x"))="x"',
		);
		expect(buildRuleFormula({ type: 'endsWith', priority: 1, text: 'x' }, 0, 0)).toBe(
			'RIGHT(A1,LEN("x"))="x"',
		);
		expect(buildRuleFormula({ type: 'containsBlanks', priority: 1 }, 0, 0)).toBe('LEN(TRIM(A1))=0');
		expect(buildRuleFormula({ type: 'notContainsBlanks', priority: 1 }, 0, 0)).toBe(
			'LEN(TRIM(A1))>0',
		);
		expect(buildRuleFormula({ type: 'containsErrors', priority: 1 }, 0, 0)).toBe('ISERROR(A1)');
		expect(buildRuleFormula({ type: 'notContainsErrors', priority: 1 }, 0, 0)).toBe(
			'NOT(ISERROR(A1))',
		);
		expect(buildRuleFormula({ type: 'timePeriod', priority: 1, timePeriod: 'today' }, 4, 1)).toBe(
			'FLOOR(B5,1)=TODAY()',
		);
		expect(
			buildRuleFormula({ type: 'timePeriod', priority: 1, timePeriod: 'last7Days' }, 0, 0),
		).toBe('AND(TODAY()-FLOOR(A1,1)<=6,FLOOR(A1,1)<=TODAY())');
		expect(
			buildRuleFormula({ type: 'timePeriod', priority: 1, timePeriod: 'lastMonth' }, 0, 0),
		).toContain('EDATE(TODAY(),0-1)');
		expect(
			buildRuleFormula({ type: 'timePeriod', priority: 1, timePeriod: 'nextWeek' }, 0, 0),
		).toContain('15-WEEKDAY(TODAY())');
		expect(buildRuleFormula({ type: 'top10', priority: 1 }, 0, 0)).toBeUndefined();
	});
});
