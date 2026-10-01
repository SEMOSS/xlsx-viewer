import { describe, expect, it } from 'vitest';
import { putCell } from '../sheet/worksheet';
import type { DifferentialFormat, FillSpec } from '../types/style';
import { ConditionalFormatEngine } from './engine';
import { addRule, cfFixture, setCells, setColumn } from './test-utils/fixtures';

const fill = (rgb: string): FillSpec => ({
	type: 'pattern',
	patternType: 'solid',
	bgColor: { rgb },
});
const RED_TEXT: DifferentialFormat = { font: { color: { rgb: 'FFFF0000' } } };
const BOLD: DifferentialFormat = { font: { bold: true } };
const YELLOW_FILL: DifferentialFormat = { fill: fill('FFFFFF00') };
const GREEN_FILL: DifferentialFormat = {
	fill: fill('FF00FF00'),
	font: { italic: true, color: { rgb: 'FF00FF00' } },
};
const always = { type: 'expression' as const, formulas: ['TRUE'] };

describe('rule priority', () => {
	it('merges matching rules per property, higher priority first', () => {
		const fx = cfFixture({ A1: 1 });
		addRule(fx.sheet, 'A1', { ...always, priority: 2, dxf: GREEN_FILL });
		addRule(fx.sheet, 'A1', { ...always, priority: 1, dxf: { ...RED_TEXT, ...YELLOW_FILL } });
		addRule(fx.sheet, 'A1', { ...always, priority: 3, dxf: BOLD });
		expect(fx.at('A1')?.style).toEqual({
			font: { color: { rgb: 'FFFF0000' }, italic: true, bold: true },
			fill: fill('FFFFFF00'),
		});
	});

	it('stops at a matching stopIfTrue rule, including visual rules below it', () => {
		const fx = cfFixture({ A1: 5, A2: 50 });
		addRule(fx.sheet, 'A1:A2', {
			type: 'cellIs',
			operator: 'lessThan',
			formulas: ['10'],
			stopIfTrue: true,
			priority: 1,
		});
		addRule(fx.sheet, 'A1:A2', { ...always, priority: 2, dxf: BOLD });
		addRule(fx.sheet, 'A1:A2', {
			type: 'iconSet',
			priority: 3,
			iconSet: { iconSet: '3Arrows', cfvos: [] },
		});
		expect(fx.at('A1')).toBeUndefined();
		expect(fx.at('A2')?.style).toBe(BOLD);
		expect(fx.at('A2')?.icon).toEqual({ iconSet: '3Arrows', index: 2 });
	});

	it('ignores stopIfTrue on rules that do not match', () => {
		const fx = cfFixture({ A1: 50 });
		addRule(fx.sheet, 'A1', {
			type: 'cellIs',
			operator: 'lessThan',
			formulas: ['10'],
			stopIfTrue: true,
			priority: 1,
			dxf: RED_TEXT,
		});
		addRule(fx.sheet, 'A1', { ...always, priority: 2, dxf: BOLD });
		expect(fx.at('A1')?.style).toBe(BOLD);
	});

	it('breaks priority ties by document order', () => {
		const fx = cfFixture({ A1: 1 });
		addRule(fx.sheet, 'A1', { ...always, priority: 1, dxf: YELLOW_FILL });
		addRule(fx.sheet, 'A1', { ...always, priority: 1, dxf: GREEN_FILL });
		expect(fx.at('A1')?.style?.fill).toEqual(YELLOW_FILL.fill);
	});

	it('lets a higher-priority color scale take the fill from lower highlight rules', () => {
		const fx = cfFixture({ A1: 1, A2: 2 });
		addRule(fx.sheet, 'A1:A2', {
			type: 'colorScale',
			priority: 1,
			colorScale: {
				cfvos: [{ type: 'min' }, { type: 'max' }],
				colors: [{ rgb: 'FFFFFFFF' }, { rgb: 'FF000000' }],
			},
		});
		addRule(fx.sheet, 'A1:A2', { ...always, priority: 2, dxf: GREEN_FILL });
		const r = fx.at('A1')!;
		expect(r.colorScale).toBe('#FFFFFF');
		expect(r.style).toEqual({ font: GREEN_FILL.font });
	});

	it('lets a higher-priority fill suppress a lower color scale', () => {
		const fx = cfFixture({ A1: 1, A2: 2 });
		addRule(fx.sheet, 'A1:A2', { ...always, priority: 1, dxf: YELLOW_FILL });
		addRule(fx.sheet, 'A1:A2', {
			type: 'colorScale',
			priority: 2,
			colorScale: {
				cfvos: [{ type: 'min' }, { type: 'max' }],
				colors: [{ rgb: 'FFFFFFFF' }, { rgb: 'FF000000' }],
			},
		});
		expect(fx.at('A1')).toEqual({ style: YELLOW_FILL });
	});

	it('combines a data bar, an icon and a highlight on one cell', () => {
		const fx = cfFixture({ A1: 1, A2: 9 });
		addRule(fx.sheet, 'A1:A2', { ...always, priority: 1, dxf: BOLD });
		addRule(fx.sheet, 'A1:A2', {
			type: 'dataBar',
			priority: 2,
			dataBar: { cfvos: [{ type: 'min' }, { type: 'max' }], color: { rgb: 'FF638EC6' } },
		});
		addRule(fx.sheet, 'A1:A2', {
			type: 'iconSet',
			priority: 3,
			iconSet: { iconSet: '3Flags', cfvos: [] },
		});
		const r = fx.at('A2')!;
		expect(r.style).toBe(BOLD);
		expect(r.dataBar?.end).toBeCloseTo(0.9);
		expect(r.icon).toEqual({ iconSet: '3Flags', index: 2 });
		expect(r.hideValue).toBeUndefined();
	});
});

describe('formats', () => {
	it('resolves dxfId from the workbook styles and prefers inline dxf', () => {
		const fx = cfFixture({ A1: 1, A2: 1 });
		fx.workbook.styles.dxfs.push(RED_TEXT);
		addRule(fx.sheet, 'A1', { ...always, dxfId: 0 });
		addRule(fx.sheet, 'A2', { ...always, dxfId: 0, dxf: BOLD });
		expect(fx.at('A1')?.style).toBe(RED_TEXT);
		expect(fx.at('A2')?.style).toBe(BOLD);
	});

	it('returns nothing for matching rules without a format', () => {
		const fx = cfFixture({ A1: 1 });
		addRule(fx.sheet, 'A1', { ...always });
		expect(fx.at('A1')).toBeUndefined();
	});
});

describe('lookup and caching', () => {
	it('only applies rules inside their ranges, across many row bands', () => {
		const fx = cfFixture();
		addRule(fx.sheet, 'B200:C300 E1000', { ...always, dxf: BOLD });
		expect(fx.at('B199')).toBeUndefined();
		expect(fx.at('B200')?.style).toBe(BOLD);
		expect(fx.at('C300')?.style).toBe(BOLD);
		expect(fx.at('D250')).toBeUndefined();
		expect(fx.at('E1000')?.style).toBe(BOLD);
		expect(fx.at('E999')).toBeUndefined();
		expect(fx.engine.getCellResult(0, 1_048_575, 16_383)).toBeUndefined();
	});

	it('memoises cell results and statistics until invalidated', () => {
		const fx = cfFixture({ B1: 1 });
		addRule(fx.sheet, 'A1:A3', { type: 'expression', formulas: ['$B$1=1'], dxf: BOLD });
		const first = fx.at('A1');
		const calls = fx.evaluator.calls;
		expect(fx.at('A1')).toBe(first);
		expect(fx.evaluator.calls).toBe(calls);
		setCells(fx.sheet, { B1: 2 });
		expect(fx.at('A1')).toBe(first);
		fx.engine.invalidateSheet(0);
		expect(fx.at('A1')).toBeUndefined();
	});

	it('recomputes range statistics only after invalidation', () => {
		const fx = cfFixture();
		setColumn(fx.sheet, 'A1', [1, 2, 3]);
		addRule(fx.sheet, 'A1:A4', { type: 'top10', rank: 1, dxf: BOLD });
		expect(fx.at('A3')?.style).toBe(BOLD);
		putCell(fx.sheet, 3, 0, { value: 10 });
		// The cached threshold is still 3 until the sheet is invalidated.
		expect(fx.at('A2')).toBeUndefined();
		expect(fx.at('A3')?.style).toBe(BOLD);
		fx.engine.invalidateSheet(0);
		expect(fx.at('A4')?.style).toBe(BOLD);
		expect(fx.at('A3')).toBeUndefined();
	});

	it('drops formula-dependent results of other sheets on invalidateSheet', () => {
		const fx = cfFixture({}, { sheets: ['Main', 'Data'] });
		const data = fx.workbook.sheets[1];
		setCells(data, { A1: 1 });
		addRule(fx.sheet, 'A1', { type: 'expression', formulas: ['Data!$A$1=1'], dxf: BOLD });
		expect(fx.at('A1')?.style).toBe(BOLD);
		setCells(data, { A1: 0 });
		fx.engine.invalidateSheet(1);
		expect(fx.at('A1')).toBeUndefined();
	});

	it('keeps caches of unrelated sheets on invalidateSheet', () => {
		const fx = cfFixture({ A1: 1 }, { sheets: ['Main', 'Data'] });
		addRule(fx.sheet, 'A1', { type: 'cellIs', operator: 'equal', formulas: ['1'], dxf: BOLD });
		const before = fx.at('A1');
		expect(before?.style).toBe(BOLD);
		fx.engine.invalidateSheet(1);
		expect(fx.at('A1')).toBe(before);
		fx.engine.invalidateSheet(0);
		expect(fx.at('A1')).not.toBe(before);
		expect(fx.at('A1')).toEqual(before);
	});

	it('invalidate() picks up rule edits', () => {
		const fx = cfFixture({ A1: 1 });
		expect(fx.engine.hasRules(0)).toBe(false);
		addRule(fx.sheet, 'A1', { ...always, dxf: BOLD });
		expect(fx.at('A1')).toBeUndefined();
		fx.engine.invalidate();
		expect(fx.engine.hasRules(0)).toBe(true);
		expect(fx.at('A1')?.style).toBe(BOLD);
	});

	it('handles unknown sheets and blank workbooks', () => {
		const fx = cfFixture();
		expect(fx.engine.getCellResult(5, 0, 0)).toBeUndefined();
		expect(fx.engine.hasRules(5)).toBe(false);
		const engine = new ConditionalFormatEngine(fx.workbook, () => null);
		expect(engine.getCellResult(0, 0, 0)).toBeUndefined();
	});

	it('skips blocks without usable ranges', () => {
		const fx = cfFixture({ A1: 1 });
		fx.sheet.conditionalFormats.push({
			ranges: [],
			rules: [{ ...always, priority: 1, dxf: BOLD }],
		});
		fx.sheet.conditionalFormats.push({
			ranges: [{ top: Number.NaN, left: 0, bottom: 0, right: 0 }],
			rules: [{ ...always, priority: 1, dxf: BOLD }],
		});
		expect(fx.at('A1')).toBeUndefined();
	});

	it('serves a full visible window quickly', () => {
		const fx = cfFixture();
		for (let r = 0; r < 2000; r++) {
			putCell(fx.sheet, r, 0, { value: r });
		}
		addRule(fx.sheet, 'A1:A2000', {
			type: 'colorScale',
			colorScale: {
				cfvos: [{ type: 'min' }, { type: 'percentile', value: '50' }, { type: 'max' }],
				colors: [{ rgb: 'FFF8696B' }, { rgb: 'FFFFEB84' }, { rgb: 'FF63BE7B' }],
			},
		});
		addRule(fx.sheet, 'A1:A2000', { type: 'top10', rank: 10, percent: true, dxf: BOLD });
		const start = performance.now();
		for (let frame = 0; frame < 20; frame++) {
			for (let r = 0; r < 60; r++) {
				for (let c = 0; c < 20; c++) {
					fx.engine.getCellResult(0, 900 + r, c);
				}
			}
		}
		expect(performance.now() - start).toBeLessThan(500);
		expect(fx.at('A2000')?.style).toBe(BOLD);
		expect(fx.at('A1000')?.colorScale).toBe('#FFEB84');
	});
});
