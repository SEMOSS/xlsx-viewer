import { describe, expect, it } from 'vitest';
import type { CfDataBar } from '../types/conditional-format';
import { layoutDataBar, mirrorDataBarLayout, isExtendedDataBar } from './data-bar';
import { addRule, cfFixture, setColumn } from './test-utils/fixtures';

const BLUE = { rgb: 'FF638EC6' };

function layout(
	value: number,
	lo: number,
	hi: number,
	axis: 'none' | 'automatic' | 'middle',
	minLength = 0,
	maxLength = 1,
) {
	return layoutDataBar({ value, lo, hi, minLength, maxLength, axis });
}

describe('layoutDataBar', () => {
	it('maps the 2007 scale onto min / max length', () => {
		expect(layout(0, 0, 100, 'none', 0.1, 0.9)).toEqual({ start: 0, end: 0.1 });
		expect(layout(100, 0, 100, 'none', 0.1, 0.9)).toEqual({ start: 0, end: 0.9 });
		expect(layout(50, 0, 100, 'none', 0.1, 0.9).end).toBeCloseTo(0.5);
		expect(layout(-20, 0, 100, 'none', 0.1, 0.9).end).toBeCloseTo(0.1);
		expect(layout(500, 0, 100, 'none', 0.1, 0.9).end).toBeCloseTo(0.9);
		expect(layout(7, 7, 7, 'none', 0.1, 0.9).end).toBeCloseTo(0.1);
	});

	it('places an automatic axis where zero falls', () => {
		const pos = layout(50, -50, 150, 'automatic');
		expect(pos.axis).toBeCloseTo(0.25);
		expect(pos.start).toBeCloseTo(0.25);
		expect(pos.end).toBeCloseTo(0.25 + 0.75 / 3);
		const neg = layout(-25, -50, 150, 'automatic');
		expect(neg.start).toBeCloseTo(0.125);
		expect(neg.end).toBeCloseTo(0.25);
		const zero = layout(0, -50, 150, 'automatic');
		expect(zero.start).toBe(zero.end);
		const max = layout(150, -50, 150, 'automatic');
		expect(max.end).toBeCloseTo(1);
		const min = layout(-80, -50, 150, 'automatic');
		expect(min.start).toBeCloseTo(0);
	});

	it('puts the axis on the right edge when everything is negative', () => {
		const bar = layout(-5, -10, 0, 'automatic');
		expect(bar.axis).toBe(1);
		expect(bar.start).toBeCloseTo(0.5);
		expect(bar.end).toBe(1);
	});

	it('has no axis when the automatic scale is non-negative', () => {
		expect(layout(5, 0, 10, 'automatic')).toEqual({ start: 0, end: 0.5 });
	});

	it('shares one scale around a middle axis', () => {
		const pos = layout(50, -100, 50, 'middle');
		expect(pos.axis).toBe(0.5);
		expect(pos.start).toBe(0.5);
		expect(pos.end).toBeCloseTo(0.75);
		const neg = layout(-100, -100, 50, 'middle');
		expect(neg.start).toBeCloseTo(0);
		expect(neg.end).toBe(0.5);
		const positiveOnly = layout(10, 0, 10, 'middle');
		expect(positiveOnly.end).toBeCloseTo(1);
	});

	it('applies min / max lengths to each side of the axis', () => {
		const bar = layout(1, -10, 10, 'automatic', 0.1, 0.9);
		expect(bar.start).toBeCloseTo(0.5);
		expect(bar.end).toBeCloseTo(0.5 + 0.5 * (0.1 + 0.1 * 0.8));
	});

	it('mirrors layouts for right-to-left bars', () => {
		expect(mirrorDataBarLayout({ start: 0, end: 0.4 })).toEqual({ start: 0.6, end: 1 });
		const m = mirrorDataBarLayout({ start: 0.25, end: 0.5, axis: 0.25 });
		expect(m).toEqual({ start: 0.5, end: 0.75, axis: 0.75 });
	});
});

describe('data bar rules', () => {
	function fixture(
		values: Array<number | string | null>,
		dataBar: Partial<CfDataBar>,
		rtl = false,
	) {
		const fx = cfFixture();
		fx.sheet.view.rightToLeft = rtl;
		setColumn(fx.sheet, 'A1', values);
		addRule(fx.sheet, `A1:A${values.length}`, {
			type: 'dataBar',
			dataBar: { cfvos: [{ type: 'min' }, { type: 'max' }], color: BLUE, ...dataBar },
		});
		return fx;
	}

	it('renders Excel 2007 bars: gradient, 10%..90%, no axis or negative colors', () => {
		const fx = fixture([0, 50, 100, -10, 'x'], {});
		expect(fx.at('A4')?.dataBar).toEqual({
			start: 0,
			end: 0.1,
			negative: false,
			color: '#638EC6',
			gradient: true,
			direction: 'ltr',
		});
		expect(fx.at('A1')?.dataBar?.end).toBeCloseTo(0.1 + (0.8 * 10) / 110);
		expect(fx.at('A3')?.dataBar?.end).toBeCloseTo(0.9);
		expect(fx.at('A3')?.dataBar?.axis).toBeUndefined();
		expect(fx.at('A5')).toBeUndefined();
		expect(isExtendedDataBar({ cfvos: [{ type: 'min' }, { type: 'max' }], color: BLUE })).toBe(
			false,
		);
	});

	it('renders Excel 2010 bars with an automatic axis and negative colors', () => {
		const fx = fixture([-50, 0, 150], {
			cfvos: [{ type: 'autoMin' }, { type: 'autoMax' }],
			minLength: 0,
			maxLength: 100,
			gradient: false,
			border: true,
			borderColor: { rgb: 'FF0000FF' },
			negativeFillColor: { rgb: 'FFFF0000' },
			negativeBorderColor: { rgb: 'FF800000' },
			axisColor: { rgb: 'FF333333' },
		});
		const neg = fx.at('A1')!.dataBar!;
		expect(neg.negative).toBe(true);
		expect(neg.color).toBe('#FF0000');
		expect(neg.borderColor).toBe('#800000');
		expect(neg.start).toBeCloseTo(0);
		expect(neg.end).toBeCloseTo(0.25);
		expect(neg.axis).toBeCloseTo(0.25);
		expect(neg.axisColor).toBe('#333333');
		expect(neg.gradient).toBe(false);
		const pos = fx.at('A3')!.dataBar!;
		expect(pos.negative).toBe(false);
		expect(pos.color).toBe('#638EC6');
		expect(pos.borderColor).toBe('#0000FF');
		expect(pos.end).toBeCloseTo(1);
		const zero = fx.at('A2')!.dataBar!;
		expect(zero.start).toBe(zero.end);
	});

	it('uses the positive color for negatives when configured and defaults to red otherwise', () => {
		const same = fixture([-1, 1], { negativeBarColorSameAsPositive: true, gradient: true });
		expect(same.at('A1')?.dataBar?.color).toBe('#638EC6');
		const dflt = fixture([-1, 1], { gradient: true });
		expect(dflt.at('A1')?.dataBar?.color).toBe('#FF0000');
		expect(dflt.at('A1')?.dataBar?.gradient).toBe(true);
	});

	it('honours axisPosition none and middle', () => {
		const none = fixture([-10, 10], { axisPosition: 'none', minLength: 0, maxLength: 100 });
		expect(none.at('A1')?.dataBar?.axis).toBeUndefined();
		expect(none.at('A1')?.dataBar?.start).toBe(0);
		expect(none.at('A1')?.dataBar?.end).toBe(0);
		expect(none.at('A2')?.dataBar?.end).toBe(1);
		const middle = fixture([-5, 10], { axisPosition: 'middle', minLength: 0, maxLength: 100 });
		expect(middle.at('A1')?.dataBar?.axis).toBe(0.5);
		expect(middle.at('A1')?.dataBar?.start).toBeCloseTo(0.25);
		expect(middle.at('A2')?.dataBar?.end).toBeCloseTo(1);
	});

	it('mirrors bars for right-to-left sheets and explicit directions', () => {
		const ctx = fixture([0, 10], { minLength: 0, maxLength: 100, gradient: true }, true);
		expect(ctx.at('A2')?.dataBar).toMatchObject({ start: 0, end: 1, direction: 'rtl' });
		expect(ctx.at('A1')?.dataBar).toMatchObject({ start: 1, end: 1 });
		const forced = fixture([5, 10], { minLength: 0, maxLength: 100, direction: 'rightToLeft' });
		expect(forced.at('A1')?.dataBar).toMatchObject({ start: 1, end: 1, direction: 'rtl' });
		const ltr = fixture([5, 10], { direction: 'leftToRight', minLength: 0, maxLength: 100 }, true);
		expect(ltr.at('A2')?.dataBar?.direction).toBe('ltr');
	});

	it('mirrors the axis and negative bars right to left', () => {
		const fx = fixture(
			[-50, 150],
			{ cfvos: [{ type: 'autoMin' }, { type: 'autoMax' }], minLength: 0, maxLength: 100 },
			true,
		);
		const neg = fx.at('A1')!.dataBar!;
		expect(neg.axis).toBeCloseTo(0.75);
		expect(neg.start).toBeCloseTo(0.75);
		expect(neg.end).toBeCloseTo(1);
		expect(neg).toMatchObject({ negative: true, direction: 'rtl', gradient: true });
		const pos = fx.at('A2')!.dataBar!;
		expect(pos.start).toBeCloseTo(0);
		expect(pos.end).toBeCloseTo(0.75);
	});

	it('resolves num, percent, percentile and formula thresholds', () => {
		const num = fixture([0, 5, 10], {
			cfvos: [
				{ type: 'num', value: '0' },
				{ type: 'num', value: '20' },
			],
			minLength: 0,
			maxLength: 100,
		});
		expect(num.at('A3')?.dataBar?.end).toBeCloseTo(0.5);
		const pct = fixture([0, 5, 10], {
			cfvos: [
				{ type: 'percent', value: '0' },
				{ type: 'percent', value: '50' },
			],
			minLength: 0,
			maxLength: 100,
		});
		expect(pct.at('A2')?.dataBar?.end).toBeCloseTo(1);
		const pctl = fixture([0, 5, 10], {
			cfvos: [
				{ type: 'percentile', value: '0' },
				{ type: 'percentile', value: '50' },
			],
			minLength: 0,
			maxLength: 100,
		});
		expect(pctl.at('A2')?.dataBar?.end).toBeCloseTo(1);
		const formula = fixture([0, 5, 10], {
			cfvos: [
				{ type: 'num', value: '0' },
				{ type: 'formula', value: '$C$1*2' },
			],
			minLength: 0,
			maxLength: 100,
		});
		formula.sheet.cells.set(2, { value: 10 });
		formula.engine.invalidate();
		expect(formula.at('A2')?.dataBar?.end).toBeCloseTo(0.25);
	});

	it('hides the value for "show bar only" and ignores lower-priority bars', () => {
		const fx = fixture([1, 2], { showValue: false });
		expect(fx.at('A1')?.hideValue).toBe(true);
		addRule(fx.sheet, 'A1:A2', {
			type: 'dataBar',
			priority: 1_000_000,
			dataBar: { cfvos: [{ type: 'min' }, { type: 'max' }], color: { rgb: 'FF00FF00' } },
		});
		fx.engine.invalidate();
		expect(fx.at('A2')?.dataBar?.color).toBe('#638EC6');
	});

	it('disables the rule when a threshold cannot be computed', () => {
		const fx = fixture([1, 2], { cfvos: [{ type: 'min' }, { type: 'formula', value: 'NA()' }] });
		expect(fx.at('A1')).toBeUndefined();
	});
});
