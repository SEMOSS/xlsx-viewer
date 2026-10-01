import { describe, expect, it } from 'vitest';
import { parseHex } from '../color/color-math';
import { createWorkbook } from '../sheet/workbook';
import { dateToSerial } from '../values/dates';
import { ERRORS } from '../values/errors';
import { interpolateColorScale, lerpRgb, rgbToCss } from './color-scale';
import { dxfBackgroundColor, mergeDifferentialFormats } from './dxf';
import { RangeStatistics } from './statistics';
import { setColumn } from './test-utils/fixtures';
import {
	cellValueText,
	createSearchMatcher,
	createTextMatcher,
	isBlankForRule,
} from './text-match';
import { cfvoNumber, resolveCfvo } from './thresholds';
import { matchesTimePeriod } from './time-period';

describe('text matching', () => {
	it('converts values like formulas do', () => {
		expect(cellValueText(null)).toBe('');
		expect(cellValueText(1.5)).toBe('1.5');
		expect(cellValueText(true)).toBe('TRUE');
		expect(cellValueText(ERRORS.NA)).toBeUndefined();
	});

	it('detects blanks like LEN(TRIM(x))=0', () => {
		expect(isBlankForRule(null)).toBe(true);
		expect(isBlankForRule('')).toBe(true);
		expect(isBlankForRule('   ')).toBe(true);
		expect(isBlankForRule(' a ')).toBe(false);
		expect(isBlankForRule(0)).toBe(false);
		expect(isBlankForRule(false)).toBe(false);
		expect(isBlankForRule(ERRORS.NA)).toBeUndefined();
	});

	it('supports SEARCH wildcards', () => {
		expect(createSearchMatcher('a*c')('xxABCxx')).toBe(true);
		expect(createSearchMatcher('a?c')('abc')).toBe(true);
		expect(createSearchMatcher('a?c')('ac')).toBe(false);
		expect(createSearchMatcher('~*')('5*3')).toBe(true);
		expect(createSearchMatcher('~*')('53')).toBe(false);
		expect(createSearchMatcher('(x)')('a(X)b')).toBe(true);
	});

	it('matches the text rule family', () => {
		const contains = createTextMatcher('containsText', 'ap');
		expect(contains('Apple')).toBe(true);
		expect(contains('pear')).toBe(false);
		expect(contains(ERRORS.NA)).toBe(false);
		expect(contains(null)).toBe(false);
		const notContains = createTextMatcher('notContainsText', 'ap');
		expect(notContains('Apple')).toBe(false);
		expect(notContains('pear')).toBe(true);
		expect(notContains(ERRORS.NA)).toBe(true);
		expect(notContains(null)).toBe(true);
		const begins = createTextMatcher('beginsWith', 'AB');
		expect(begins('abc')).toBe(true);
		expect(begins('cab')).toBe(false);
		expect(begins(ERRORS.NA)).toBe(false);
		const ends = createTextMatcher('endsWith', '.5');
		expect(ends(2.5)).toBe(true);
		expect(ends('x.5')).toBe(true);
		expect(createTextMatcher('beginsWith', 'a*')('abc')).toBe(false);
		expect(createTextMatcher('containsText', '')(null)).toBe(true);
	});
});

describe('matchesTimePeriod', () => {
	// Wednesday 2024-03-13.
	const today = dateToSerial(2024, 3, 13);
	const d = (y: number, m: number, day: number) => dateToSerial(y, m, day);

	it('handles single days and the last 7 days', () => {
		expect(matchesTimePeriod(today + 0.75, 'today', today)).toBe(true);
		expect(matchesTimePeriod(today - 1, 'today', today)).toBe(false);
		expect(matchesTimePeriod(today - 1, 'yesterday', today)).toBe(true);
		expect(matchesTimePeriod(today + 1, 'tomorrow', today)).toBe(true);
		expect(matchesTimePeriod(today - 6, 'last7Days', today)).toBe(true);
		expect(matchesTimePeriod(today - 7, 'last7Days', today)).toBe(false);
		expect(matchesTimePeriod(today + 1, 'last7Days', today)).toBe(false);
		expect(matchesTimePeriod(today, 'last7Days', today)).toBe(true);
	});

	it('uses Sunday-to-Saturday weeks', () => {
		expect(matchesTimePeriod(d(2024, 3, 10), 'thisWeek', today)).toBe(true);
		expect(matchesTimePeriod(d(2024, 3, 16), 'thisWeek', today)).toBe(true);
		expect(matchesTimePeriod(d(2024, 3, 9), 'thisWeek', today)).toBe(false);
		expect(matchesTimePeriod(d(2024, 3, 3), 'lastWeek', today)).toBe(true);
		expect(matchesTimePeriod(d(2024, 3, 9), 'lastWeek', today)).toBe(true);
		expect(matchesTimePeriod(d(2024, 3, 10), 'lastWeek', today)).toBe(false);
		expect(matchesTimePeriod(d(2024, 3, 17), 'nextWeek', today)).toBe(true);
		expect(matchesTimePeriod(d(2024, 3, 23), 'nextWeek', today)).toBe(true);
		expect(matchesTimePeriod(d(2024, 3, 24), 'nextWeek', today)).toBe(false);
	});

	it('compares calendar months across year boundaries', () => {
		expect(matchesTimePeriod(d(2024, 3, 1), 'thisMonth', today)).toBe(true);
		expect(matchesTimePeriod(d(2023, 3, 13), 'thisMonth', today)).toBe(false);
		expect(matchesTimePeriod(d(2024, 2, 29), 'lastMonth', today)).toBe(true);
		expect(matchesTimePeriod(d(2024, 4, 30), 'nextMonth', today)).toBe(true);
		const jan = d(2024, 1, 15);
		expect(matchesTimePeriod(d(2023, 12, 31), 'lastMonth', jan)).toBe(true);
		const dec = d(2023, 12, 5);
		expect(matchesTimePeriod(d(2024, 1, 1), 'nextMonth', dec)).toBe(true);
		expect(matchesTimePeriod(-5, 'today', today)).toBe(false);
	});

	it('supports the 1904 date system', () => {
		const t1904 = dateToSerial(2024, 3, 13, true);
		expect(matchesTimePeriod(dateToSerial(2024, 3, 10, true), 'thisWeek', t1904, true)).toBe(true);
		expect(matchesTimePeriod(dateToSerial(2024, 2, 10, true), 'lastMonth', t1904, true)).toBe(true);
	});
});

describe('thresholds', () => {
	const sheet = createWorkbook().sheets[0];
	setColumn(sheet, 'A1', [-10, 0, 10, 20, 30]);
	const stats = new RangeStatistics(sheet, [{ top: 0, left: 0, bottom: 9, right: 0 }]);
	const evaluate = (f: string) => (f === '$B$1' ? 7 : undefined);

	it('resolves every cfvo type', () => {
		expect(resolveCfvo({ type: 'min' }, stats, evaluate)).toBe(-10);
		expect(resolveCfvo({ type: 'max' }, stats, evaluate)).toBe(30);
		expect(resolveCfvo({ type: 'autoMin' }, stats, evaluate)).toBe(-10);
		expect(resolveCfvo({ type: 'autoMax' }, stats, evaluate)).toBe(30);
		expect(resolveCfvo({ type: 'num', value: '12.5' }, stats, evaluate)).toBe(12.5);
		expect(resolveCfvo({ type: 'percent', value: '50' }, stats, evaluate)).toBe(10);
		expect(resolveCfvo({ type: 'percentile', value: '25' }, stats, evaluate)).toBe(0);
		expect(resolveCfvo({ type: 'formula', value: '$B$1' }, stats, evaluate)).toBe(7);
		expect(resolveCfvo({ type: 'num', value: '=$B$1' }, stats, evaluate)).toBe(7);
		expect(resolveCfvo({ type: 'formula', value: 'XYZ' }, stats, evaluate)).toBeUndefined();
		expect(resolveCfvo(undefined, stats, evaluate)).toBeUndefined();
	});

	it('clamps automatic bounds to zero', () => {
		const positive = createWorkbook().sheets[0];
		setColumn(positive, 'A1', [5, 8]);
		const s = new RangeStatistics(positive, [{ top: 0, left: 0, bottom: 1, right: 0 }]);
		expect(resolveCfvo({ type: 'autoMin' }, s, evaluate)).toBe(0);
		expect(resolveCfvo({ type: 'min' }, s, evaluate)).toBe(5);
		expect(cfvoNumber(' ', evaluate)).toBeUndefined();
		expect(cfvoNumber(undefined, evaluate)).toBeUndefined();
	});
});

describe('color scale interpolation', () => {
	const red = parseHex('F8696B')!;
	const yellow = parseHex('FFEB84')!;
	const green = parseHex('63BE7B')!;

	it('interpolates two colors', () => {
		const white = parseHex('FFFFFF')!;
		const black = parseHex('000000')!;
		expect(rgbToCss(interpolateColorScale(0, [0, 10], [white, black]))).toBe('#FFFFFF');
		expect(rgbToCss(interpolateColorScale(10, [0, 10], [white, black]))).toBe('#000000');
		expect(rgbToCss(interpolateColorScale(5, [0, 10], [white, black]))).toBe('#808080');
		expect(rgbToCss(interpolateColorScale(-5, [0, 10], [white, black]))).toBe('#FFFFFF');
		expect(rgbToCss(interpolateColorScale(50, [0, 10], [white, black]))).toBe('#000000');
	});

	it('interpolates three colors per segment', () => {
		const t = [0, 50, 100];
		const c = [red, yellow, green];
		expect(rgbToCss(interpolateColorScale(50, t, c))).toBe('#FFEB84');
		expect(rgbToCss(interpolateColorScale(25, t, c))).toBe(rgbToCss(lerpRgb(red, yellow, 0.5)));
		expect(rgbToCss(interpolateColorScale(75, t, c))).toBe(rgbToCss(lerpRgb(yellow, green, 0.5)));
		expect(rgbToCss(interpolateColorScale(100, t, c))).toBe('#63BE7B');
		expect(rgbToCss(interpolateColorScale(5, [5, 5, 5], c))).toBe('#F8696B');
	});
});

describe('differential format merging', () => {
	it('lets the higher-priority format win per property', () => {
		const high = { font: { bold: true, color: { rgb: 'FFFF0000' } } };
		const low = {
			font: { italic: true, color: { rgb: 'FF0000FF' } },
			fill: {
				type: 'pattern' as const,
				patternType: 'solid' as const,
				bgColor: { rgb: 'FFFFFF00' },
			},
			border: { left: { style: 'thin' as const } },
		};
		const merged = mergeDifferentialFormats(high, low);
		expect(merged.font).toEqual({ bold: true, color: { rgb: 'FFFF0000' }, italic: true });
		expect(merged.fill).toEqual(low.fill);
		expect(merged.border).toEqual(low.border);
		expect(high).toEqual({ font: { bold: true, color: { rgb: 'FFFF0000' } } });
		expect(mergeDifferentialFormats(undefined, low)).toBe(low);
		expect(mergeDifferentialFormats(undefined, low, true).fill).toBeUndefined();
		expect(mergeDifferentialFormats(high, low, true).fill).toBeUndefined();
	});

	it('reads the DXF background color under both conventions', () => {
		expect(
			dxfBackgroundColor({ type: 'pattern', patternType: 'solid', bgColor: { rgb: 'FF00FF00' } }),
		).toEqual({ rgb: 'FF00FF00' });
		expect(
			dxfBackgroundColor({ type: 'pattern', patternType: 'solid', fgColor: { rgb: 'FF0000FF' } }),
		).toEqual({ rgb: 'FF0000FF' });
		expect(
			dxfBackgroundColor({ type: 'pattern', patternType: 'none', bgColor: { rgb: 'FF00FF00' } }),
		).toBeUndefined();
		expect(dxfBackgroundColor(undefined)).toBeUndefined();
	});
});
