import { describe, expect, it } from 'vitest';
import { resolveDrawingColorCss } from '../../color/resolve-color';
import { createDefaultTheme } from '../../styles/defaults';
import { autoChartColor, autoMarkerSymbol, chartStyleIndex } from './palette';

const theme = createDefaultTheme();
const css = (i: number, n = 1, style?: number) =>
	resolveDrawingColorCss(autoChartColor(i, n, style), theme);

describe('autoChartColor', () => {
	it('cycles the six accents, then darker and lighter variations (style 2)', () => {
		expect([0, 1, 2, 3, 4, 5].map((i) => css(i))).toEqual([
			'#4472C4',
			'#ED7D31',
			'#A5A5A5',
			'#FFC000',
			'#5B9BD5',
			'#70AD47',
		]);
		expect(autoChartColor(6, 1)).toEqual({
			kind: 'scheme',
			value: 'accent1',
			transforms: [{ name: 'lumMod', val: 60000 }],
		});
		expect(css(6)).toBe('#264478');
		expect(autoChartColor(13, 1)).toEqual({
			kind: 'scheme',
			value: 'accent2',
			transforms: [
				{ name: 'lumMod', val: 80000 },
				{ name: 'lumOff', val: 20000 },
			],
		});
		expect(css(0, 1, 102)).toBe('#4472C4');
	});

	it('uses grays and monochrome shades for the other style columns', () => {
		expect(css(0, 3, 1)).toBe('#595959');
		const mono = [0, 1, 2].map((i) => css(i, 3, 3));
		expect(new Set(mono).size).toBe(3);
		expect(autoChartColor(1, 3, 3)).toMatchObject({ value: 'accent1' });
		expect(autoChartColor(0, 1, 4)).toEqual({ kind: 'scheme', value: 'accent2' });
	});
});

describe('chartStyleIndex / autoMarkerSymbol', () => {
	it('normalises c14 styles and cycles markers like Excel', () => {
		expect(chartStyleIndex(102)).toBe(2);
		expect(chartStyleIndex(undefined)).toBe(2);
		expect(chartStyleIndex(99)).toBe(2);
		expect([0, 1, 2, 3, 4, 5, 6, 7].map(autoMarkerSymbol)).toEqual([
			'diamond',
			'square',
			'triangle',
			'x',
			'star',
			'circle',
			'plus',
			'diamond',
		]);
	});
});
