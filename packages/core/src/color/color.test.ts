import { describe, expect, it } from 'vitest';

import { createDefaultTheme } from '../styles/defaults';
import { applyDrawingTransform, applyTint, parseHex, toHex } from './color-math';
import {
	resolveColorCss,
	resolveColorHex,
	resolveDrawingColorCss,
	themeColorByIndex,
} from './resolve-color';

const theme = createDefaultTheme();
const ctx = { theme };

/** Channel-wise distance between two RRGGBB colors. */
function distance(a: string, b: string): number {
	const x = parseHex(a)!;
	const y = parseHex(b)!;
	return Math.max(Math.abs(x.r - y.r), Math.abs(x.g - y.g), Math.abs(x.b - y.b));
}

describe('SpreadsheetML colors', () => {
	it('maps theme indexes with the light / dark swap Excel uses', () => {
		// Index 0 is lt1 (background) and 1 is dk1 (text), the reverse of the scheme order.
		expect(themeColorByIndex(theme, 0)).toBe(theme.colors.lt1);
		expect(themeColorByIndex(theme, 1)).toBe(theme.colors.dk1);
		expect(themeColorByIndex(theme, 4)).toBe(theme.colors.accent1);
	});

	it('applies tints like the Excel palette (Office theme, Blue Accent 1)', () => {
		const accent = { theme: 4 };
		const cases: Array<[number, string]> = [
			[0.7999816888943144, 'D9E1F2'],
			[0.5999938962981048, 'B4C6E7'],
			[0.3999755851924192, '8EA9DB'],
			[-0.249977111117893, '2F5597'],
			[-0.499984740745262, '203764'],
		];
		for (const [tint, expected] of cases) {
			expect(distance(resolveColorHex({ ...accent, tint }, ctx)!, expected)).toBeLessThanOrEqual(1);
		}
	});

	it('resolves rgb, indexed and automatic colors', () => {
		expect(resolveColorHex({ rgb: 'FF336699' }, ctx)).toBe('336699');
		expect(resolveColorHex({ indexed: 10 }, ctx)).toBe('FF0000');
		expect(resolveColorHex({ indexed: 64 }, ctx)).toBe('000000');
		expect(resolveColorHex({ auto: true }, ctx)).toBeUndefined();
		expect(resolveColorCss(undefined, ctx, '#123456')).toBe('#123456');
		expect(
			resolveColorHex(
				{ indexed: 8 },
				{
					theme,
					indexedColors: [
						'FF000000',
						'FFFFFFFF',
						'FF00FF00',
						'FF000000',
						'FF000000',
						'FF000000',
						'FF000000',
						'FF000000',
						'FFABCDEF',
					],
				},
			),
		).toBe('ABCDEF');
	});

	it('keeps colors stable without a tint and clamps extremes', () => {
		expect(toHex(applyTint(parseHex('4472C4')!, 0))).toBe('4472C4');
		expect(toHex(applyTint(parseHex('4472C4')!, 1))).toBe('FFFFFF');
		expect(toHex(applyTint(parseHex('4472C4')!, -1))).toBe('000000');
	});
});

describe('DrawingML colors', () => {
	it('applies lumMod / lumOff like Office (accent1 lighter 40% = lumMod 60000 lumOff 40000)', () => {
		const css = resolveDrawingColorCss(
			{
				kind: 'scheme',
				value: 'accent1',
				transforms: [
					{ name: 'lumMod', val: 60000 },
					{ name: 'lumOff', val: 40000 },
				],
			},
			theme,
		);
		expect(css).toBeDefined();
		expect(
			distance(css!.replace(/^#/, '').slice(0, 6).toUpperCase(), '8FAADC'),
		).toBeLessThanOrEqual(2);
	});

	it('applies alpha', () => {
		const c = applyDrawingTransform({ r: 10, g: 20, b: 30, a: 1 }, 'alpha', 50000);
		expect(c.a).toBeCloseTo(0.5);
	});
});
