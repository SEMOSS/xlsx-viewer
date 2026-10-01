import { describe, expect, it } from 'vitest';
import type { ColorContext, FillSpec, PatternType } from 'xlsx-viewer-core';
import { createDefaultTheme } from 'xlsx-viewer-core';

import { fillSwatchColor, fillToCss } from './fills';

const ctx: ColorContext = { theme: createDefaultTheme() };

function pattern(patternType: PatternType, fg?: string, bg?: string): FillSpec {
	return {
		type: 'pattern',
		patternType,
		fgColor: fg ? { rgb: fg } : undefined,
		bgColor: bg ? { rgb: bg } : undefined,
	};
}

/** Decoded SVG markup of a `url("data:image/svg+xml,...")` background image. */
function svgOf(image: string | undefined): string {
	const m = /^url\("data:image\/svg\+xml,(.*)"\)$/.exec(image ?? '');
	return m ? decodeURIComponent(m[1]) : '';
}

function rectCount(svg: string): number {
	return svg.match(/<rect /g)?.length ?? 0;
}

describe('fillToCss: pattern fills', () => {
	it('returns nothing for a missing fill or the none pattern', () => {
		expect(fillToCss(undefined, ctx)).toBeUndefined();
		expect(fillToCss(pattern('none', 'FFFF0000'), ctx)).toBeUndefined();
	});

	it('paints solid fills with the foreground color', () => {
		expect(fillToCss(pattern('solid', 'FFFF0000', 'FF00FF00'), ctx)).toEqual({
			backgroundColor: '#FF0000',
		});
		expect(fillToCss(pattern('solid', '00B0F0'), ctx)).toEqual({ backgroundColor: '#00B0F0' });
	});

	it('resolves theme, tinted and indexed colors', () => {
		const themed: FillSpec = { type: 'pattern', patternType: 'solid', fgColor: { theme: 4 } };
		expect(fillToCss(themed, ctx)?.backgroundColor).toBe('#4472C4');
		const tinted: FillSpec = {
			type: 'pattern',
			patternType: 'solid',
			fgColor: { theme: 4, tint: 0.7999 },
		};
		expect(fillToCss(tinted, ctx)?.backgroundColor).toBe('#DAE3F3');
		const indexed: FillSpec = { type: 'pattern', patternType: 'solid', fgColor: { indexed: 13 } };
		expect(fillToCss(indexed, ctx)?.backgroundColor).toBe('#FFFF00');
		const custom = { ...ctx, indexedColors: Array.from({ length: 64 }, () => 'FF123456') };
		expect(fillToCss(indexed, custom)?.backgroundColor).toBe('#123456');
	});

	it('treats a solid fill without a foreground color as black, like Excel', () => {
		expect(fillToCss(pattern('solid'), ctx)).toEqual({ backgroundColor: '#000000' });
	});

	it('draws other patterns as an 8x8 tile of the foreground over the background', () => {
		const css = fillToCss(pattern('gray125', 'FFFF0000', 'FF00FF00'), ctx);
		expect(css?.backgroundColor).toBe('#00FF00');
		const svg = svgOf(css?.backgroundImage);
		expect(svg).toContain('width="8" height="8"');
		expect(svg).toContain('<g fill="#FF0000">');
		// gray125 sets 8 of the 64 pixels; the first one is the top-left corner.
		expect(rectCount(svg)).toBe(8);
		expect(svg).toContain('<rect x="0" y="0" width="1" height="1"/>');
	});

	it('defaults the pattern to black on white', () => {
		const css = fillToCss(pattern('darkGrid'), ctx);
		expect(css?.backgroundColor).toBe('#ffffff');
		expect(svgOf(css?.backgroundImage)).toContain('<g fill="#000000">');
	});

	it('has a tile for every Excel pattern, with the expected density', () => {
		const types: PatternType[] = [
			'gray0625',
			'gray125',
			'lightGray',
			'mediumGray',
			'darkGray',
			'darkHorizontal',
			'darkVertical',
			'darkDown',
			'darkUp',
			'darkGrid',
			'darkTrellis',
			'lightHorizontal',
			'lightVertical',
			'lightDown',
			'lightUp',
			'lightGrid',
			'lightTrellis',
		];
		const counts = new Map(
			types.map((t) => [
				t,
				rectCount(svgOf(fillToCss(pattern(t, 'FF000000'), ctx)?.backgroundImage)),
			]),
		);
		for (const t of types) {
			expect(counts.get(t), t).toBeGreaterThan(0);
			expect(counts.get(t), t).toBeLessThan(64);
		}
		expect(counts.get('gray0625')).toBe(2);
		expect(counts.get('mediumGray')).toBe(32);
		expect(counts.get('lightHorizontal')).toBe(16);
		expect(counts.get('darkHorizontal')).toBe(32);
		// Denser grays have more foreground pixels.
		expect(counts.get('gray0625')!).toBeLessThan(counts.get('gray125')!);
		expect(counts.get('gray125')!).toBeLessThan(counts.get('lightGray')!);
		expect(counts.get('lightGray')!).toBeLessThan(counts.get('mediumGray')!);
		expect(counts.get('mediumGray')!).toBeLessThan(counts.get('darkGray')!);
	});

	it('encodes the pattern color into the tile', () => {
		const red = fillToCss(pattern('lightUp', 'FFFF0000'), ctx)?.backgroundImage;
		const blue = fillToCss(pattern('lightUp', 'FF0000FF'), ctx)?.backgroundImage;
		expect(red).not.toBe(blue);
		expect(fillToCss(pattern('lightUp', 'FFFF0000'), ctx)?.backgroundImage).toBe(red);
		expect(red).toContain('%23FF0000');
	});
});

describe('fillToCss: gradient fills', () => {
	const stops = [
		{ position: 1, color: { rgb: 'FF0000FF' } },
		{ position: 0, color: { rgb: 'FFFF0000' } },
	];

	it('turns linear gradients into CSS with sorted stops (Excel 0 deg is left to right)', () => {
		expect(fillToCss({ type: 'gradient', gradientType: 'linear', stops }, ctx)).toEqual({
			backgroundImage: 'linear-gradient(90deg, #FF0000 0%, #0000FF 100%)',
		});
		const vertical = fillToCss(
			{ type: 'gradient', gradientType: 'linear', degree: 90, stops },
			ctx,
		);
		expect(vertical?.backgroundImage).toBe('linear-gradient(180deg, #FF0000 0%, #0000FF 100%)');
	});

	it('rounds stop positions to percents and falls back to white for automatic colors', () => {
		const css = fillToCss(
			{
				type: 'gradient',
				gradientType: 'linear',
				degree: 45,
				stops: [
					{ position: 0.333, color: { auto: true } },
					{ position: 0, color: { theme: 4 } },
				],
			},
			ctx,
		);
		expect(css?.backgroundImage).toBe('linear-gradient(135deg, #4472C4 0%, #ffffff 33%)');
	});

	it('centres path gradients on the fill rectangle', () => {
		const centred = fillToCss({ type: 'gradient', gradientType: 'path', stops }, ctx);
		expect(centred?.backgroundImage).toBe(
			'radial-gradient(circle at 50% 50%, #FF0000 0%, #0000FF 100%)',
		);
		const corner = fillToCss(
			{ type: 'gradient', gradientType: 'path', left: 1, right: 1, top: 0, bottom: 0, stops },
			ctx,
		);
		expect(corner?.backgroundImage).toMatch(/^radial-gradient\(circle at 100% 0%, /);
	});
});

describe('fillSwatchColor', () => {
	it('reports the single color shown for a fill', () => {
		expect(fillSwatchColor(pattern('solid', 'FF00B050'), ctx)).toBe('#00B050');
		expect(fillSwatchColor(pattern('lightGrid', 'FF000000', 'FFFFFF00'), ctx)).toBe('#FFFF00');
		expect(fillSwatchColor(undefined, ctx)).toBeUndefined();
		expect(
			fillSwatchColor({ type: 'gradient', gradientType: 'linear', stops: [] }, ctx),
		).toBeUndefined();
	});
});
