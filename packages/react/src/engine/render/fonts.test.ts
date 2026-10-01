import { describe, expect, it } from 'vitest';

import { canvasFont, fontFamilyStack, fontSizePx, fontToCss } from './fonts';

const CALIBRI = "Calibri, Carlito, 'Segoe UI', 'Helvetica Neue', Arial, sans-serif";

describe('fontFamilyStack', () => {
	it('maps Office fonts to metric-compatible stacks, ignoring case and spaces', () => {
		expect(fontFamilyStack('Calibri')).toBe(CALIBRI);
		expect(fontFamilyStack('  ARIAL ')).toBe(
			"Arial, 'Liberation Sans', Arimo, 'Helvetica Neue', Helvetica, sans-serif",
		);
		expect(fontFamilyStack('Times New Roman')).toBe(
			"'Times New Roman', 'Liberation Serif', Tinos, Times, serif",
		);
		expect(fontFamilyStack('Consolas')).toMatch(/^Consolas, .*monospace$/);
	});

	it('falls back to Calibri without a name', () => {
		expect(fontFamilyStack(undefined)).toBe(CALIBRI);
		expect(fontFamilyStack(undefined, 1)).toBe(CALIBRI);
	});

	it('keeps unknown fonts first and picks the generic family from the family class', () => {
		const tail = "Calibri, Carlito, 'Helvetica Neue', Arial";
		expect(fontFamilyStack('Roboto')).toBe(`Roboto, ${tail}, sans-serif`);
		expect(fontFamilyStack('Roboto', 2)).toBe(`Roboto, ${tail}, sans-serif`);
		expect(fontFamilyStack('Bodoni', 1)).toBe(`Bodoni, ${tail}, serif`);
		expect(fontFamilyStack('Hack', 3)).toBe(`Hack, ${tail}, monospace`);
		expect(fontFamilyStack('Brush', 4)).toBe(`Brush, ${tail}, cursive`);
		expect(fontFamilyStack('Deco', 5)).toBe(`Deco, ${tail}, sans-serif`);
	});

	it('quotes names that are not plain words and escapes apostrophes', () => {
		expect(fontFamilyStack('Source Serif', 1).startsWith("'Source Serif', ")).toBe(true);
		expect(fontFamilyStack('Font2').startsWith("'Font2', ")).toBe(true);
		expect(fontFamilyStack("O'Hare").startsWith("'O\\'Hare', ")).toBe(true);
		expect(fontFamilyStack('Noto-Sans').startsWith('Noto-Sans, ')).toBe(true);
	});
});

describe('fontSizePx', () => {
	it('converts points to CSS pixels at 96 DPI and scales with zoom', () => {
		expect(fontSizePx(12, 1)).toBe(16);
		expect(fontSizePx(11, 1)).toBeCloseTo(14.667, 3);
		expect(fontSizePx(9, 2)).toBe(24);
		expect(fontSizePx(12, 0.5)).toBe(8);
	});

	it('uses 11 pt when the size is missing', () => {
		expect(fontSizePx(undefined, 1)).toBe(fontSizePx(11, 1));
	});
});

describe('fontToCss', () => {
	it('describes a plain font', () => {
		expect(fontToCss({ name: 'Arial', size: 12 }, 1)).toEqual({
			fontFamily: fontFamilyStack('Arial'),
			fontSize: 16,
			fontWeight: 400,
			fontStyle: 'normal',
			textDecorationLine: undefined,
			textDecorationStyle: undefined,
			verticalAlign: undefined,
		});
	});

	it('maps bold, italic and zoom', () => {
		const css = fontToCss({ name: 'Calibri', size: 12, bold: true, italic: true }, 1.5);
		expect(css.fontWeight).toBe(700);
		expect(css.fontStyle).toBe('italic');
		expect(css.fontSize).toBe(24);
	});

	it('combines underline and strikethrough decorations', () => {
		expect(fontToCss({ underline: 'single', strike: true }, 1).textDecorationLine).toBe(
			'underline line-through',
		);
		expect(fontToCss({ strike: true }, 1).textDecorationLine).toBe('line-through');
		expect(fontToCss({ underline: 'none' }, 1).textDecorationLine).toBeUndefined();
	});

	it('draws double and double-accounting underlines as double lines', () => {
		const double = fontToCss({ underline: 'double' }, 1);
		expect(double.textDecorationLine).toBe('underline');
		expect(double.textDecorationStyle).toBe('double');
		expect(fontToCss({ underline: 'doubleAccounting' }, 1).textDecorationStyle).toBe('double');
		const single = fontToCss({ underline: 'singleAccounting' }, 1);
		expect(single.textDecorationLine).toBe('underline');
		expect(single.textDecorationStyle).toBeUndefined();
	});

	it('shrinks superscript and subscript text to 70%', () => {
		const sup = fontToCss({ size: 12, vertAlign: 'superscript' }, 1);
		expect(sup.verticalAlign).toBe('super');
		expect(sup.fontSize).toBeCloseTo(16 * 0.7, 6);
		const sub = fontToCss({ size: 12, vertAlign: 'subscript' }, 1);
		expect(sub.verticalAlign).toBe('sub');
		expect(sub.fontSize).toBeCloseTo(11.2, 6);
		const base = fontToCss({ size: 12, vertAlign: 'baseline' }, 1);
		expect(base.verticalAlign).toBeUndefined();
		expect(base.fontSize).toBe(16);
	});
});

describe('canvasFont', () => {
	it('builds the canvas font shorthand: style, weight, size and family', () => {
		const css = fontToCss({ name: 'Arial', size: 12, bold: true, italic: true }, 1);
		expect(canvasFont(css)).toBe(`italic 700 16px ${fontFamilyStack('Arial')}`);
		expect(canvasFont(fontToCss({ name: 'Calibri', size: 9 }, 1))).toBe(
			`normal 400 12px ${CALIBRI}`,
		);
	});
});
