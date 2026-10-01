import { describe, expect, it } from 'vitest';
import {
	baselineOffset,
	blockWidth,
	defaultMeasureText,
	ellipsize,
	lineHeightFor,
	wrapText,
} from './text';

const font = { family: 'Calibri', size: 10 };
const m = defaultMeasureText;

describe('defaultMeasureText', () => {
	it('uses Calibri advance widths', () => {
		expect(m('0', font)).toBeCloseTo(5.07, 2);
		expect(m('0123456789', font)).toBeCloseTo(50.7, 1);
		expect(m('iiii', font)).toBeLessThan(m('WWWW', font));
		expect(m('', font)).toBe(0);
	});

	it('accounts for weight, family and wide characters', () => {
		expect(m('Sales', { ...font, bold: true })).toBeGreaterThan(m('Sales', font));
		expect(m('Sales', { ...font, family: 'Arial' })).toBeGreaterThan(m('Sales', font));
		expect(m('漢字', font)).toBeCloseTo(20, 5);
	});
});

describe('wrapText / ellipsize', () => {
	it('wraps words to the width', () => {
		expect(wrapText('North America East', 50, font, m)).toEqual(['North', 'America', 'East']);
		expect(wrapText('Q1 Q2', 100, font, m)).toEqual(['Q1 Q2']);
		expect(wrapText('a\nb', 100, font, m)).toEqual(['a', 'b']);
	});

	it('breaks words longer than a line', () => {
		const lines = wrapText('Supercalifragilistic', 30, font, m);
		expect(lines.length).toBeGreaterThan(2);
		expect(lines.join('')).toBe('Supercalifragilistic');
		for (const l of lines) {
			expect(m(l, font)).toBeLessThanOrEqual(30);
		}
	});

	it('limits the number of lines with an ellipsis', () => {
		const lines = wrapText('one two three four five six', 30, font, m, 2);
		expect(lines).toHaveLength(2);
		expect(lines[1].endsWith('…')).toBe(true);
	});

	it('ellipsizes to fit', () => {
		expect(ellipsize('Short', 100, font, m)).toBe('Short');
		const cut = ellipsize('A rather long category label', 60, font, m);
		expect(cut.endsWith('…')).toBe(true);
		expect(m(cut, font)).toBeLessThanOrEqual(60);
		expect(ellipsize('Wide', 1, font, m)).toBe('');
	});

	it('measures blocks and line metrics', () => {
		expect(blockWidth(['a', 'abc'], font, m)).toBe(m('abc', font));
		expect(lineHeightFor(10)).toBeCloseTo(12.2);
		expect(baselineOffset(10)).toBeCloseTo(9.1);
	});
});
