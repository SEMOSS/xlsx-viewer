import { describe, expect, it } from 'vitest';

import {
	formatHeaderDate,
	formatHeaderFooter,
	formatHeaderTime,
	headerFooterFor,
	headerFooterText,
	isHeaderFooterEmpty,
	joinHeaderFooter,
	splitHeaderFooter,
	type HeaderFooterContext,
} from './header-footer';

const DATE = new Date(2026, 8, 30, 14, 5);
const CTX: HeaderFooterContext = {
	page: 3,
	pages: 7,
	sheetName: 'Budget',
	fileName: 'Report.xlsx',
	filePath: 'C:\\Docs\\Report.xlsx',
	date: DATE,
};

const text = (code: string, ctx: HeaderFooterContext = CTX) =>
	headerFooterText(formatHeaderFooter(code, ctx));

describe('formatHeaderFooter sections and fields', () => {
	it('puts text without a section code in the center', () => {
		expect(text('Hello')).toEqual({ left: '', center: 'Hello', right: '' });
	});

	it('splits &L, &C and &R sections', () => {
		expect(text('&LLeft&CMiddle&RRight')).toEqual({
			left: 'Left',
			center: 'Middle',
			right: 'Right',
		});
		expect(text('&RA&LB&RC')).toEqual({ left: 'B', center: '', right: 'AC' });
	});

	it('replaces page fields, including &P+n and &P-n', () => {
		expect(text('Page &P of &N').center).toBe('Page 3 of 7');
		expect(text('&P+10').center).toBe('13');
		expect(text('&P-1').center).toBe('2');
		expect(text('&P-x').center).toBe('3-x');
	});

	it('replaces date, time, file, path and sheet fields', () => {
		expect(text('&D').center).toBe('9/30/2026');
		expect(text('&T').center).toBe('2:05 PM');
		expect(text('&F').center).toBe('Report.xlsx');
		expect(text('&Z').center).toBe('C:\\Docs\\Report.xlsx');
		expect(text('&A').center).toBe('Budget');
		expect(formatHeaderDate(new Date(2024, 0, 5))).toBe('1/5/2024');
		expect(formatHeaderTime(new Date(2024, 0, 5, 0, 7))).toBe('12:07 AM');
		expect(formatHeaderTime(new Date(2024, 0, 5, 12, 0))).toBe('12:00 PM');
	});

	it('keeps literal ampersands and skips pictures and unknown codes', () => {
		expect(text('R&&D').center).toBe('R&D');
		expect(text('&GLogo&O&H').center).toBe('Logo');
		expect(text('end&').center).toBe('end');
	});

	it('formats nothing for an empty code', () => {
		expect(isHeaderFooterEmpty(formatHeaderFooter('', CTX))).toBe(true);
		expect(isHeaderFooterEmpty(formatHeaderFooter(undefined, CTX))).toBe(true);
		expect(isHeaderFooterEmpty(null)).toBe(true);
	});
});

describe('formatHeaderFooter fonts', () => {
	it('reads &"Font,Style" and &nn sizes into runs', () => {
		const s = formatHeaderFooter('&C&"Arial,Bold Italic"&14Title&"-,Regular" plain', CTX);
		expect(s.center).toEqual([
			{ text: 'Title', font: { name: 'Arial', bold: true, italic: true, size: 14 } },
			{ text: ' plain', font: { name: 'Arial', size: 14 } },
		]);
	});

	it('toggles bold, italic, underline, double underline, strike, superscript and subscript', () => {
		const s = formatHeaderFooter('&Bb&B &Ii&I &Uu&U &Ee&E &Ss&S &Xx&X &Yy&Y', CTX);
		expect(s.center.map((r) => [r.text, r.font])).toEqual([
			['b', { bold: true }],
			[' ', {}],
			['i', { italic: true }],
			[' ', {}],
			['u', { underline: 'single' }],
			[' ', {}],
			['e', { underline: 'double' }],
			[' ', {}],
			['s', { strike: true }],
			[' ', {}],
			['x', { vertAlign: 'superscript' }],
			[' ', {}],
			['y', { vertAlign: 'subscript' }],
		]);
	});

	it('reads RGB and theme colors', () => {
		const s = formatHeaderFooter('&KFF0000red&K01+050tint&K03-025dark&K02+000plain', CTX);
		expect(s.center.map((r) => r.font.color)).toEqual([
			{ rgb: 'FFFF0000' },
			{ theme: 1, tint: 0.5 },
			{ theme: 3, tint: -0.25 },
			{ theme: 2 },
		]);
	});

	it('resets the font at each section', () => {
		const s = formatHeaderFooter('&L&BBold&RPlain', CTX);
		expect(s.left[0].font).toEqual({ bold: true });
		expect(s.right[0].font).toEqual({});
	});

	it('merges adjacent text with the same font', () => {
		const s = formatHeaderFooter('Page &P', CTX);
		expect(s.center).toEqual([{ text: 'Page 3', font: {} }]);
	});
});

describe('header / footer helpers', () => {
	it('splits and joins raw sections', () => {
		expect(splitHeaderFooter('&L&"Calibri,Bold"Left &&C&CMid&RRight')).toEqual({
			left: '&"Calibri,Bold"Left &&C',
			center: 'Mid',
			right: 'Right',
		});
		expect(splitHeaderFooter('Page &P')).toEqual({ left: '', center: 'Page &P', right: '' });
		expect(splitHeaderFooter(undefined)).toEqual({ left: '', center: '', right: '' });
		expect(joinHeaderFooter({ left: 'A', center: '', right: '&P' })).toBe('&LA&R&P');
		expect(joinHeaderFooter({})).toBe('');
	});

	it('picks the first, even or odd variant', () => {
		const hf = {
			oddHeader: 'odd',
			evenHeader: 'even',
			firstHeader: 'first',
			oddFooter: 'oddF',
			differentFirst: true,
			differentOddEven: true,
		};
		expect(headerFooterFor(hf, 0, 1)).toEqual({ header: 'first', footer: undefined });
		expect(headerFooterFor(hf, 1, 2)).toEqual({ header: 'even', footer: undefined });
		expect(headerFooterFor(hf, 2, 3)).toEqual({ header: 'odd', footer: 'oddF' });
		expect(headerFooterFor({ ...hf, differentFirst: false }, 0, 1).header).toBe('odd');
		expect(headerFooterFor({ ...hf, differentOddEven: false }, 1, 2).header).toBe('odd');
		expect(headerFooterFor(undefined, 0, 1)).toEqual({});
	});
});
