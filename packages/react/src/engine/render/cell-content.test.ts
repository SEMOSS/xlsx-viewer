import { describe, expect, it } from 'vitest';
import type { Cell, ResolvedCellStyle, StylePatch } from 'xlsx-viewer-core';
import { cellError, createWorkbook, StylePool, StyleResolver } from 'xlsx-viewer-core';

import {
	canOverflow,
	describeCellContent,
	INDENT_PX,
	type CellContent,
	type DescribeOptions,
} from './cell-content';

const wb = createWorkbook();
const pool = new StylePool(wb.styles);

function style(patch: StylePatch = {}): ResolvedCellStyle {
	return new StyleResolver(wb.styles, wb.theme).resolve(pool.derive(undefined, patch));
}

const opts: DescribeOptions = {
	zoom: 1,
	colorCtx: { theme: wb.theme },
	date1904: false,
	defaultTextColor: '#000000',
};

function contentOf(
	cell: Cell | undefined,
	patch: StylePatch = {},
	extra: Partial<DescribeOptions> = {},
): CellContent | null {
	return describeCellContent(cell, style(patch), { ...opts, ...extra });
}

const ACCOUNTING = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)';

describe('describeCellContent: what is drawn', () => {
	it('draws nothing for missing and blank cells', () => {
		expect(contentOf(undefined)).toBeNull();
		expect(contentOf({ value: null })).toBeNull();
		expect(contentOf({ value: '' })).toBeNull();
		const yellow = { type: 'pattern', patternType: 'solid', fgColor: { rgb: 'FFFFFF00' } } as const;
		expect(contentOf({ value: null, style: 1 }, { fill: yellow })).toBeNull();
	});

	it('formats numbers with the number format, right-aligned', () => {
		expect(contentOf({ value: 1234.5 })).toMatchObject({
			kind: 'number',
			text: '1234.5',
			hAlign: 'right',
			vAlign: 'bottom',
			color: '#000000',
			invalid: false,
		});
		expect(contentOf({ value: 1234.5 }, { numFmt: '#,##0.00' })?.text).toBe('1,234.50');
		expect(contentOf({ value: 0.256 }, { numFmt: '0%' })?.text).toBe('26%');
		expect(contentOf({ value: 1 / 3 }, { numFmt: '0.000' })?.text).toBe('0.333');
	});

	it('formats dates and times in both date systems', () => {
		expect(contentOf({ value: 45306 }, { numFmt: 'm/d/yyyy' })?.text).toBe('1/15/2024');
		expect(contentOf({ value: 43844 }, { numFmt: 'm/d/yyyy' }, { date1904: true })?.text).toBe(
			'1/15/2024',
		);
		expect(contentOf({ value: 0.75 }, { numFmt: 'h:mm AM/PM' })?.text).toBe('6:00 PM');
		expect(contentOf({ value: 45306 }, { numFmt: 'm/d/yyyy' })?.kind).toBe('number');
	});

	it('flags values the format cannot show (negative dates)', () => {
		expect(contentOf({ value: -1 }, { numFmt: 'm/d/yyyy' })).toMatchObject({
			kind: 'number',
			text: '#####',
			invalid: true,
		});
	});

	it('takes the color of a format section over the font color', () => {
		const fmt = '[Red]0.00;[Blue]-0.00';
		expect(contentOf({ value: 5 }, { numFmt: fmt })).toMatchObject({
			text: '5.00',
			color: '#FF0000',
		});
		expect(contentOf({ value: -3 }, { numFmt: fmt })).toMatchObject({
			text: '-3.00',
			color: '#0000FF',
		});
		const red = { font: { color: { rgb: 'FF00B050' } } };
		expect(contentOf({ value: 5 }, red)?.color).toBe('#00B050');
		expect(contentOf({ value: 5 }, { ...red, numFmt: fmt })?.color).toBe('#FF0000');
	});

	it('exposes the repeat fill of accounting formats', () => {
		const c = contentOf({ value: 1234.5 }, { numFmt: ACCOUNTING });
		expect(c?.repeatFill).toEqual({ before: ' $', char: ' ', after: '1,234.50 ' });
		expect(c?.text).toBe(' $1,234.50 ');
		expect(c ? canOverflow(c) : true).toBe(false);
	});

	it('draws text left-aligned and applies the text section of the format', () => {
		expect(contentOf({ value: 'Hello' })).toMatchObject({
			kind: 'text',
			text: 'Hello',
			hAlign: 'left',
		});
		expect(contentOf({ value: 'abc' }, { numFmt: '"pre "@' })?.text).toBe('pre abc');
		// Numeric text stays text.
		expect(contentOf({ value: '42' })).toMatchObject({ kind: 'text', hAlign: 'left' });
	});

	it('centres booleans and errors', () => {
		expect(contentOf({ value: true })).toMatchObject({
			kind: 'boolean',
			text: 'TRUE',
			hAlign: 'center',
		});
		expect(contentOf({ value: false })?.text).toBe('FALSE');
		expect(contentOf({ value: cellError('#DIV/0!') })).toMatchObject({
			kind: 'error',
			text: '#DIV/0!',
			hAlign: 'center',
		});
		expect(contentOf({ value: cellError('#N/A') }, { numFmt: '0.00' })?.text).toBe('#N/A');
	});

	it('uses the default text color for automatic font colors', () => {
		const auto = { font: { color: { auto: true } } };
		expect(contentOf({ value: 'x' }, auto, { defaultTextColor: '#E0E0E0' })?.color).toBe('#E0E0E0');
		expect(contentOf({ value: 'x' }, { font: { color: { theme: 4 } } })?.color).toBe('#4472C4');
	});
});

describe('describeCellContent: show formulas and hide zeros', () => {
	const formula: Cell = { value: 3, formula: 'SUM(A1:A2)' };

	it('shows the formula text, left-aligned, when showing formulas', () => {
		expect(contentOf(formula, {}, { showFormulas: true })).toMatchObject({
			kind: 'text',
			text: '=SUM(A1:A2)',
			hAlign: 'left',
		});
		expect(contentOf(formula)).toMatchObject({ kind: 'number', text: '3', hAlign: 'right' });
		// Explicit alignment still applies.
		const right = contentOf(
			formula,
			{ alignment: { horizontal: 'right' } },
			{ showFormulas: true },
		);
		expect(right?.hAlign).toBe('right');
	});

	it('shows formulas that have no value yet only in formula view', () => {
		const pending: Cell = { value: null, formula: 'A1*2' };
		expect(contentOf(pending)).toBeNull();
		expect(contentOf(pending, {}, { showFormulas: true })?.text).toBe('=A1*2');
	});

	it('leaves text constants alone and drops rich text runs of formulas in formula view', () => {
		expect(contentOf({ value: 'abc' }, {}, { showFormulas: true })?.text).toBe('abc');
		const rich: Cell = { value: 'ab', formula: 'X()', richText: [{ text: 'ab' }] };
		expect(contentOf(rich, {}, { showFormulas: true })?.runs).toBeUndefined();
	});

	// Excel's formula view (Ctrl+`) does not apply number formats to constants: a date shows its serial.
	it('shows number constants unformatted in formula view, like Excel', () => {
		const date = contentOf({ value: 45306 }, { numFmt: 'm/d/yyyy' }, { showFormulas: true });
		expect(date?.text).toBe('45306');
	});

	it('hides exact zeros but not formulas shown as text or other values', () => {
		expect(contentOf({ value: 0 }, {}, { hideZeros: true })).toBeNull();
		expect(contentOf({ value: 0 }, { numFmt: '0.00' }, { hideZeros: true })).toBeNull();
		expect(contentOf({ value: 0 }, {}, { hideZeros: false })?.text).toBe('0');
		expect(contentOf({ value: 0.001 }, { numFmt: '0' }, { hideZeros: true })?.text).toBe('0');
		expect(contentOf({ value: '0' }, {}, { hideZeros: true })?.text).toBe('0');
		const zeroFormula: Cell = { value: 0, formula: 'A1-A1' };
		expect(contentOf(zeroFormula, {}, { hideZeros: true })).toBeNull();
		expect(contentOf(zeroFormula, {}, { hideZeros: true, showFormulas: true })?.text).toBe(
			'=A1-A1',
		);
	});
});

describe('describeCellContent: alignment and fonts', () => {
	it('honours explicit horizontal and vertical alignment', () => {
		const c = contentOf({ value: 5 }, { alignment: { horizontal: 'center', vertical: 'top' } });
		expect(c).toMatchObject({ hAlign: 'center', vAlign: 'top' });
		expect(contentOf({ value: 'x' }, { alignment: { horizontal: 'general' } })?.hAlign).toBe(
			'left',
		);
		expect(contentOf({ value: 'x' }, { alignment: { horizontal: 'fill' } })?.hAlign).toBe('fill');
	});

	it('wraps for wrap text, justify and distributed', () => {
		expect(contentOf({ value: 'x' })?.wrap).toBe(false);
		expect(contentOf({ value: 'x' }, { alignment: { wrapText: true } })?.wrap).toBe(true);
		expect(contentOf({ value: 'x' }, { alignment: { horizontal: 'justify' } })?.wrap).toBe(true);
		expect(contentOf({ value: 'x' }, { alignment: { horizontal: 'distributed' } })?.wrap).toBe(
			true,
		);
	});

	it('shrinks to fit only when not wrapping', () => {
		expect(contentOf({ value: 'x' }, { alignment: { shrinkToFit: true } })?.shrink).toBe(true);
		const both = contentOf({ value: 'x' }, { alignment: { shrinkToFit: true, wrapText: true } });
		expect(both).toMatchObject({ shrink: false, wrap: true });
	});

	it('converts indent levels to zoomed pixels', () => {
		expect(contentOf({ value: 'x' }, { alignment: { indent: 2 } })?.indent).toBe(2 * INDENT_PX);
		const zoomed = contentOf({ value: 'x' }, { alignment: { indent: 3 } }, { zoom: 2 });
		expect(zoomed?.indent).toBe(3 * INDENT_PX * 2);
		expect(contentOf({ value: 'x' })?.indent).toBe(0);
	});

	it('maps Excel text rotation to CSS degrees (clockwise positive)', () => {
		const rot = (textRotation: number) =>
			contentOf({ value: 'x' }, { alignment: { textRotation } });
		expect(rot(45)).toMatchObject({ rotation: -45, vertical: false });
		expect(rot(90)?.rotation).toBe(-90);
		expect(rot(135)?.rotation).toBe(45);
		expect(rot(180)?.rotation).toBe(90);
		expect(rot(255)).toMatchObject({ rotation: 0, vertical: true });
		// No rotation comes out as -0, which compares equal to 0 where it is used.
		expect(contentOf({ value: 'x' })?.rotation === 0).toBe(true);
	});

	it('describes the font at the zoom level and accounting underlines', () => {
		const c = contentOf({ value: 'x' }, { font: { size: 12, bold: true } }, { zoom: 2 });
		expect(c?.font).toMatchObject({ fontSize: 32, fontWeight: 700 });
		expect(
			contentOf({ value: 'x' }, { font: { underline: 'singleAccounting' } })?.accountingUnderline,
		).toBe(true);
		expect(contentOf({ value: 'x' }, { font: { underline: 'single' } })?.accountingUnderline).toBe(
			false,
		);
	});

	it('builds rich text runs that inherit the cell font', () => {
		const cell: Cell = {
			value: 'Hello World',
			richText: [
				{ text: 'Hello ' },
				{ text: 'World', font: { bold: true, color: { rgb: 'FFFF0000' } } },
			],
		};
		const c = contentOf(cell, { font: { italic: true, color: { rgb: 'FF0000FF' } } });
		expect(c?.text).toBe('Hello World');
		expect(c?.runs).toHaveLength(2);
		expect(c?.runs?.[0]).toMatchObject({
			text: 'Hello ',
			color: '#0000FF',
			font: { fontWeight: 400, fontStyle: 'italic' },
		});
		expect(c?.runs?.[1]).toMatchObject({
			text: 'World',
			color: '#FF0000',
			font: { fontWeight: 700, fontStyle: 'italic' },
		});
	});
});

describe('canOverflow', () => {
	const text = contentOf({ value: 'long text' })!;

	it('lets plain single-line text spill', () => {
		expect(canOverflow(text)).toBe(true);
		expect(canOverflow({ ...text, hAlign: 'right' })).toBe(true);
		expect(canOverflow({ ...text, hAlign: 'center' })).toBe(true);
	});

	it('keeps numbers, booleans, errors and formatted layouts inside the cell', () => {
		expect(canOverflow({ ...text, kind: 'number' })).toBe(false);
		expect(canOverflow({ ...text, kind: 'boolean' })).toBe(false);
		expect(canOverflow({ ...text, kind: 'error' })).toBe(false);
		expect(canOverflow({ ...text, wrap: true })).toBe(false);
		expect(canOverflow({ ...text, shrink: true })).toBe(false);
		expect(canOverflow({ ...text, vertical: true })).toBe(false);
		expect(canOverflow({ ...text, rotation: 45 })).toBe(false);
		expect(canOverflow({ ...text, repeatFill: { before: '', char: '-', after: '' } })).toBe(false);
		for (const hAlign of ['fill', 'justify', 'distributed', 'centerContinuous'] as const) {
			expect(canOverflow({ ...text, hAlign }), hAlign).toBe(false);
		}
	});
});
