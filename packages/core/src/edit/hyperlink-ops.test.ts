import { describe, expect, it } from 'vitest';

import { cellKey } from '../address/cell-address';
import { parseRange } from '../address/range';
import { createWorkbook } from '../sheet/workbook';
import { getCell } from '../sheet/worksheet';
import { createDefaultStyleSheet } from '../styles/defaults';
import { StylePool } from '../styles/style-pool';
import type { Workbook } from '../types/workbook';
import {
	EditTransaction,
	ensureHyperlinkCellStyle,
	findHyperlinkCellStyle,
	hyperlinkAt,
	hyperlinkCellXf,
	hyperlinkDisplayText,
	mergeRange,
	removeHyperlinks,
	setCellInput,
	setHyperlink,
} from './index';

const r = (a1: string) => parseRange(a1)!;

function seed(wb: Workbook, entries: Record<string, string>): void {
	const tx = new EditTransaction(wb, 'seed');
	for (const [ref, text] of Object.entries(entries)) {
		setCellInput(tx, wb, 0, r(ref).top, r(ref).left, text);
	}
	tx.commit();
}

const cellAt = (wb: Workbook, a1: string) => getCell(wb.sheets[0], r(a1).top, r(a1).left);

describe('Hyperlink cell style', () => {
	it('creates the built-in style once, appending only', () => {
		const styles = createDefaultStyleSheet();
		const before = JSON.stringify({
			fonts: styles.fonts,
			xfs: styles.cellStyleXfs,
			cs: styles.cellStyles,
		});
		expect(findHyperlinkCellStyle(styles)).toBeUndefined();
		const xfId = ensureHyperlinkCellStyle(styles);
		expect(xfId).toBe(1);
		expect(styles.cellStyles[1]).toEqual({ name: 'Hyperlink', xfId: 1, builtinId: 8 });
		const font = styles.fonts[styles.cellStyleXfs[1].fontId];
		expect(font).toMatchObject({
			name: 'Calibri',
			size: 11,
			color: { theme: 10 },
			underline: 'single',
		});
		expect(styles.cellStyleXfs[1]).toMatchObject({
			numFmtId: 0,
			fillId: 0,
			borderId: 0,
			applyFill: false,
		});
		const after = JSON.parse(
			JSON.stringify({ fonts: styles.fonts, xfs: styles.cellStyleXfs, cs: styles.cellStyles }),
		);
		const old = JSON.parse(before);
		expect(after.fonts.slice(0, old.fonts.length)).toEqual(old.fonts);
		expect(after.xfs.slice(0, old.xfs.length)).toEqual(old.xfs);
		expect(ensureHyperlinkCellStyle(styles)).toBe(1);
		expect(styles.cellStyles).toHaveLength(2);
	});

	it('reuses an existing Hyperlink style by built-in id or name', () => {
		const styles = createDefaultStyleSheet();
		styles.cellStyleXfs.push({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 });
		styles.cellStyles.push({ name: 'Link', xfId: 1, builtinId: 8 });
		expect(ensureHyperlinkCellStyle(styles)).toBe(1);
		const named = createDefaultStyleSheet();
		named.cellStyleXfs.push({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 });
		named.cellStyles.push({ name: 'hyperlink', xfId: 1 });
		expect(ensureHyperlinkCellStyle(named)).toBe(1);
		expect(named.cellStyles).toHaveLength(2);
	});

	it('shares one cell format based on the style', () => {
		const styles = createDefaultStyleSheet();
		const xf = hyperlinkCellXf(styles);
		expect(styles.cellXfs[xf]).toEqual({
			numFmtId: 0,
			fontId: styles.cellStyleXfs[1].fontId,
			fillId: 0,
			borderId: 0,
			xfId: 1,
			applyFont: true,
		});
		expect(hyperlinkCellXf(styles)).toBe(xf);
		expect(styles.cellXfs).toHaveLength(2);
	});
});

describe('setHyperlink', () => {
	it('links a cell, writes the text, applies the style and undoes', () => {
		const wb = createWorkbook();
		const tx = new EditTransaction(wb, 'Insert Hyperlink');
		const link = setHyperlink(
			tx,
			wb,
			0,
			r('B2'),
			{ target: 'http://example.com', tooltip: 'Go' },
			'Example',
		);
		const entry = tx.commit();
		expect(link).toEqual({
			ref: r('B2'),
			target: 'http://example.com',
			tooltip: 'Go',
			display: 'Example',
		});
		expect(wb.sheets[0].hyperlinks).toEqual([link]);
		const cell = cellAt(wb, 'B2')!;
		expect(cell.value).toBe('Example');
		expect(wb.styles.cellXfs[cell.style!].xfId).toBe(findHyperlinkCellStyle(wb.styles));
		expect(entry.changedCells).toContainEqual({ sheet: 0, row: 1, col: 1 });

		entry.undo();
		expect(wb.sheets[0].hyperlinks).toEqual([]);
		expect(cellAt(wb, 'B2')).toBeUndefined();
		entry.redo();
		expect(hyperlinkAt(wb.sheets[0], 1, 1)?.target).toBe('http://example.com');
		expect(cellAt(wb, 'B2')?.value).toBe('Example');
	});

	it('keeps text that already shows the display text and stores new text as text', () => {
		const wb = createWorkbook();
		seed(wb, { A3: 'old' });
		const fmt = new StylePool(wb.styles).derive(undefined, { numFmt: '#,##0.00' });
		wb.sheets[0].cells.set(cellKey(0, 0), { value: 1234.5, style: fmt });
		wb.sheets[0].cells.set(cellKey(1, 0), { value: 2, formula: '1+1' });
		const tx = new EditTransaction(wb, 'Insert Hyperlink');
		setHyperlink(tx, wb, 0, r('A1'), { location: 'Sheet1!C1' }, '1,234.50');
		setHyperlink(tx, wb, 0, r('A2'), { location: 'Sheet1!C2' }, '2');
		setHyperlink(tx, wb, 0, r('A3'), { location: 'Sheet1!C3' }, '123');
		tx.commit();
		expect(cellAt(wb, 'A1')).toEqual({ value: 1234.5, style: fmt });
		expect(cellAt(wb, 'A2')?.formula).toBe('1+1');
		expect(cellAt(wb, 'A3')?.value).toBe('123');
		expect(cellAt(wb, 'A3')?.formula).toBeUndefined();
	});

	it('replaces overlapping links and keeps their other cells linked', () => {
		const wb = createWorkbook();
		wb.sheets[0].hyperlinks.push(
			{ ref: r('A1:A3'), target: 'http://old.example' },
			{ ref: r('D1'), location: 'Sheet1!Z1' },
		);
		const tx = new EditTransaction(wb, 'Insert Hyperlink');
		setHyperlink(tx, wb, 0, r('A2'), { target: 'http://new.example' });
		const links = wb.sheets[0].hyperlinks;
		expect(links.map((h) => `${h.ref.top}-${h.ref.bottom}:${h.target ?? h.location}`)).toEqual([
			'0-0:http://old.example',
			'2-2:http://old.example',
			'0-0:Sheet1!Z1',
			'1-1:http://new.example',
		]);
		expect(hyperlinkAt(wb.sheets[0], 1, 0)?.target).toBe('http://new.example');
		expect(cellAt(wb, 'A2')?.value).toBeNull();
	});

	it('styles unformatted cells of the range only, skipping merged-away cells', () => {
		const wb = createWorkbook();
		const bold = new StylePool(wb.styles).derive(undefined, { font: { bold: true } });
		wb.sheets[0].cells.set(cellKey(1, 1), { value: 'keep', style: bold });
		const setup = new EditTransaction(wb, 'merge');
		mergeRange(setup, wb, 0, r('C1:C2'));
		setup.commit();
		const tx = new EditTransaction(wb, 'Insert Hyperlink');
		setHyperlink(tx, wb, 0, r('A1:C2'), { target: 'http://x.example' });
		const xf = hyperlinkCellXf(wb.styles);
		expect(cellAt(wb, 'A1')?.style).toBe(xf);
		expect(cellAt(wb, 'B1')?.style).toBe(xf);
		expect(cellAt(wb, 'A2')?.style).toBe(xf);
		expect(cellAt(wb, 'B2')?.style).toBe(bold);
		expect(cellAt(wb, 'C1')?.style).toBe(xf);
		expect(cellAt(wb, 'C2')).toBeUndefined();
	});

	it('styles only existing cells and the top-left cell of a whole column', () => {
		const wb = createWorkbook();
		seed(wb, { B5: 'x', C5: 'outside' });
		const tx = new EditTransaction(wb, 'Insert Hyperlink');
		setHyperlink(tx, wb, 0, r('B:B'), { target: 'http://x.example' }, 'Top');
		const xf = hyperlinkCellXf(wb.styles);
		expect(cellAt(wb, 'B1')).toEqual({ value: 'Top', style: xf });
		expect(cellAt(wb, 'B5')?.style).toBe(xf);
		expect(cellAt(wb, 'C5')?.style).toBeUndefined();
		expect(wb.sheets[0].cells.size).toBe(3);
	});
});

describe('removeHyperlinks', () => {
	function linked(): Workbook {
		const wb = createWorkbook();
		seed(wb, { A1: 'Home', A2: 'Docs', B1: 'plain' });
		const tx = new EditTransaction(wb, 'Insert Hyperlink');
		setHyperlink(tx, wb, 0, r('A1:A3'), { target: 'http://example.com' });
		setHyperlink(tx, wb, 0, r('C1'), { location: 'Sheet1!A1' });
		tx.commit();
		return wb;
	}

	it('unlinks the range and returns the cells to the Normal style', () => {
		const wb = linked();
		const tx = new EditTransaction(wb, 'Remove Hyperlinks');
		expect(removeHyperlinks(tx, wb, 0, [r('A1:B3')])).toBe(1);
		const entry = tx.commit();
		expect(wb.sheets[0].hyperlinks.map((h) => h.location)).toEqual(['Sheet1!A1']);
		expect(cellAt(wb, 'A1')).toEqual({ value: 'Home' });
		expect(cellAt(wb, 'A2')).toEqual({ value: 'Docs' });
		expect(cellAt(wb, 'A3')).toBeUndefined();
		expect(cellAt(wb, 'B1')).toEqual({ value: 'plain' });
		expect(cellAt(wb, 'C1')?.style).toBe(hyperlinkCellXf(wb.styles));

		entry.undo();
		expect(wb.sheets[0].hyperlinks).toHaveLength(2);
		expect(cellAt(wb, 'A1')?.style).toBe(hyperlinkCellXf(wb.styles));
		entry.redo();
		expect(wb.sheets[0].hyperlinks).toHaveLength(1);
	});

	it('keeps the rest of a partly removed link and other formats', () => {
		const wb = linked();
		const bold = new StylePool(wb.styles).derive(undefined, { font: { bold: true } });
		wb.sheets[0].cells.set(cellKey(1, 0), { value: 'Docs', style: bold });
		const tx = new EditTransaction(wb, 'Remove Hyperlinks');
		expect(removeHyperlinks(tx, wb, 0, [r('A2'), r('Z9')])).toBe(1);
		expect(wb.sheets[0].hyperlinks.map((h) => h.ref)).toEqual([r('A1'), r('A3'), r('C1')]);
		expect(cellAt(wb, 'A2')?.style).toBe(bold);
		expect(removeHyperlinks(tx, wb, 0, [r('Z9')])).toBe(0);
	});
});

describe('hyperlinkAt / hyperlinkDisplayText', () => {
	it('finds the link of a cell and its displayed text', () => {
		const wb = createWorkbook();
		seed(wb, { A1: '7%', A2: 'text' });
		wb.sheets[0].hyperlinks.push({ ref: r('A1:B2'), target: 'a' }, { ref: r('B2'), target: 'b' });
		expect(hyperlinkAt(wb.sheets[0], 0, 1)?.target).toBe('a');
		expect(hyperlinkAt(wb.sheets[0], 1, 1)?.target).toBe('b');
		expect(hyperlinkAt(wb.sheets[0], 5, 5)).toBeUndefined();
		expect(hyperlinkDisplayText(wb, 0, 0, 0)).toBe('7%');
		expect(hyperlinkDisplayText(wb, 0, 1, 0)).toBe('text');
		expect(hyperlinkDisplayText(wb, 0, 9, 9)).toBe('');
	});
});
