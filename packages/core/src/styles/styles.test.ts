import { describe, expect, it } from 'vitest';

import { builtinNumberFormatId, numberFormatCode } from './builtin-number-formats';
import {
	columnWidthToPixels,
	displayCharsToStoredWidth,
	pixelsToColumnWidth,
	storedWidthToDisplayChars,
} from './column-width';
import { createDefaultStyleSheet, createDefaultTheme } from './defaults';
import { StyleResolver } from './resolve-style';
import { StylePool } from './style-pool';

describe('column widths (Excel formula, Calibri 11 = 7 px digits)', () => {
	it('converts stored widths and pixels', () => {
		expect(columnWidthToPixels(9.140625)).toBe(64); // the default column
		expect(columnWidthToPixels(8.43)).toBe(59);
		expect(pixelsToColumnWidth(64)).toBeCloseTo(9.140625, 4);
		expect(columnWidthToPixels(pixelsToColumnWidth(100))).toBe(100);
	});

	it('converts between stored widths and the characters Excel shows', () => {
		expect(storedWidthToDisplayChars(9.140625)).toBeCloseTo(8.43, 2);
		expect(displayCharsToStoredWidth(8.43)).toBeCloseTo(9.140625, 2);
	});
});

describe('built-in number formats', () => {
	it('maps ids and codes both ways', () => {
		expect(numberFormatCode(9, {})).toBe('0%');
		expect(numberFormatCode(14, {})).toBe('m/d/yyyy');
		expect(numberFormatCode(164, { 164: '0.000' })).toBe('0.000');
		expect(builtinNumberFormatId('0.00%')).toBe(10);
		expect(builtinNumberFormatId('#,##0.000')).toBeUndefined();
	});
});

describe('StylePool', () => {
	it('derives formats append-only and shares identical ones', () => {
		const styles = createDefaultStyleSheet();
		const pool = new StylePool(styles);
		const xfsBefore = styles.cellXfs.length;
		const fontsBefore = styles.fonts.length;
		const bold = pool.derive(0, { font: { bold: true } });
		expect(bold).toBe(xfsBefore);
		expect(styles.fonts.length).toBe(fontsBefore + 1);
		expect(pool.derive(0, { font: { bold: true } })).toBe(bold);
		const boldRed = pool.derive(bold, { font: { color: { rgb: 'FFFF0000' } } });
		expect(boldRed).not.toBe(bold);
		const resolved = new StyleResolver(styles, createDefaultTheme()).resolve(boldRed);
		expect(resolved.font).toMatchObject({ bold: true, color: { rgb: 'FFFF0000' } });
		// Existing entries never change.
		expect(styles.cellXfs[0]).toEqual(createDefaultStyleSheet().cellXfs[0]);
	});

	it('interns number formats as custom ids from 164', () => {
		const styles = createDefaultStyleSheet();
		const pool = new StylePool(styles);
		const xf = pool.derive(0, { numFmt: '0.000 "kg"' });
		const id = styles.cellXfs[xf].numFmtId;
		expect(id).toBeGreaterThanOrEqual(164);
		expect(pool.derive(0, { numFmt: '0.000 "kg"' })).toBe(xf);
		expect(styles.cellXfs[pool.derive(0, { numFmt: '0%' })].numFmtId).toBe(9);
	});
});
