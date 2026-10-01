import { describe, expect, it } from 'vitest';
import type { FillSpec, ResolvedCellStyle } from 'xlsx-viewer-core';
import { createWorkbook, StylePool, StyleResolver } from 'xlsx-viewer-core';

import { applyDxf } from './apply-dxf';

const RED = { rgb: 'FFFF0000' };
const BLUE = { rgb: 'FF0000FF' };
const solid = (fgColor = RED): FillSpec => ({ type: 'pattern', patternType: 'solid', fgColor });

/** A resolved style: italic Arial 12, yellow fill, thin top and left borders, centred, "0.00". */
function base(): ResolvedCellStyle {
	const wb = createWorkbook();
	const id = new StylePool(wb.styles).derive(undefined, {
		// Like the font-name command: a named font no longer follows the theme's body font.
		font: { name: 'Arial', scheme: undefined, size: 12, italic: true },
		fill: solid({ rgb: 'FFFFFF00' }),
		border: { top: { style: 'thin' }, left: { style: 'thin', color: BLUE } },
		alignment: { horizontal: 'center', wrapText: true },
		numFmt: '0.00',
	});
	return new StyleResolver(wb.styles, wb.theme).resolve(id);
}

describe('applyDxf', () => {
	it('overrides only the font properties the dxf sets', () => {
		const style = base();
		const next = applyDxf(style, { font: { bold: true, color: RED, italic: undefined } });
		expect(next.font).toMatchObject({
			name: 'Arial',
			size: 12,
			italic: true,
			bold: true,
			color: RED,
		});
		expect(next.fill).toBe(style.fill);
		expect(next.border).toBe(style.border);
	});

	it('never mutates the base style', () => {
		const style = base();
		const snapshot = JSON.stringify(style);
		applyDxf(style, {
			font: { bold: true },
			fill: solid(),
			border: { top: { style: 'thick' } },
			alignment: { vertical: 'top' },
			numFmt: { id: 10, code: '0.00%' },
		});
		expect(JSON.stringify(style)).toBe(snapshot);
	});

	it('reads the color of solid dxf fills from bgColor, as Excel stores it', () => {
		const next = applyDxf(base(), {
			fill: { type: 'pattern', patternType: 'solid', bgColor: { rgb: 'FFFFC7CE' } },
		});
		expect(next.fill).toEqual({
			type: 'pattern',
			patternType: 'solid',
			fgColor: { rgb: 'FFFFC7CE' },
		});
		// bgColor wins when both are present; fgColor is used when it is the only color.
		const both = applyDxf(base(), {
			fill: { type: 'pattern', patternType: 'solid', fgColor: BLUE, bgColor: RED },
		});
		expect(both.fill).toEqual(solid(RED));
		expect(applyDxf(base(), { fill: solid(BLUE) }).fill).toEqual(solid(BLUE));
	});

	it('treats colored none patterns as solid and keeps the cell fill for colorless ones', () => {
		const style = base();
		const colored = applyDxf(style, {
			fill: { type: 'pattern', patternType: 'none', bgColor: RED },
		});
		expect(colored.fill).toEqual(solid(RED));
		const empty = applyDxf(style, { fill: { type: 'pattern', patternType: 'none' } });
		expect(empty.fill).toBe(style.fill);
	});

	it('keeps other patterns and gradients as they are', () => {
		const grid: FillSpec = {
			type: 'pattern',
			patternType: 'lightGrid',
			fgColor: RED,
			bgColor: BLUE,
		};
		expect(applyDxf(base(), { fill: grid }).fill).toBe(grid);
		const gradient: FillSpec = {
			type: 'gradient',
			gradientType: 'linear',
			degree: 90,
			stops: [
				{ position: 0, color: RED },
				{ position: 1, color: BLUE },
			],
		};
		expect(applyDxf(base(), { fill: gradient }).fill).toBe(gradient);
	});

	it('overrides border edges one by one', () => {
		const next = applyDxf(base(), {
			border: { top: { style: 'thick', color: RED }, bottom: { style: 'dashed' } },
		});
		expect(next.border).toEqual({
			top: { style: 'thick', color: RED },
			left: { style: 'thin', color: BLUE },
			bottom: { style: 'dashed' },
		});
	});

	it('replaces the number format code and id', () => {
		const style = base();
		expect(style.numFmt).toBe('0.00');
		const next = applyDxf(style, { numFmt: { id: 164, code: '[Red]0.0' } });
		expect([next.numFmt, next.numFmtId]).toEqual(['[Red]0.0', 164]);
	});

	it('merges alignment', () => {
		const next = applyDxf(base(), { alignment: { vertical: 'top', horizontal: 'right' } });
		expect(next.alignment).toEqual({ horizontal: 'right', vertical: 'top', wrapText: true });
	});

	it('returns an equal copy for an empty dxf', () => {
		const style = base();
		const next = applyDxf(style, {});
		expect(next).not.toBe(style);
		expect(next).toEqual(style);
	});
});
