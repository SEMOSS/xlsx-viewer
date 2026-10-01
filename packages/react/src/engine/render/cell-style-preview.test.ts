import { createWorkbook, ensureCellStyle } from 'xlsx-viewer-core';
import { describe, expect, it } from 'vitest';

import { cellStylePreview } from './cell-style-preview';

describe('cellStylePreview', () => {
	it('previews built-in styles the workbook does not have yet', () => {
		const wb = createWorkbook();
		const good = cellStylePreview(wb.styles, wb.theme, 'Good')!;
		expect(good.color).toBe('#006100');
		expect(good.backgroundColor?.toLowerCase()).toBe('#c6efce');
		const h1 = cellStylePreview(wb.styles, wb.theme, 'Heading 1')!;
		expect(h1.fontWeight).toBe(700);
		expect(h1.fontSize).toBe(20);
		expect(h1.borderBottom).toMatch(/^3px solid/);
		expect(cellStylePreview(wb.styles, wb.theme, 'Nope')).toBeUndefined();
	});

	it('uses the workbook definition once the style exists', () => {
		const wb = createWorkbook();
		const id = ensureCellStyle(wb.styles, 'Note', wb.theme)!;
		// Change the workbook's Note style fill: the preview follows it.
		wb.styles.fills.push({ type: 'pattern', patternType: 'solid', fgColor: { rgb: 'FF00FF00' } });
		wb.styles.cellStyleXfs[id] = {
			...wb.styles.cellStyleXfs[id],
			fillId: wb.styles.fills.length - 1,
		};
		expect(cellStylePreview(wb.styles, wb.theme, 'Note')?.backgroundColor?.toLowerCase()).toBe(
			'#00ff00',
		);
	});
});
