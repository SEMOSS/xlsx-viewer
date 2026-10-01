import { createDefaultTheme, getBuiltinTableStyle } from 'xlsx-viewer-core';
import { describe, expect, it } from 'vitest';

import { buildTableThumbnail } from './table-thumbnail';

const ctx = { theme: createDefaultTheme() };

describe('buildTableThumbnail', () => {
	it('draws a 5 x 5 preview with a filled header for Medium 2', () => {
		const thumb = buildTableThumbnail(getBuiltinTableStyle('TableStyleMedium2')!, {}, ctx);
		expect(thumb.cells).toHaveLength(25);
		expect(thumb.width).toBe(55);
		const header = thumb.cells.filter((c) => c.y === 0);
		expect(header.every((c) => c.fill && c.bold)).toBe(true);
		// Row stripes: the first data row is filled, the second is not.
		expect(thumb.cells.find((c) => c.y === 8 && c.x === 0)?.fill).toBeDefined();
		expect(thumb.cells.find((c) => c.y === 16 && c.x === 0)?.fill).toBeUndefined();
	});

	it('follows the stripe options', () => {
		const style = getBuiltinTableStyle('TableStyleMedium2')!;
		const thumb = buildTableThumbnail(style, { showRowStripes: false }, ctx);
		expect(thumb.cells.filter((c) => c.y > 0).every((c) => !c.fill)).toBe(true);
	});

	it('emits border lines for light styles', () => {
		const thumb = buildTableThumbnail(getBuiltinTableStyle('TableStyleLight9')!, {}, ctx);
		expect(thumb.lines.length).toBeGreaterThan(0);
	});
});
