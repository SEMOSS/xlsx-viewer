import { createWorksheet } from './worksheet';
import { describe, expect, it } from 'vitest';

import { computeOutline } from './outline';

describe('computeOutline', () => {
	it('finds nested row groups with summary rows below', () => {
		const sheet = createWorksheet('S', 1);
		// Rows 1-3 at level 1, rows 2-3 also at level 2; row 4 is the summary.
		sheet.rows.set(1, { outlineLevel: 1 });
		sheet.rows.set(2, { outlineLevel: 2, hidden: true });
		sheet.rows.set(3, { outlineLevel: 2, hidden: true });
		sheet.rows.set(4, { collapsed: false });
		sheet.rows.set(6, { outlineLevel: 1 });
		const outline = computeOutline(sheet, 'rows');
		expect(outline.levels).toBe(2);
		expect(outline.groups).toEqual([
			{ level: 1, start: 1, end: 3, summary: 4, collapsed: false },
			{ level: 1, start: 6, end: 6, summary: 7, collapsed: false },
			{ level: 2, start: 2, end: 3, summary: 4, collapsed: true },
		]);
	});

	it('puts summaries above when summaryBelow is off, and reads columns', () => {
		const sheet = createWorksheet('S', 1);
		sheet.summaryBelow = false;
		sheet.rows.set(5, { outlineLevel: 1 });
		expect(computeOutline(sheet, 'rows').groups[0].summary).toBe(4);
		sheet.cols.push({ min: 2, max: 4, outlineLevel: 1, hidden: true });
		sheet.cols.push({ min: 5, max: 5, collapsed: true });
		expect(computeOutline(sheet, 'cols').groups).toEqual([
			{ level: 1, start: 2, end: 4, summary: 5, collapsed: true },
		]);
	});

	it('is empty without outline levels', () => {
		expect(computeOutline(createWorksheet('S', 1), 'rows')).toEqual({ levels: 0, groups: [] });
	});
});
