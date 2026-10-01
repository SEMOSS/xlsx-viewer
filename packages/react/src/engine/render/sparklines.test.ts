import { cellKey, createWorkbook, createDefaultTheme } from 'xlsx-viewer-core';
import type { SparklineGroup, Workbook } from 'xlsx-viewer-core';
import { describe, expect, it } from 'vitest';

import { SheetGeometry } from '../geometry/sheet-geometry';
import { layoutSparklines, sparklineValues } from './sparklines';

function setup(group: Partial<SparklineGroup>, values: Array<number | string | null>): Workbook {
	const wb = createWorkbook();
	const sheet = wb.sheets[0];
	values.forEach((v, i) => {
		if (v !== null) {
			sheet.cells.set(cellKey(0, i), { value: v });
		}
	});
	sheet.sparklineGroups.push({
		type: 'line',
		sparklines: [{ formula: 'Sheet1!A1:E1', location: { row: 1, col: 0 } }],
		...group,
	});
	return wb;
}

const layout = (wb: Workbook) =>
	layoutSparklines(
		wb,
		wb.sheets[0],
		new SheetGeometry(wb.sheets[0], { zoom: 1 }),
		{ start: 0, end: 10 },
		{ start: 0, end: 10 },
		{ theme: createDefaultTheme() },
	);

describe('sparklines', () => {
	it('reads source values in order, numbers only', () => {
		const wb = setup({}, [1, 'x', 3, null, 5]);
		expect(sparklineValues(wb, wb.sheets[0], 'Sheet1!A1:E1')).toEqual([1, null, 3, null, 5]);
		expect(sparklineValues(wb, wb.sheets[0], 'Nope!A1:B1')).toEqual([]);
	});

	it('draws a line with gaps and high / low markers', () => {
		const wb = setup({ high: true, low: true, highColor: { rgb: 'FF00FF00' } }, [1, 4, null, 2, 3]);
		const [shape] = layout(wb);
		expect(shape.paths).toHaveLength(2);
		expect(shape.markers).toHaveLength(2);
		expect(shape.markers.some((m) => m.color.toLowerCase() === '#00ff00')).toBe(true);
		// Every point stays inside the cell.
		const cell = shape.clip;
		for (const m of shape.markers) {
			expect(m.cx).toBeGreaterThanOrEqual(cell.x);
			expect(m.cy).toBeLessThanOrEqual(cell.y + cell.height);
		}
	});

	it('spans or zeroes empty cells when asked', () => {
		expect(
			layout(setup({ displayEmptyCellsAs: 'span' }, [1, 4, null, 2, 3]))[0].paths,
		).toHaveLength(1);
		expect(
			layout(setup({ displayEmptyCellsAs: 'zero' }, [1, 4, null, 2, 3]))[0].paths,
		).toHaveLength(1);
	});

	it('draws columns from zero, negative ones only highlighted when asked', () => {
		const plain = layout(setup({ type: 'column' }, [3, -2, 5, 1, 4]))[0];
		expect(plain.bars).toHaveLength(5);
		expect(new Set(plain.bars.map((b) => b.color)).size).toBe(1);
		const neg = layout(setup({ type: 'column', negative: true }, [3, -2, 5, 1, 4]))[0];
		expect(new Set(neg.bars.map((b) => b.color)).size).toBe(2);
		// The negative bar hangs below the positive ones' baseline.
		const negBar = neg.bars[1];
		expect(negBar.y).toBeGreaterThanOrEqual(neg.bars[0].y + neg.bars[0].height - 0.01);
	});

	it('draws win/loss bars of equal height and skips zeros', () => {
		const [shape] = layout(setup({ type: 'stacked' }, [1, -1, 0, 2, -3]));
		expect(shape.bars).toHaveLength(4);
		expect(new Set(shape.bars.map((b) => b.height)).size).toBe(1);
	});
});
