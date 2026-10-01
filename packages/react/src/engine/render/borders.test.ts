import { describe, expect, it } from 'vitest';
import type { BorderEdge, BorderStyle, ColorContext } from 'xlsx-viewer-core';
import { createDefaultTheme } from 'xlsx-viewer-core';

import { BorderCollector, borderStroke, segmentLines, type BorderSegment } from './borders';

const ctx: ColorContext = { theme: createDefaultTheme() };

const edge = (style: BorderStyle, rgb?: string): BorderEdge => ({
	style,
	color: rgb ? { rgb } : undefined,
});

/** Excel's conflict order, weakest first. */
const ORDER: Array<Exclude<BorderStyle, 'none'>> = [
	'hair',
	'dotted',
	'dashDotDot',
	'dashDot',
	'dashed',
	'thin',
	'mediumDashDotDot',
	'slantDashDot',
	'mediumDashDot',
	'mediumDashed',
	'medium',
	'thick',
	'double',
];

describe('borderStroke', () => {
	it('has no stroke for none and ranks the other styles from hair to double', () => {
		expect(borderStroke('none')).toBeUndefined();
		const weights = ORDER.map((s) => borderStroke(s)!.weight);
		expect(weights).toEqual([...weights].sort((a, b) => a - b));
		expect(new Set(weights).size).toBe(ORDER.length);
	});

	it('maps widths, dashes and double lines', () => {
		expect(borderStroke('thin')).toEqual({ width: 1, weight: 6 });
		expect(borderStroke('medium')?.width).toBe(2);
		expect(borderStroke('thick')?.width).toBe(3);
		expect(borderStroke('double')?.double).toBe(true);
		expect(borderStroke('dashed')?.dash).toBe('3 1');
		expect(borderStroke('mediumDashed')).toMatchObject({ width: 2, dash: '9 3' });
		expect(borderStroke('hair')?.dash).toBe('1 1');
	});
});

describe('BorderCollector', () => {
	it('records a segment with its resolved color', () => {
		const c = new BorderCollector(ctx, 1);
		c.add('h', 3, 2, 60, 128, 192, edge('thin', 'FFFF0000'));
		expect(c.segments()).toEqual([
			{
				orientation: 'h',
				at: 60,
				from: 128,
				to: 192,
				color: '#FF0000',
				stroke: borderStroke('thin'),
				zoom: 1,
			},
		]);
	});

	it('ignores missing and none edges', () => {
		const c = new BorderCollector(ctx, 1);
		c.add('h', 0, 0, 0, 0, 64, undefined);
		c.add('v', 0, 0, 0, 0, 20, edge('none'));
		expect(c.segments()).toEqual([]);
	});

	it('uses the default color for automatic edges, and resolves theme colors', () => {
		const c = new BorderCollector(ctx, 1, '#123456');
		c.add('h', 0, 0, 0, 0, 64, edge('thin'));
		c.add('h', 0, 1, 0, 64, 128, { style: 'thin', color: { auto: true } });
		c.add('h', 0, 2, 0, 128, 192, { style: 'thin', color: { theme: 5 } });
		expect(c.segments().map((s) => s.color)).toEqual(['#123456', '#123456', '#ED7D31']);
		const black = new BorderCollector(ctx, 1);
		black.add('v', 1, 0, 64, 0, 20, edge('thin'));
		expect(black.segments()[0].color).toBe('#000000');
	});

	it('lets the stronger style win a shared edge, in either order', () => {
		const upper = new BorderCollector(ctx, 1);
		upper.add('h', 2, 0, 40, 0, 64, edge('thick', 'FFFF0000'));
		upper.add('h', 2, 0, 40, 0, 64, edge('thin', 'FF0000FF'));
		expect(upper.segments()).toHaveLength(1);
		expect(upper.segments()[0]).toMatchObject({ color: '#FF0000', stroke: { width: 3 } });

		const lower = new BorderCollector(ctx, 1);
		lower.add('h', 2, 0, 40, 0, 64, edge('dotted', 'FF0000FF'));
		lower.add('h', 2, 0, 40, 0, 64, edge('double', 'FF00FF00'));
		expect(lower.segments()[0]).toMatchObject({ color: '#00FF00', stroke: { double: true } });
	});

	it('gives ties to the edge added last (the right or lower cell)', () => {
		const c = new BorderCollector(ctx, 1);
		c.add('v', 1, 0, 64, 0, 20, edge('thin', 'FFFF0000'));
		c.add('v', 1, 0, 64, 0, 20, edge('thin', 'FF0000FF'));
		expect(c.segments().map((s) => s.color)).toEqual(['#0000FF']);
	});

	it('resolves every pair of styles by weight', () => {
		for (const a of ORDER) {
			for (const b of ORDER) {
				const c = new BorderCollector(ctx, 1);
				c.add('h', 1, 0, 20, 0, 64, edge(a, 'FFAA0000'));
				c.add('h', 1, 0, 20, 0, 64, edge(b, 'FF00AA00'));
				const winner = ORDER.indexOf(a) > ORDER.indexOf(b) ? a : b;
				expect(c.segments()[0].stroke, `${a} vs ${b}`).toBe(borderStroke(winner));
			}
		}
	});

	it('keeps edges of different lines, cells and orientations apart', () => {
		const c = new BorderCollector(ctx, 1);
		c.add('h', 1, 0, 20, 0, 64, edge('thin'));
		c.add('h', 1, 1, 20, 64, 128, edge('thin'));
		c.add('h', 2, 0, 40, 0, 64, edge('thin'));
		c.add('v', 1, 1, 64, 20, 40, edge('thin'));
		expect(c.segments()).toHaveLength(4);
	});
});

function seg(
	style: Exclude<BorderStyle, 'none'>,
	zoom: number,
	orientation: 'h' | 'v' = 'h',
): BorderSegment {
	return {
		orientation,
		at: 40,
		from: 0,
		to: 64,
		color: '#000000',
		stroke: borderStroke(style)!,
		zoom,
	};
}

describe('segmentLines', () => {
	it('centres odd widths on the gridline pixel and even widths on the line', () => {
		expect(segmentLines(seg('thin', 1))).toEqual([
			{ x1: 0, y1: 39.5, x2: 64, y2: 39.5, width: 1, dash: undefined },
		]);
		expect(segmentLines(seg('medium', 1))).toEqual([
			{ x1: 0, y1: 40, x2: 64, y2: 40, width: 2, dash: undefined },
		]);
		expect(segmentLines(seg('thick', 1, 'v'))).toEqual([
			{ x1: 39.5, y1: 0, x2: 39.5, y2: 64, width: 3, dash: undefined },
		]);
	});

	it('scales widths and dash patterns with zoom', () => {
		expect(segmentLines(seg('medium', 2))[0]).toMatchObject({ width: 4, y1: 40 });
		expect(segmentLines(seg('thin', 1.5))[0]).toMatchObject({ width: 2, y1: 40 });
		expect(segmentLines(seg('dashed', 1))[0].dash).toBe('3 1');
		expect(segmentLines(seg('dashed', 2))[0].dash).toBe('6 2');
		expect(segmentLines(seg('dashDot', 2))[0].dash).toBe('18 6 6 6');
		expect(segmentLines(seg('mediumDashed', 1.5))[0]).toMatchObject({
			width: 3,
			dash: '13.5 4.5',
		});
	});

	it('never goes below 1px or half-size dashes when zoomed out', () => {
		expect(segmentLines(seg('thin', 0.25))[0]).toMatchObject({ width: 1, y1: 39.5 });
		expect(segmentLines(seg('thick', 0.1))[0].width).toBe(2);
		expect(segmentLines(seg('hair', 0.25))[0].dash).toBe('0.5 0.5');
	});

	it('draws double borders as two 1px lines around the gridline pixel', () => {
		expect(segmentLines(seg('double', 1))).toEqual([
			{ x1: 0, y1: 38.5, x2: 64, y2: 38.5, width: 1 },
			{ x1: 0, y1: 40.5, x2: 64, y2: 40.5, width: 1 },
		]);
		expect(segmentLines(seg('double', 1, 'v'))).toEqual([
			{ x1: 38.5, y1: 0, x2: 38.5, y2: 64, width: 1 },
			{ x1: 40.5, y1: 0, x2: 40.5, y2: 64, width: 1 },
		]);
	});

	it('widens the gap of double borders with zoom', () => {
		const [a, b] = segmentLines(seg('double', 2));
		// 3px * 2 = 6px wide (even): centred on the line, 2px gap on each side.
		expect([a.y1, b.y1]).toEqual([38, 42]);
		const [c, d] = segmentLines(seg('double', 0.3));
		expect(d.y1 - c.y1).toBe(2);
	});
});
