import { describe, expect, it } from 'vitest';
import { linePath, markerNode, polar, sectorPath, smoothPath } from './geometry';

describe('geometry', () => {
	it('builds polylines and smooth curves through the points', () => {
		expect(
			linePath([
				[0, 0],
				[10, 5],
			]),
		).toBe('M0 0 L10 5');
		const d = smoothPath([
			[0, 0],
			[10, 10],
			[20, 0],
		]);
		expect(d.startsWith('M0 0 C')).toBe(true);
		expect(d.endsWith(' 20 0')).toBe(true);
		expect(d.match(/C/g)).toHaveLength(2);
		expect(
			smoothPath([
				[0, 0],
				[1, 1],
			]),
		).toBe('M0 0 L1 1');
	});

	it("measures angles clockwise from 12 o'clock", () => {
		const [x, y] = polar(0, 0, 10, 90);
		expect(x).toBeCloseTo(10);
		expect(y).toBeCloseTo(0);
		expect(polar(0, 0, 10, 0)[1]).toBeCloseTo(-10);
	});

	it('draws pie sectors, doughnut segments and full rings', () => {
		expect(sectorPath(50, 50, 40, 0, 0, 90)).toBe('M50 50 L50 10 A40 40 0 0 1 90 50 Z');
		expect(sectorPath(50, 50, 40, 0, 0, 270)).toContain('A40 40 0 1 1');
		const ring = sectorPath(50, 50, 40, 20, 0, 360);
		expect(ring.match(/A/g)).toHaveLength(4);
		expect(sectorPath(50, 50, 40, 20, 0, 90)).toContain('A20 20 0 0 0');
	});

	it('draws every marker symbol', () => {
		const paint = { fill: '#f00', stroke: '#000', strokeWidth: 1 };
		expect(markerNode('none', 0, 0, 8, paint)).toBeUndefined();
		expect(markerNode('circle', 5, 5, 8, paint)).toMatchObject({
			type: 'circle',
			r: 4,
			role: 'marker',
		});
		expect(markerNode('square', 5, 5, 8, paint)).toMatchObject({
			type: 'rect',
			x: 1,
			y: 1,
			width: 8,
		});
		for (const symbol of [
			'diamond',
			'triangle',
			'x',
			'plus',
			'star',
			'dash',
			'dot',
			'picture',
			'auto',
		] as const) {
			expect(markerNode(symbol, 5, 5, 8, paint)).toBeDefined();
		}
		expect(markerNode('x', 5, 5, 8, { fill: '#f00' })).toMatchObject({ stroke: '#f00' });
	});
});
