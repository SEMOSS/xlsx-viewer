import { describe, expect, it } from 'vitest';

import { isLinePreset, presetPath } from './shape-paths';

type Point = [number, number];

/** Vertices of a single-polygon path `M x yL x y ... Z`. */
function polygon(path: string): Point[] {
	expect(path.startsWith('M')).toBe(true);
	expect(path.endsWith('Z')).toBe(true);
	return path
		.slice(1, -1)
		.split('L')
		.map((p) => p.split(' ').map(Number) as Point);
}

function expectPoints(actual: Point[], expected: Point[]): void {
	expect(actual).toHaveLength(expected.length);
	actual.forEach(([x, y], i) => {
		expect(x, `x of point ${i}`).toBeCloseTo(expected[i][0], 6);
		expect(y, `y of point ${i}`).toBeCloseTo(expected[i][1], 6);
	});
}

describe('presetPath: boxes and ellipses', () => {
	it('draws rectangles (and unknown presets) around the whole box', () => {
		expect(presetPath('rect', 100, 50)).toBe('M0 0H100V50H0Z');
		expect(presetPath('flowChartProcess', 100, 50)).toBe('M0 0H100V50H0Z');
		expect(presetPath('cloud', 30, 20)).toBe('M0 0H30V20H0Z');
	});

	it('rounds rectangle corners by 1/6 of the shorter side, or the adjust value', () => {
		const r = 50 * 0.16667;
		expect(presetPath('roundRect', 100, 50).startsWith(`M${r} 0H${100 - r}A${r} ${r} `)).toBe(true);
		expect(presetPath('roundRect', 100, 50, { adj: 20000 }).startsWith('M10 0H90A10 10 ')).toBe(
			true,
		);
		// The radius never exceeds half the box.
		expect(presetPath('roundRect', 100, 40, { adj: 90000 }).startsWith('M20 0H80A20 20 ')).toBe(
			true,
		);
	});

	it('draws ellipses as two arcs through the middle of the left and right sides', () => {
		expect(presetPath('ellipse', 80, 40)).toBe('M0 20A40 20 0 1 0 80 20A40 20 0 1 0 0 20Z');
		expect(presetPath('flowChartConnector', 80, 40)).toBe(presetPath('ellipse', 80, 40));
	});

	it('draws cans with a separate front edge for the lid', () => {
		const path = presetPath('can', 40, 100);
		// ry = 40 * 0.25 / 2.
		expect(path).toBe('M0 5A20 5 0 0 1 40 5V95A20 5 0 0 1 0 95Z M0 5A20 5 0 0 0 40 5');
	});
});

describe('presetPath: polygons', () => {
	it('places the triangle apex from the adjust value', () => {
		expectPoints(polygon(presetPath('triangle', 100, 50)), [
			[50, 0],
			[100, 50],
			[0, 50],
		]);
		expectPoints(polygon(presetPath('triangle', 100, 50, { adj: 0 })), [
			[0, 0],
			[100, 50],
			[0, 50],
		]);
		expectPoints(polygon(presetPath('rtTriangle', 10, 20)), [
			[0, 0],
			[10, 20],
			[0, 20],
		]);
	});

	it('draws diamonds through the side midpoints', () => {
		expectPoints(polygon(presetPath('diamond', 60, 40)), [
			[30, 0],
			[60, 20],
			[30, 40],
			[0, 20],
		]);
		expect(presetPath('flowChartDecision', 60, 40)).toBe(presetPath('diamond', 60, 40));
	});

	it('offsets parallelograms, trapezoids and hexagons by a quarter of the shorter side', () => {
		expectPoints(polygon(presetPath('parallelogram', 100, 40)), [
			[10, 0],
			[100, 0],
			[90, 40],
			[0, 40],
		]);
		expectPoints(polygon(presetPath('trapezoid', 100, 40)), [
			[10, 0],
			[90, 0],
			[100, 40],
			[0, 40],
		]);
		expectPoints(polygon(presetPath('hexagon', 100, 40)), [
			[10, 0],
			[90, 0],
			[100, 20],
			[90, 40],
			[10, 40],
			[0, 20],
		]);
	});

	it('draws pentagons (home plates) and chevrons pointing right', () => {
		expectPoints(polygon(presetPath('homePlate', 100, 40)), [
			[0, 0],
			[80, 0],
			[100, 20],
			[80, 40],
			[0, 40],
		]);
		expect(presetPath('pentagon', 100, 40)).toBe(presetPath('homePlate', 100, 40));
		expectPoints(polygon(presetPath('chevron', 100, 40)), [
			[0, 0],
			[80, 0],
			[100, 20],
			[80, 40],
			[0, 40],
			[20, 20],
		]);
	});

	it('cuts octagon corners and builds a plus from the arm offset', () => {
		const octagon = polygon(presetPath('octagon', 100, 100));
		expect(octagon).toHaveLength(8);
		expect(octagon[0][0]).toBeCloseTo(29.289, 3);
		const plus = polygon(presetPath('plus', 40, 40));
		expect(plus).toHaveLength(12);
		expectPoints(plus.slice(0, 4), [
			[10, 0],
			[30, 0],
			[30, 10],
			[40, 10],
		]);
	});

	it('draws arrows with a half-height shaft and a head as long as half the shorter side', () => {
		const right = polygon(presetPath('rightArrow', 100, 50));
		expectPoints(right, [
			[0, 12.5],
			[75, 12.5],
			[75, 0],
			[100, 25],
			[75, 50],
			[75, 37.5],
			[0, 37.5],
		]);
		// Left arrows mirror right arrows horizontally.
		expectPoints(
			polygon(presetPath('leftArrow', 100, 50)),
			right.map(([x, y]) => [100 - x, y]),
		);
		const up = polygon(presetPath('upArrow', 50, 100));
		expectPoints(up, [
			[12.5, 100],
			[12.5, 25],
			[0, 25],
			[25, 0],
			[50, 25],
			[37.5, 25],
			[37.5, 100],
		]);
		expectPoints(
			polygon(presetPath('downArrow', 50, 100)),
			up.map(([x, y]) => [x, 100 - y]),
		);
	});

	it('honours arrow adjust values', () => {
		const thin = polygon(presetPath('rightArrow', 100, 50, { adj1: 20000, adj2: 100000 }));
		expect(thin[0]).toEqual([0, 20]);
		expect(thin[1]).toEqual([50, 20]);
	});

	it('draws a five-pointed star inside the box', () => {
		const star = polygon(presetPath('star5', 100, 100));
		expect(star).toHaveLength(10);
		expect(star[0][0]).toBeCloseTo(50, 6);
		expect(star[0][1]).toBeLessThan(5);
		for (const [x, y] of star) {
			expect(x).toBeGreaterThanOrEqual(0);
			expect(x).toBeLessThanOrEqual(100);
			expect(y).toBeGreaterThanOrEqual(0);
			expect(y).toBeLessThanOrEqual(100);
		}
	});

	it('draws callouts with a tail below the box', () => {
		expect(presetPath('wedgeRectCallout', 100, 40)).toMatch(
			/^M0 0H100V40H42L29\.16\d* 45L25 40H0Z$/,
		);
		expect(presetPath('wedgeRoundRectCallout', 100, 40, { adj1: 0, adj2: 50000 })).toBe(
			'M0 0H100V40H42L50 40L25 40H0Z',
		);
	});
});

describe('lines and connectors', () => {
	it('draws a diagonal across the box', () => {
		expect(presetPath('line', 30, 40)).toBe('M0 0L30 40');
		expect(presetPath('straightConnector1', 30, 40)).toBe('M0 0L30 40');
		// Elbow connectors bend at their adjust handle (half way by default).
		expect(presetPath('bentConnector3', 5, 5)).toBe('M0 0L2.5 0L2.5 5L5 5');
		expect(presetPath('bentConnector3', 10, 4, { adj1: 25000 })).toBe('M0 0L2.5 0L2.5 4L10 4');
	});

	it('recognises line and connector presets', () => {
		for (const p of [
			'line',
			'straightConnector',
			'straightConnector1',
			'bentConnector3',
			'curvedConnector2',
		]) {
			expect(isLinePreset(p), p).toBe(true);
		}
		for (const p of ['rect', 'lineInv', 'bentConnector', 'rightArrow', 'curvedRightArrow']) {
			expect(isLinePreset(p), p).toBe(false);
		}
	});

	it('draws every connector that isLinePreset accepts as an open line', () => {
		const connectors = [
			'bentConnector2',
			'bentConnector3',
			'bentConnector4',
			'bentConnector5',
			'curvedConnector2',
			'curvedConnector3',
			'curvedConnector4',
			'curvedConnector5',
		];
		for (const p of connectors) {
			expect(isLinePreset(p), p).toBe(true);
			expect(presetPath(p, 40, 20), p).not.toMatch(/Z$/);
		}
	});
});
