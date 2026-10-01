import { describe, expect, it, vi } from 'vitest';

import type { GridHit } from '../../../engine/geometry/hit-test';
import { fillDragTarget, hoverCursor } from './pointer-targets';

const SOURCE = { top: 2, left: 2, bottom: 3, right: 3 };

describe('fillDragTarget', () => {
	it('keeps the source while the pointer is inside it', () => {
		expect(fillDragTarget(SOURCE, { row: 3, col: 2 })).toEqual(SOURCE);
	});

	it('extends toward the side the pointer moved furthest', () => {
		expect(fillDragTarget(SOURCE, { row: 9, col: 5 })).toEqual({ ...SOURCE, bottom: 9 });
		expect(fillDragTarget(SOURCE, { row: 0, col: 3 })).toEqual({ ...SOURCE, top: 0 });
		expect(fillDragTarget(SOURCE, { row: 4, col: 8 })).toEqual({ ...SOURCE, right: 8 });
		expect(fillDragTarget(SOURCE, { row: 2, col: 0 })).toEqual({ ...SOURCE, left: 0 });
	});

	it('prefers down, then up, then right on a tie', () => {
		expect(fillDragTarget(SOURCE, { row: 5, col: 5 })).toEqual({ ...SOURCE, bottom: 5 });
		expect(fillDragTarget(SOURCE, { row: 0, col: 5 })).toEqual({ ...SOURCE, top: 0 });
	});
});

describe('hoverCursor', () => {
	const hit = (over: Partial<GridHit>): GridHit => ({
		kind: 'cell',
		row: 0,
		col: 0,
		x: 0,
		y: 0,
		inBodyX: true,
		inBodyY: true,
		...over,
	});

	it('shows resize cursors on header edges before anything else', () => {
		expect(hoverCursor(hit({ kind: 'colHeader', colResize: 1 }), true, () => true)).toBe(
			'col-resize',
		);
		expect(hoverCursor(hit({ kind: 'rowHeader', rowResize: 1 }), false, () => false)).toBe(
			'row-resize',
		);
	});

	it('points down columns and across rows on the headers', () => {
		expect(hoverCursor(hit({ kind: 'colHeader' }), false, () => false)).toBe(
			'var(--xv-cursor-down, s-resize)',
		);
		expect(hoverCursor(hit({ kind: 'rowHeader' }), false, () => false)).toBe(
			'var(--xv-cursor-right, e-resize)',
		);
	});

	it('uses a crosshair on the fill handle, a pointer on links and the cell cursor elsewhere', () => {
		expect(hoverCursor(hit({}), true, () => true)).toBe('crosshair');
		expect(hoverCursor(hit({}), false, () => true)).toBe('pointer');
		expect(hoverCursor(hit({}), false, () => false)).toBe('cell');
		expect(hoverCursor(hit({ kind: 'corner' }), false, () => true)).toBe('cell');
	});

	it('only looks up hyperlinks for cells', () => {
		const overLink = vi.fn(() => true);
		hoverCursor(hit({ kind: 'corner' }), false, overLink);
		hoverCursor(hit({ kind: 'colHeader' }), false, overLink);
		expect(overLink).not.toHaveBeenCalled();
	});
});
