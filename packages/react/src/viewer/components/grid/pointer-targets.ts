import type { RangeAddress } from 'xlsx-viewer-core';

import type { GridHit } from '../../../engine/geometry/hit-test';

/**
 * The range a fill-handle drag covers with the pointer over `hit`: `source`
 * extended in the direction the pointer moved furthest beyond it (`source`
 * itself while the pointer is inside it).
 */
export function fillDragTarget(
	source: RangeAddress,
	hit: { row: number; col: number },
): RangeAddress {
	const down = hit.row - source.bottom;
	const up = source.top - hit.row;
	const right = hit.col - source.right;
	const left = source.left - hit.col;
	const best = Math.max(down, up, right, left);
	if (best <= 0) {
		return source;
	}
	if (best === down) {
		return { ...source, bottom: hit.row };
	}
	if (best === up) {
		return { ...source, top: hit.row };
	}
	if (best === right) {
		return { ...source, right: hit.col };
	}
	return { ...source, left: hit.col };
}

/**
 * The mouse cursor over `hit` while no button is pressed. `overLink` is only
 * asked for cells (hyperlink lookups are not free).
 */
export function hoverCursor(
	hit: GridHit,
	overFillHandle: boolean,
	overLink: () => boolean,
): string {
	if (hit.colResize !== undefined) {
		return 'col-resize';
	}
	if (hit.rowResize !== undefined) {
		return 'row-resize';
	}
	if (hit.kind === 'colHeader') {
		return 'var(--xv-cursor-down, s-resize)';
	}
	if (hit.kind === 'rowHeader') {
		return 'var(--xv-cursor-right, e-resize)';
	}
	if (overFillHandle) {
		return 'crosshair';
	}
	return hit.kind === 'cell' && overLink() ? 'pointer' : 'cell';
}
