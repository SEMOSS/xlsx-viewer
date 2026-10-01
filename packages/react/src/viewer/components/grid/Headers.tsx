import { columnToLetters } from 'xlsx-viewer-core';
import { memo } from 'react';

import type { Span } from '../../../engine/layout/types';
import {
	isColFullySelected,
	isRowFullySelected,
	selectionTouchesCol,
	selectionTouchesRow,
	type SelectionState,
} from '../../../engine/selection/selection';
import { useGridRender } from './grid-context';

interface ColumnHeadersProps {
	cols: Span;
	height: number;
	/** Offset below the column outline gutter. */
	top?: number;
	selection: SelectionState;
}

/** Column letters for a span of columns, in sheet x coordinates. */
export const ColumnHeaders = memo(function ColumnHeaders({
	cols,
	height,
	top = 0,
	selection,
}: ColumnHeadersProps) {
	const { geometry: g, zoom } = useGridRender();
	const cells = [];
	for (let c = cols.start; c <= cols.end; c++) {
		const w = g.colWidth(c);
		if (w <= 0) {
			continue;
		}
		const full = isColFullySelected(selection, c);
		const touched = full || selectionTouchesCol(selection, c);
		cells.push(
			<div
				key={c}
				className='xv-col-header'
				data-selected={touched || undefined}
				data-full={full || undefined}
				data-col={c}
				style={{
					left: g.colLeft(c),
					top,
					width: w,
					height,
					fontSize: `${Math.max(9, 11 * zoom)}px`,
				}}
			>
				{columnToLetters(c)}
			</div>,
		);
	}
	return <>{cells}</>;
});

interface RowHeadersProps {
	rows: Span;
	width: number;
	/** Offset right of the row outline gutter. */
	left?: number;
	selection: SelectionState;
}

/** Row numbers for a span of rows, in sheet y coordinates. */
export const RowHeaders = memo(function RowHeaders({
	rows,
	width,
	left = 0,
	selection,
}: RowHeadersProps) {
	const { geometry: g, zoom } = useGridRender();
	const cells = [];
	for (let r = rows.start; r <= rows.end; r++) {
		const h = g.rowHeight(r);
		if (h <= 0) {
			continue;
		}
		const full = isRowFullySelected(selection, r);
		const touched = full || selectionTouchesRow(selection, r);
		cells.push(
			<div
				key={r}
				className='xv-row-header'
				data-selected={touched || undefined}
				data-full={full || undefined}
				data-row={r}
				style={{
					top: g.rowTop(r),
					left,
					height: h,
					width,
					fontSize: `${Math.max(9, 11 * zoom)}px`,
				}}
			>
				{r + 1}
			</div>,
		);
	}
	return <>{cells}</>;
});
