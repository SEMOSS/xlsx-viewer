import type { RangeAddress } from 'xlsx-viewer-core';
import { memo } from 'react';

import type { SheetGeometry } from '../../../engine/geometry/sheet-geometry';
import type { Span } from '../../../engine/layout/types';
import { currentRange, type SelectionState } from '../../../engine/selection/selection';
import type { MergeIndex } from '../../../engine/layout/style-at';

export interface RefHighlight {
	range: RangeAddress;
	color: string;
}

interface SelectionOverlayProps {
	geometry: SheetGeometry;
	merges: MergeIndex;
	rows: Span;
	cols: Span;
	selection: SelectionState;
	showFillHandle: boolean;
	copyRanges?: RangeAddress[];
	fillPreview?: RangeAddress | null;
	refHighlights?: RefHighlight[];
	/** Hide the selection (e.g. while a floating object is selected). */
	hidden?: boolean;
}

function clampToSpan(r: RangeAddress, rows: Span, cols: Span): RangeAddress | null {
	const top = Math.max(r.top, rows.start - 1);
	const bottom = Math.min(r.bottom, rows.end + 1);
	const left = Math.max(r.left, cols.start - 1);
	const right = Math.min(r.right, cols.end + 1);
	if (top > bottom || left > right) {
		return null;
	}
	return { top: Math.max(0, top), bottom, left: Math.max(0, left), right };
}

/**
 * Selection shading, the current-range border, the active cell, the fill
 * handle, the copy marquee and formula reference highlights for one pane.
 */
export const SelectionOverlay = memo(function SelectionOverlay(props: SelectionOverlayProps) {
	const { geometry: g, rows, cols, selection } = props;
	if (props.hidden) {
		return null;
	}
	const rect = (r: RangeAddress) => g.rangeRect(r);
	const active = props.merges.find(selection.active.row, selection.active.col) ?? {
		top: selection.active.row,
		left: selection.active.col,
		bottom: selection.active.row,
		right: selection.active.col,
	};
	const activeRect = rect(active);
	const current = currentRange(selection);
	const multi =
		selection.ranges.length > 1 || current.top !== current.bottom || current.left !== current.right;

	let shade = '';
	if (multi) {
		for (const r of selection.ranges) {
			const clipped = clampToSpan(r, rows, cols);
			if (!clipped) {
				continue;
			}
			const b = rect(clipped);
			shade += `M${b.x} ${b.y}h${b.width}v${b.height}h${-b.width}Z`;
		}
		shade += `M${activeRect.x} ${activeRect.y}h${activeRect.width}v${activeRect.height}h${-activeRect.width}Z`;
	}

	const curClipped = clampToSpan(current, rows, cols);
	const cur = curClipped ? rect(curClipped) : null;
	const fullCur = rect(current);
	const handleSize = 6;

	return (
		<div className='xv-selection' aria-hidden>
			{shade && (
				<svg
					className='xv-selection-shade'
					style={{ position: 'absolute', left: 0, top: 0, overflow: 'visible' }}
					width={1}
					height={1}
				>
					<path d={shade} fillRule='evenodd' fill='var(--xlsx-selection-fill)' />
				</svg>
			)}
			{cur && (
				<div
					className='xv-selection-border'
					style={{ left: cur.x - 1, top: cur.y - 1, width: cur.width + 1, height: cur.height + 1 }}
				/>
			)}
			{props.showFillHandle && selection.ranges.length === 1 && (
				<div
					className='xv-fill-handle'
					data-fill-handle
					style={{
						left: fullCur.x + fullCur.width - handleSize / 2 - 1,
						top: fullCur.y + fullCur.height - handleSize / 2 - 1,
						width: handleSize,
						height: handleSize,
					}}
				/>
			)}
			{props.fillPreview && (
				<div className='xv-fill-preview' style={boxStyle(rect(props.fillPreview))} />
			)}
			{props.copyRanges?.map((r, i) => (
				<svg
					// biome-ignore lint: positional
					key={i}
					className='xv-copy-marquee'
					style={{
						position: 'absolute',
						left: rect(r).x - 1,
						top: rect(r).y - 1,
						overflow: 'visible',
					}}
					width={rect(r).width + 1}
					height={rect(r).height + 1}
				>
					<rect
						x={0.5}
						y={0.5}
						width={rect(r).width}
						height={rect(r).height}
						fill='none'
						stroke='var(--xlsx-selection-border)'
						strokeWidth={2}
						strokeDasharray='5 3'
					/>
				</svg>
			))}
			{props.refHighlights?.map((h, i) => (
				<div
					// biome-ignore lint: positional
					key={i}
					className='xv-ref-highlight'
					style={{
						...boxStyle(rect(h.range)),
						borderColor: h.color,
						backgroundColor: `${h.color}14`,
					}}
				/>
			))}
		</div>
	);
});

function boxStyle(b: { x: number; y: number; width: number; height: number }) {
	return { left: b.x - 1, top: b.y - 1, width: b.width + 1, height: b.height + 1 };
}
