import type { AxisOutline } from 'xlsx-viewer-core';
import type { CSSProperties, ReactNode } from 'react';

import type { SheetGeometry } from '../../../engine/geometry/sheet-geometry';
import type { PaneLayout } from '../../../engine/layout/panes';
import type { MergeIndex } from '../../../engine/layout/style-at';
import type { Span } from '../../../engine/layout/types';
import type { SelectionState } from '../../../engine/selection/selection';
import type { EditState } from '../../state/types';
import { CellEditorLayer } from './CellEditorLayer';
import { ColumnHeaders, RowHeaders } from './Headers';
import { ColOutlineGutter, OutlineLevelButtons, RowOutlineGutter } from './Outline';
import { Pane } from './Pane';
import type { ViewportState } from './useGridViewport';

/** The four regions of the grid: the frozen corner, frozen rows, frozen columns and the body. */
export type PaneKind = 'body' | 'frozenRows' | 'frozenCols' | 'corner';

const EMPTY: Span = { start: 0, end: -1 };

/** Whether a span has any lines. Panes without frozen lines get no overlay (no hidden copies). */
const hasLines = (s: Span) => s.end >= s.start;

interface GridPanesProps {
	geometry: SheetGeometry;
	panes: PaneLayout;
	viewport: ViewportState;
	/** Horizontal scroll offset of the body (how far right the cell editor may grow). */
	scrollLeft: number;
	/** Header sizes; the outline gutters are part of the header width and height. */
	header: { width: number; height: number; rowGutter: number; colGutter: number };
	outline: { rows: AxisOutline; cols: AxisOutline };
	selection: SelectionState;
	merges: MergeIndex;
	editing: EditState | null;
	onEditorDone: () => void;
	/** The selection, decorations and drawings shown over one pane. */
	overlay: (pane: PaneKind, rows: Span, cols: Span) => ReactNode;
}

/**
 * The scrolling content of the grid, laid out as a 3 x 3 CSS grid: the corner,
 * the column and row headers (with the outline gutters), the frozen and body
 * panes, and the in-cell editor.
 */
export function GridPanes({
	geometry: g,
	panes,
	viewport: vp,
	scrollLeft,
	header,
	outline,
	selection,
	merges,
	editing,
	onEditorDone,
	overlay,
}: GridPanesProps) {
	const { width: headerWidth, height: headerHeight, rowGutter, colGutter } = header;
	const frozenRowSpan = panes.frozenRows
		? { start: panes.originRow, end: panes.bodyRow - 1 }
		: EMPTY;
	const frozenColSpan = panes.frozenCols
		? { start: panes.originCol, end: panes.bodyCol - 1 }
		: EMPTY;
	const bodyW = g.colLeft(vp.extentCols) - g.colLeft(panes.bodyCol);
	const bodyH = g.rowTop(vp.extentRows) - g.rowTop(panes.bodyRow);
	const originXFrozen = g.colLeft(panes.originCol);
	const originYFrozen = g.rowTop(panes.originRow);
	const originXBody = g.colLeft(panes.bodyCol);
	const originYBody = g.rowTop(panes.bodyRow);
	const viewportRight =
		g.colLeft(panes.bodyCol) + scrollLeft + vp.width - headerWidth - panes.frozenWidth;
	const content: CSSProperties = {
		width: headerWidth + panes.frozenWidth + bodyW,
		height: headerHeight + panes.frozenHeight + bodyH,
		gridTemplateColumns: `${headerWidth}px ${panes.frozenWidth}px ${bodyW}px`,
		gridTemplateRows: `${headerHeight}px ${panes.frozenHeight}px ${bodyH}px`,
	};

	return (
		<div className='xv-grid-content' style={content}>
			<div className='xv-corner' style={{ gridArea: '1 / 1', top: 0, left: 0 }}>
				{(outline.rows.levels > 0 || outline.cols.levels > 0) && (
					<OutlineLevelButtons
						rows={outline.rows}
						cols={outline.cols}
						rowGutter={rowGutter}
						colGutter={colGutter}
						headerHeight={headerHeight}
					/>
				)}
			</div>
			<div
				className='xv-col-headers'
				style={{ gridArea: '1 / 2', top: 0, left: headerWidth, zIndex: 6 }}
			>
				<div className='xv-pane-inner' style={{ transform: `translateX(${-originXFrozen}px)` }}>
					<ColumnHeaders
						cols={frozenColSpan}
						height={headerHeight - colGutter}
						top={colGutter}
						selection={selection}
					/>
					{colGutter > 0 && (
						<ColOutlineGutter outline={outline.cols} span={frozenColSpan} size={colGutter} />
					)}
				</div>
			</div>
			<div className='xv-col-headers' style={{ gridArea: '1 / 3', top: 0, zIndex: 5 }}>
				<div className='xv-pane-inner' style={{ transform: `translateX(${-originXBody}px)` }}>
					<ColumnHeaders
						cols={vp.cols}
						height={headerHeight - colGutter}
						top={colGutter}
						selection={selection}
					/>
					{colGutter > 0 && (
						<ColOutlineGutter outline={outline.cols} span={vp.cols} size={colGutter} />
					)}
				</div>
			</div>
			<div
				className='xv-row-headers'
				style={{ gridArea: '2 / 1', left: 0, top: headerHeight, zIndex: 6 }}
			>
				<div className='xv-pane-inner' style={{ transform: `translateY(${-originYFrozen}px)` }}>
					<RowHeaders
						rows={frozenRowSpan}
						width={headerWidth - rowGutter}
						left={rowGutter}
						selection={selection}
					/>
					{rowGutter > 0 && (
						<RowOutlineGutter outline={outline.rows} span={frozenRowSpan} size={rowGutter} />
					)}
				</div>
			</div>
			<Pane
				rows={frozenRowSpan}
				cols={frozenColSpan}
				originX={originXFrozen}
				originY={originYFrozen}
				className='xv-pane-corner'
				freezeBottom={panes.frozenRows > 0}
				freezeRight={panes.frozenCols > 0}
				style={{
					gridArea: '2 / 2',
					position: 'sticky',
					top: headerHeight,
					left: headerWidth,
					zIndex: 4,
				}}
			>
				{hasLines(frozenRowSpan) &&
					hasLines(frozenColSpan) &&
					overlay('corner', frozenRowSpan, frozenColSpan)}
			</Pane>
			<Pane
				rows={frozenRowSpan}
				cols={vp.cols}
				originX={originXBody}
				originY={originYFrozen}
				className='xv-pane-frozen-rows'
				freezeBottom={panes.frozenRows > 0}
				style={{ gridArea: '2 / 3', position: 'sticky', top: headerHeight, zIndex: 3 }}
			>
				{hasLines(frozenRowSpan) && overlay('frozenRows', frozenRowSpan, vp.cols)}
			</Pane>
			<div className='xv-row-headers' style={{ gridArea: '3 / 1', left: 0, zIndex: 3 }}>
				<div className='xv-pane-inner' style={{ transform: `translateY(${-originYBody}px)` }}>
					<RowHeaders
						rows={vp.rows}
						width={headerWidth - rowGutter}
						left={rowGutter}
						selection={selection}
					/>
					{rowGutter > 0 && (
						<RowOutlineGutter outline={outline.rows} span={vp.rows} size={rowGutter} />
					)}
				</div>
			</div>
			<Pane
				rows={vp.rows}
				cols={frozenColSpan}
				originX={originXFrozen}
				originY={originYBody}
				className='xv-pane-frozen-cols'
				freezeRight={panes.frozenCols > 0}
				style={{ gridArea: '3 / 2', position: 'sticky', left: headerWidth, zIndex: 2 }}
			>
				{hasLines(frozenColSpan) && overlay('frozenCols', vp.rows, frozenColSpan)}
			</Pane>
			<Pane
				rows={vp.rows}
				cols={vp.cols}
				originX={originXBody}
				originY={originYBody}
				className='xv-pane-body'
				style={{ gridArea: '3 / 3' }}
			>
				{overlay('body', vp.rows, vp.cols)}
			</Pane>
			{editing && (
				<CellEditorLayer
					edit={editing}
					geometry={g}
					merges={merges}
					panes={panes}
					header={{ headerWidth, headerHeight }}
					origins={{
						frozenX: originXFrozen,
						frozenY: originYFrozen,
						bodyX: originXBody,
						bodyY: originYBody,
					}}
					maxRight={
						editing.col < panes.bodyCol ? originXFrozen + vp.width - headerWidth : viewportRight
					}
					onDone={onEditorDone}
				/>
			)}
		</div>
	);
}
