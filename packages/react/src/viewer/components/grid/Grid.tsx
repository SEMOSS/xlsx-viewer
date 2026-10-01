import { useCallback, useEffect, useMemo, useRef, useState } from 'react';

import type { Note, ThreadedComment } from 'xlsx-viewer-core';
import { computeOutline, MAX_COLS, MAX_ROWS } from 'xlsx-viewer-core';

import { useT } from '../../../i18n';
import { SheetGeometry } from '../../../engine/geometry/sheet-geometry';
import {
	sheetToViewportX,
	sheetToViewportY,
	type GridHit,
} from '../../../engine/geometry/hit-test';
import { columnHeaderHeight, computePanes, rowHeaderWidth } from '../../../engine/layout/panes';
import { MergeIndex } from '../../../engine/layout/style-at';
import { createCellOverlays } from '../../../engine/layout/overlays';
import type { Span } from '../../../engine/layout/types';
import type { GridCommand } from '../../../engine/keymap';
import { currentRange } from '../../../engine/selection/selection';
import { hyperlinkForCell } from '../../commands/hyperlink-commands';
import { useStoreSelector, useViewerStore } from '../../state/use-store';
import { CellDecorations, type HoverLink } from './CellDecorations';
import { CommentThreadCard } from './CommentThreadCard';
import { DrawingsLayer } from '../drawings/DrawingsLayer';
import { GridRenderContextReact, type GridRenderContext } from './grid-context';
import { GridFocusSink } from './GridFocusSink';
import { GridPanes, type PaneKind } from './GridPanes';
import { outlineGutterSize } from './Outline';
import { SelectionOverlay, type RefHighlight } from './SelectionOverlay';
import { TraceArrowsLayer } from './TraceArrowsLayer';
import { useGridKeyboard } from './useGridKeyboard';
import { useGridPointer } from './useGridPointer';
import { useGridViewport } from './useGridViewport';

export interface GridProps {
	onAction?: (action: GridCommand['id']) => void;
	onContextMenu?: (e: { clientX: number; clientY: number; hit?: GridHit }) => void;
	refHighlights?: RefHighlight[];
	cf?: GridRenderContext['cf'];
	/** Extra content rendered in sheet coordinates inside the body pane (drawings, comment markers). */
	renderOverlay?: (args: { geometry: SheetGeometry; pane: PaneKind }) => React.ReactNode;
}

/** The spreadsheet grid of the active sheet. */
export function Grid({ onAction, onContextMenu, refHighlights, cf, renderOverlay }: GridProps) {
	const t = useT();
	const store = useViewerStore();
	const snapshot = useStoreSelector((s) => s);
	const { ui, docVersion } = snapshot;
	const sheetIndex = ui.activeSheet;
	const sheet = store.workbook.sheets[sheetIndex];
	const zoom = (sheet.view.zoomScale || 100) / 100;
	const selection = store.selectionOf(sheetIndex);
	const sinkRef = useRef<HTMLTextAreaElement>(null);

	const geometry = useMemo(
		() => new SheetGeometry(sheet, { zoom }),
		// eslint-disable-next-line react-hooks/exhaustive-deps
		[sheet, zoom, docVersion],
	);
	const panes = useMemo(() => computePanes(sheet, geometry), [sheet, geometry]);
	// eslint-disable-next-line react-hooks/exhaustive-deps
	const merges = useMemo(() => new MergeIndex(sheet.merges), [sheet, docVersion]);
	// eslint-disable-next-line react-hooks/exhaustive-deps
	const used = useMemo(() => ({ ...store.usedExtentOf(sheet) }), [store, sheet, docVersion]);
	const overlays = useMemo(
		() => createCellOverlays(store.workbook, sheetIndex, store.conditionalFormats),
		// eslint-disable-next-line react-hooks/exhaustive-deps
		[store, sheetIndex, docVersion],
	);
	// eslint-disable-next-line react-hooks/exhaustive-deps
	const rowOutline = useMemo(() => computeOutline(sheet, 'rows'), [sheet, docVersion]);
	// eslint-disable-next-line react-hooks/exhaustive-deps
	const colOutline = useMemo(() => computeOutline(sheet, 'cols'), [sheet, docVersion]);
	const showHeaders = sheet.view.showRowColHeaders;
	const rowGutter = showHeaders ? outlineGutterSize(rowOutline.levels, zoom) : 0;
	const colGutter = showHeaders ? outlineGutterSize(colOutline.levels, zoom) : 0;
	const headerHeight = showHeaders ? colGutter + columnHeaderHeight(zoom) : 0;

	const scrollStore = useRef<Record<string, { left: number; top: number }>>({});
	// The row header width depends on the largest visible row number (known after the
	// first render); use the previous render's value.
	const lastRowEnd = useRef(100);
	const headerWidth = showHeaders ? rowGutter + rowHeaderWidth(lastRowEnd.current + 1, zoom) : 0;
	const viewport = useGridViewport({
		geometry,
		panes,
		headerWidth,
		headerHeight,
		usedRows: Math.max(1, used.rows),
		usedCols: Math.max(1, used.cols),
		sheetId: sheet.id,
		initialScroll: scrollStore.current[sheet.id],
		onScrollPosition: (pos) => {
			scrollStore.current[sheet.id] = pos;
		},
	});
	const { state: vp } = viewport;
	lastRowEnd.current = vp.rows.end;
	const metrics = viewport.metrics;

	const focusGrid = useCallback(() => {
		sinkRef.current?.focus({ preventScroll: true });
	}, []);

	/**
	 * Scroll the active cell into view, or with `end` the moving end of an extended
	 * selection (Shift+arrows). Whole rows never scroll sideways, whole columns never
	 * up or down: clicking a row header must not jump to column XFD.
	 */
	const ensureActiveVisible = useCallback(
		(target: 'active' | 'end' = 'active') => {
			const sel = store.selectionOf();
			const r = currentRange(sel);
			const end = target === 'end';
			const row = end ? (sel.anchor.row === r.top ? r.bottom : r.top) : sel.active.row;
			const col = end ? (sel.anchor.col === r.left ? r.right : r.left) : sel.active.col;
			const allRows = r.top === 0 && r.bottom >= MAX_ROWS - 1;
			const allCols = r.left === 0 && r.right >= MAX_COLS - 1;
			viewport.ensureVisible(
				allRows ? null : Math.min(row, MAX_ROWS - 1),
				allCols ? null : Math.min(col, MAX_COLS - 1),
			);
		},
		[store, viewport],
	);

	// Keep the active cell visible after programmatic selection changes (Go To, Name Box...).
	const lastActive = useRef('');
	useEffect(() => {
		const key = `${sheet.id}:${selection.active.row}:${selection.active.col}`;
		if (lastActive.current && lastActive.current !== key) {
			ensureActiveVisible();
		}
		lastActive.current = key;
	}, [sheet.id, selection.active.row, selection.active.col, ensureActiveVisible]);

	const [hoverNote, setHoverNote] = useState<Note | null>(null);
	const [hoverLink, setHoverLink] = useState<HoverLink | null>(null);
	const [hoverThread, setHoverThread] = useState<ThreadedComment | null>(null);
	const pointer = useGridPointer({
		store,
		scrollerRef: viewport.scrollerRef,
		metrics,
		scrollBy: viewport.scrollBy,
		onContextMenu,
		focusGrid,
		onHover: (hit) => {
			const note =
				hit && hit.kind === 'cell'
					? (sheet.notes.find((n) => n.ref.row === hit.row && n.ref.col === hit.col) ?? null)
					: null;
			setHoverNote((cur) => (cur === note ? cur : note));
			const thread =
				hit && hit.kind === 'cell'
					? (sheet.threadedComments.find(
							(c) => !c.parentId && c.ref.row === hit.row && c.ref.col === hit.col,
						) ?? null)
					: null;
			setHoverThread((cur) => (cur === thread ? cur : thread));
			const link =
				hit && hit.kind === 'cell' ? hyperlinkForCell(store, hit.row, hit.col) : undefined;
			setHoverLink((cur) =>
				link && hit
					? cur?.link === link && cur.row === hit.row && cur.col === hit.col
						? cur
						: { link, row: hit.row, col: hit.col }
					: null,
			);
		},
	});

	const onKeyDown = useGridKeyboard({
		store,
		ensureActiveVisible,
		pageRows: () => Math.max(1, vp.rows.end - vp.rows.start - 24),
		pageCols: () => Math.max(1, vp.cols.end - vp.cols.start - 8),
		onAction,
		openContextMenu: () => {
			const m = metrics();
			const rect = viewport.scrollerRef.current?.getBoundingClientRect();
			const sel = store.selectionOf();
			const x =
				sheetToViewportX(m, geometry.colLeft(sel.active.col), sel.active.col) + (rect?.left ?? 0);
			const y =
				sheetToViewportY(m, geometry.rowTop(sel.active.row + 1), sel.active.row) + (rect?.top ?? 0);
			onContextMenu?.({ clientX: x, clientY: y });
		},
	});

	// Focus the grid on mount and whenever editing ends.
	useEffect(() => {
		if (!ui.edit) {
			focusGrid();
		}
	}, [ui.edit, focusGrid]);

	const ctx: GridRenderContext = useMemo(
		() => ({
			sheet,
			sheetIndex,
			geometry,
			panes,
			merges,
			styles: store.styles,
			colorCtx: { theme: store.workbook.theme, indexedColors: store.workbook.styles.indexedColors },
			zoom,
			docVersion,
			date1904: store.workbook.date1904,
			showGridLines: sheet.view.showGridLines,
			showFormulas: sheet.view.showFormulas,
			showZeros: sheet.view.showZeros,
			gridlineColor: 'var(--xlsx-gridline)',
			cf: cf ?? overlays.cf,
			table: overlays.table,
			bodyRows: vp.rows,
			bodyCols: vp.cols,
		}),
		[
			sheet,
			sheetIndex,
			geometry,
			panes,
			merges,
			store,
			zoom,
			docVersion,
			cf,
			overlays,
			vp.rows,
			vp.cols,
		],
	);

	const g = geometry;
	const clipMark =
		ui.clipboard && ui.clipboard.sheetId === sheet.id ? ui.clipboard.ranges : undefined;
	const editing = ui.edit && ui.edit.sheet === sheetIndex ? ui.edit : null;

	const overlay = (pane: PaneKind, rows: Span, cols: Span) => (
		<>
			<SelectionOverlay
				geometry={g}
				merges={merges}
				rows={rows}
				cols={cols}
				selection={selection}
				showFillHandle={!store.readOnly && !editing}
				copyRanges={clipMark}
				fillPreview={pointer.fillPreview}
				refHighlights={refHighlights}
				hidden={ui.selectedObject !== null}
			/>
			<CellDecorations
				geometry={g}
				rows={rows}
				cols={cols}
				hoverNote={hoverNote}
				hoverLink={hoverLink}
			/>
			{hoverThread && <CommentThreadCard root={hoverThread} geometry={g} rows={rows} cols={cols} />}
			<DrawingsLayer geometry={g} />
			<TraceArrowsLayer geometry={g} />
			{renderOverlay?.({ geometry: g, pane })}
		</>
	);

	return (
		<GridRenderContextReact.Provider value={ctx}>
			<div className='xv-grid' data-sheet={sheet.name}>
				<div
					ref={viewport.scrollerRef}
					className='xv-grid-scroller'
					onScroll={viewport.onScroll}
					onPointerDown={pointer.onPointerDown}
					onPointerMove={pointer.onPointerMove}
					onPointerUp={pointer.onPointerUp}
					onPointerCancel={pointer.onPointerCancel}
					onPointerLeave={() => {
						setHoverNote(null);
						setHoverLink(null);
						setHoverThread(null);
					}}
					onDoubleClick={pointer.onDoubleClick}
					onContextMenu={pointer.onContextMenu}
					role='grid'
					aria-rowcount={vp.extentRows}
					aria-colcount={vp.extentCols}
					aria-label={t('grid.sheet', { name: sheet.name })}
				>
					<GridPanes
						geometry={g}
						panes={panes}
						viewport={vp}
						scrollLeft={viewport.metrics().scrollLeft ?? 0}
						header={{ width: headerWidth, height: headerHeight, rowGutter, colGutter }}
						outline={{ rows: rowOutline, cols: colOutline }}
						selection={selection}
						merges={merges}
						editing={editing}
						onEditorDone={focusGrid}
						overlay={overlay}
					/>
				</div>
				{pointer.guide && (
					<div
						className={
							pointer.guide.orientation === 'v' ? 'xv-resize-guide-v' : 'xv-resize-guide-h'
						}
						style={
							pointer.guide.orientation === 'v'
								? { left: pointer.guide.at }
								: { top: pointer.guide.at }
						}
					>
						<span className='xv-resize-label'>{pointer.guide.label}</span>
					</div>
				)}
				<GridFocusSink ref={sinkRef} onKeyDown={onKeyDown} />
			</div>
		</GridRenderContextReact.Provider>
	);
}
