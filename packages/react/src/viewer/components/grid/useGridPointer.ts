import type { Hyperlink, RangeAddress } from 'xlsx-viewer-core';
import { useCallback, useEffect, useRef, useState, type PointerEvent, type RefObject } from 'react';

import { hitTest, type GridHit, type GridMetrics } from '../../../engine/geometry/hit-test';
import {
	currentRange,
	selectAll,
	wholeColumns,
	wholeRows,
} from '../../../engine/selection/selection';
import {
	autoFitColumns,
	autoFitRows,
	resizeColumns,
	resizeRows,
} from '../../commands/layout-commands';
import { fillFromHandle } from '../../commands/fill-commands';
import { beginEdit, commitEdit } from '../../commands/edit-commands';
import { tryPointReference, pointDragTo } from '../../commands/point-commands';
import { followHyperlink, hyperlinkForCell } from '../../commands/hyperlink-commands';
import {
	dragSelectionTo,
	selectCell,
	selectColumns,
	selectRows,
} from '../../commands/selection-commands';
import type { ViewerStore } from '../../state/viewer-store';
import { fillDragTarget, hoverCursor } from './pointer-targets';

/** Excel: click a link to follow it, click and hold to just select the cell. */
const LINK_HOLD_MS = 500;

type Drag =
	| { kind: 'cells'; link?: { link: Hyperlink; row: number; col: number; at: number } }
	| { kind: 'cols'; anchor: number }
	| { kind: 'rows'; anchor: number }
	| { kind: 'colResize'; col: number; startX: number; startWidth: number }
	| { kind: 'rowResize'; row: number; startY: number; startHeight: number }
	| { kind: 'fill'; source: RangeAddress }
	| { kind: 'point' };

export interface ResizeGuide {
	orientation: 'v' | 'h';
	/** Viewport coordinate of the guide line. */
	at: number;
	/** Size label shown while dragging. */
	label: string;
}

interface Options {
	store: ViewerStore;
	scrollerRef: RefObject<HTMLDivElement | null>;
	metrics: () => GridMetrics;
	scrollBy: (dx: number, dy: number) => void;
	onContextMenu?: (e: { clientX: number; clientY: number; hit: GridHit }) => void;
	focusGrid: () => void;
	/** Pointer hovering (no button pressed): the hit under the pointer, or null when leaving. */
	onHover?: (hit: GridHit | null) => void;
}

const EDGE = 24;

/** Pointer interactions of the grid: selection, resizing, the fill handle and double-click actions. */
export function useGridPointer(opts: Options) {
	const optsRef = useRef(opts);
	optsRef.current = opts;
	const drag = useRef<Drag | null>(null);
	const lastPoint = useRef<{ px: number; py: number } | null>(null);
	const raf = useRef<number | null>(null);
	const [guide, setGuide] = useState<ResizeGuide | null>(null);
	const [fillPreview, setFillPreview] = useState<RangeAddress | null>(null);

	const local = (e: { clientX: number; clientY: number }) => {
		const el = optsRef.current.scrollerRef.current;
		const rect = el?.getBoundingClientRect();
		return { px: e.clientX - (rect?.left ?? 0), py: e.clientY - (rect?.top ?? 0) };
	};

	const applyDrag = useCallback((px: number, py: number) => {
		const { store, metrics } = optsRef.current;
		const d = drag.current;
		if (!d) {
			return;
		}
		const m = metrics();
		const hit = hitTest(m, Math.max(px, m.headerWidth), Math.max(py, m.headerHeight));
		switch (d.kind) {
			case 'cells':
				dragSelectionTo(store, hit.row, hit.col);
				break;
			case 'point':
				pointDragTo(store, hit.row, hit.col);
				break;
			case 'cols': {
				const sel = store.selectionOf();
				store.setSelection({
					...sel,
					ranges: [...sel.ranges.slice(0, -1), wholeColumns(d.anchor, hit.col)],
				});
				break;
			}
			case 'rows': {
				const sel = store.selectionOf();
				store.setSelection({
					...sel,
					ranges: [...sel.ranges.slice(0, -1), wholeRows(d.anchor, hit.row)],
				});
				break;
			}
			case 'colResize': {
				const width = Math.max(0, d.startWidth + (px - d.startX));
				const size = Math.round(width / m.geometry.zoom);
				setGuide({ orientation: 'v', at: px, label: store.t('grid.resizeWidth', { size }) });
				break;
			}
			case 'rowResize': {
				const height = Math.max(0, d.startHeight + (py - d.startY));
				const size = Math.round(height / m.geometry.zoom);
				setGuide({ orientation: 'h', at: py, label: store.t('grid.resizeHeight', { size }) });
				break;
			}
			case 'fill':
				setFillPreview(fillDragTarget(d.source, hit));
				break;
		}
	}, []);

	const autoScrollLoop = useCallback(() => {
		raf.current = null;
		const pt = lastPoint.current;
		const d = drag.current;
		const el = optsRef.current.scrollerRef.current;
		if (!pt || !d || !el || d.kind === 'colResize' || d.kind === 'rowResize') {
			return;
		}
		const m = optsRef.current.metrics();
		const minX = m.headerWidth + m.panes.frozenWidth;
		const minY = m.headerHeight + m.panes.frozenHeight;
		let dx = 0;
		let dy = 0;
		if (pt.px > el.clientWidth - EDGE) {
			dx = Math.min(60, pt.px - (el.clientWidth - EDGE));
		} else if (pt.px < minX && d.kind !== 'rows') {
			dx = -Math.min(60, minX - pt.px);
		}
		if (pt.py > el.clientHeight - EDGE) {
			dy = Math.min(60, pt.py - (el.clientHeight - EDGE));
		} else if (pt.py < minY && d.kind !== 'cols') {
			dy = -Math.min(60, minY - pt.py);
		}
		if (dx !== 0 || dy !== 0) {
			optsRef.current.scrollBy(dx, dy);
			applyDrag(pt.px, pt.py);
			raf.current = requestAnimationFrame(autoScrollLoop);
		}
	}, [applyDrag]);

	const onPointerDown = useCallback(
		(e: PointerEvent<HTMLDivElement>) => {
			const { store, metrics, focusGrid } = optsRef.current;
			const el = optsRef.current.scrollerRef.current;
			if (!el || (e.button !== 0 && e.button !== 2)) {
				return;
			}
			const { px, py } = local(e);
			if (px >= el.clientWidth || py >= el.clientHeight) {
				return; // scrollbar
			}
			const target = e.target as HTMLElement;
			if (target.closest('[data-xv-object]')) {
				return; // floating objects handle their own pointer events
			}
			// The focus sink keeps the keyboard: cancelling the press stops the browser from
			// moving focus to the clicked element (click and dblclick still follow).
			if (e.button === 2) {
				if (!store.state.edit) {
					focusGrid();
					e.preventDefault();
				}
				return; // the context menu handler selects
			}
			const m = metrics();
			const hit = hitTest(m, px, py);
			const edit = store.state.edit;

			if (target.closest('[data-fill-handle]') && !store.readOnly) {
				if (edit) {
					commitEdit(store, { move: 'none' });
				}
				focusGrid();
				e.preventDefault();
				drag.current = { kind: 'fill', source: currentRange(store.selectionOf()) };
			} else if (
				hit.kind === 'cell' &&
				edit &&
				tryPointReference(store, hit.row, hit.col, e.shiftKey)
			) {
				drag.current = { kind: 'point' };
				e.preventDefault();
			} else {
				if (edit) {
					if (!commitEdit(store, { move: 'none' })) {
						return;
					}
				}
				focusGrid();
				e.preventDefault();
				switch (hit.kind) {
					case 'corner':
						store.setSelection({
							active: { row: 0, col: 0 },
							anchor: { row: 0, col: 0 },
							ranges: [selectAll()],
						});
						return;
					case 'colHeader':
						if (hit.colResize !== undefined && !store.readOnly) {
							drag.current = {
								kind: 'colResize',
								col: hit.colResize,
								startX: px,
								startWidth: m.geometry.colWidth(hit.colResize),
							};
						} else {
							const anchor = e.shiftKey ? store.selectionOf().anchor.col : hit.col;
							if (e.shiftKey) {
								const sel = store.selectionOf();
								store.setSelection({
									...sel,
									ranges: [...sel.ranges.slice(0, -1), wholeColumns(anchor, hit.col)],
								});
							} else {
								selectColumns(store, hit.col, hit.col, { add: e.ctrlKey || e.metaKey });
							}
							drag.current = { kind: 'cols', anchor };
						}
						break;
					case 'rowHeader':
						if (hit.rowResize !== undefined && !store.readOnly) {
							drag.current = {
								kind: 'rowResize',
								row: hit.rowResize,
								startY: py,
								startHeight: m.geometry.rowHeight(hit.rowResize),
							};
						} else {
							const anchor = e.shiftKey ? store.selectionOf().anchor.row : hit.row;
							if (e.shiftKey) {
								const sel = store.selectionOf();
								store.setSelection({
									...sel,
									ranges: [...sel.ranges.slice(0, -1), wholeRows(anchor, hit.row)],
								});
							} else {
								selectRows(store, hit.row, hit.row, { add: e.ctrlKey || e.metaKey });
							}
							drag.current = { kind: 'rows', anchor };
						}
						break;
					case 'cell': {
						const plain = !e.shiftKey && !e.ctrlKey && !e.metaKey && !e.altKey;
						const link = plain ? hyperlinkForCell(store, hit.row, hit.col) : undefined;
						selectCell(store, hit.row, hit.col, {
							extend: e.shiftKey,
							add: e.ctrlKey || e.metaKey,
						});
						drag.current = {
							kind: 'cells',
							link: link && { link, row: hit.row, col: hit.col, at: e.timeStamp },
						};
						break;
					}
				}
			}
			lastPoint.current = { px, py };
			el.setPointerCapture(e.pointerId);
			applyDrag(px, py);
		},
		[applyDrag],
	);

	const onPointerMove = useCallback(
		(e: PointerEvent<HTMLDivElement>) => {
			const { px, py } = local(e);
			const el = optsRef.current.scrollerRef.current;
			if (!drag.current) {
				if (el) {
					const hit = hitTest(optsRef.current.metrics(), px, py);
					optsRef.current.onHover?.(hit);
					const cursor = hoverCursor(
						hit,
						!!(e.target as HTMLElement).closest('[data-fill-handle]'),
						() => !!hyperlinkForCell(optsRef.current.store, hit.row, hit.col),
					);
					if (el.style.cursor !== cursor) {
						el.style.cursor = cursor;
					}
				}
				return;
			}
			lastPoint.current = { px, py };
			applyDrag(px, py);
			if (raf.current === null) {
				raf.current = requestAnimationFrame(autoScrollLoop);
			}
		},
		[applyDrag, autoScrollLoop],
	);

	const finish = useCallback((e: PointerEvent<HTMLDivElement>) => {
		const d = drag.current;
		drag.current = null;
		if (raf.current !== null) {
			cancelAnimationFrame(raf.current);
			raf.current = null;
		}
		const { store } = optsRef.current;
		const { px, py } = local(e);
		if (d?.kind === 'colResize') {
			resizeColumns(store, d.col, Math.max(0, d.startWidth + (px - d.startX)));
			setGuide(null);
		} else if (d?.kind === 'rowResize') {
			resizeRows(store, d.row, Math.max(0, d.startHeight + (py - d.startY)));
			setGuide(null);
		} else if (
			d?.kind === 'cells' &&
			d.link &&
			e.type === 'pointerup' &&
			e.timeStamp - d.link.at < LINK_HOLD_MS
		) {
			const hit = hitTest(optsRef.current.metrics(), px, py);
			if (hit.kind === 'cell' && hit.row === d.link.row && hit.col === d.link.col) {
				followHyperlink(store, d.link.link);
			}
		} else if (d?.kind === 'fill') {
			setFillPreview((preview) => {
				if (preview) {
					fillFromHandle(store, d.source, preview);
				}
				return null;
			});
		}
	}, []);

	const onDoubleClick = useCallback((e: React.MouseEvent<HTMLDivElement>) => {
		const { store, metrics } = optsRef.current;
		const { px, py } = local(e);
		const hit = hitTest(metrics(), px, py);
		if ((e.target as HTMLElement).closest('[data-fill-handle]')) {
			fillFromHandle(store, currentRange(store.selectionOf()), 'auto');
			return;
		}
		if (hit.kind === 'colHeader' && hit.colResize !== undefined) {
			autoFitColumns(store, hit.colResize);
		} else if (hit.kind === 'rowHeader' && hit.rowResize !== undefined) {
			autoFitRows(store, hit.rowResize);
		} else if (hit.kind === 'cell' && !store.state.edit) {
			beginEdit(store, { mode: 'edit' });
		}
	}, []);

	const onContextMenu = useCallback((e: React.MouseEvent<HTMLDivElement>) => {
		const { store, metrics, onContextMenu: open } = optsRef.current;
		const { px, py } = local(e);
		const hit = hitTest(metrics(), px, py);
		e.preventDefault();
		const sel = store.selectionOf();
		const inside = sel.ranges.some(
			(r) => hit.row >= r.top && hit.row <= r.bottom && hit.col >= r.left && hit.col <= r.right,
		);
		if (!inside) {
			if (hit.kind === 'colHeader') {
				selectColumns(store, hit.col, hit.col);
			} else if (hit.kind === 'rowHeader') {
				selectRows(store, hit.row, hit.row);
			} else if (hit.kind === 'cell') {
				selectCell(store, hit.row, hit.col);
			}
		}
		open?.({ clientX: e.clientX, clientY: e.clientY, hit });
	}, []);

	useEffect(
		() => () => {
			if (raf.current !== null) {
				cancelAnimationFrame(raf.current);
			}
		},
		[],
	);

	return {
		onPointerDown,
		onPointerMove,
		onPointerUp: finish,
		onPointerCancel: finish,
		onDoubleClick,
		onContextMenu,
		guide,
		fillPreview,
	};
}
