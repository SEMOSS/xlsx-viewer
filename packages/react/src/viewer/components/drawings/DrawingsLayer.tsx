import type { DrawingObject } from 'xlsx-viewer-core';
import { useRef, useState, type PointerEvent as ReactPointerEvent } from 'react';

import { useT } from '../../../i18n';
import { anchorRect, rectToAnchor } from '../../../engine/geometry/anchor';
import type { Rect, SheetGeometry } from '../../../engine/geometry/sheet-geometry';
import { useStoreSelector, useViewerStore } from '../../state/use-store';
import { DRAWING_KIND_KEYS } from './drawing-labels';
import { DrawingObjectView } from './DrawingObjectView';

type Handle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w' | 'move';

const HANDLES: Array<Exclude<Handle, 'move'>> = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];

function resize(r: Rect, h: Handle, dx: number, dy: number): Rect {
	let { x, y, width, height } = r;
	if (h === 'move') {
		return { x: Math.max(0, x + dx), y: Math.max(0, y + dy), width, height };
	}
	if (h.includes('w')) {
		x += dx;
		width -= dx;
	}
	if (h.includes('e')) {
		width += dx;
	}
	if (h.includes('n')) {
		y += dy;
		height -= dy;
	}
	if (h.includes('s')) {
		height += dy;
	}
	if (width < 8) {
		width = 8;
	}
	if (height < 8) {
		height = 8;
	}
	return { x: Math.max(0, x), y: Math.max(0, y), width, height };
}

/**
 * Floating objects of the active sheet, in sheet coordinates. Click selects,
 * dragging moves, handles resize; changes are committed as undoable edits.
 */
export function DrawingsLayer({ geometry }: { geometry: SheetGeometry }) {
	const t = useT();
	const store = useViewerStore();
	const snap = useStoreSelector((s) => s);
	const sheetIndex = snap.ui.activeSheet;
	const sheet = store.workbook.sheets[sheetIndex];
	const selected = snap.ui.selectedObject;
	const [live, setLive] = useState<{ id: string; rect: Rect } | null>(null);
	const drag = useRef<{
		id: string;
		handle: Handle;
		startX: number;
		startY: number;
		rect: Rect;
	} | null>(null);

	if (sheet.drawings.length === 0) {
		return null;
	}

	const begin = (e: ReactPointerEvent, obj: DrawingObject, rect: Rect, handle: Handle) => {
		e.stopPropagation();
		e.preventDefault();
		store.setUi({ selectedObject: obj.id, edit: null });
		if (store.readOnly) {
			return;
		}
		drag.current = { id: obj.id, handle, startX: e.clientX, startY: e.clientY, rect };
		(e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
	};
	const move = (e: ReactPointerEvent) => {
		const d = drag.current;
		if (!d) {
			return;
		}
		setLive({
			id: d.id,
			rect: resize(d.rect, d.handle, e.clientX - d.startX, e.clientY - d.startY),
		});
	};
	const end = (e: ReactPointerEvent) => {
		const d = drag.current;
		drag.current = null;
		if (!d) {
			return;
		}
		const next = resize(d.rect, d.handle, e.clientX - d.startX, e.clientY - d.startY);
		setLive(null);
		if (
			Math.abs(next.x - d.rect.x) < 1 &&
			Math.abs(next.y - d.rect.y) < 1 &&
			Math.abs(next.width - d.rect.width) < 1 &&
			Math.abs(next.height - d.rect.height) < 1
		) {
			return;
		}
		store.transact(store.t(d.handle === 'move' ? 'undo.moveObject' : 'undo.sizeObject'), (tx) => {
			tx.touchSheet(sheetIndex);
			sheet.drawings = sheet.drawings.map((o) =>
				o.id === d.id
					? ({ ...o, anchor: rectToAnchor(next, o.anchor, geometry), dirty: true } as DrawingObject)
					: o,
			);
			sheet.drawingsDirty = true;
		});
	};

	return (
		<div className='xv-drawings' style={{ position: 'absolute', left: 0, top: 0, zIndex: 15 }}>
			{sheet.drawings.map((obj) => {
				if (obj.hidden) {
					return null;
				}
				const rect = live?.id === obj.id ? live.rect : anchorRect(obj.anchor, geometry);
				const isSelected = selected === obj.id;
				return (
					<div
						key={obj.id}
						data-xv-object={obj.id}
						role='img'
						aria-label={obj.description ?? obj.name ?? t(DRAWING_KIND_KEYS[obj.kind])}
						style={{
							position: 'absolute',
							left: rect.x,
							top: rect.y,
							width: rect.width,
							height: rect.height,
							cursor: store.readOnly ? 'default' : 'move',
						}}
						onPointerDown={(e) => begin(e, obj, rect, 'move')}
						onPointerMove={move}
						onPointerUp={end}
						onPointerCancel={end}
					>
						<DrawingObjectView
							object={obj}
							rect={rect}
							workbook={store.workbook}
							sheetIndex={sheetIndex}
							zoom={geometry.zoom}
							revision={snap.docVersion}
						/>
						{isSelected && (
							<div className='xv-object-selection' aria-hidden>
								{!store.readOnly &&
									HANDLES.map((h) => (
										<span
											key={h}
											className={`xv-object-handle xv-handle-${h}`}
											onPointerDown={(e) => begin(e, obj, rect, h)}
											onPointerMove={move}
											onPointerUp={end}
											onPointerCancel={end}
										/>
									))}
							</div>
						)}
					</div>
				);
			})}
		</div>
	);
}
