import type { ChartObject, PictureObject } from 'xlsx-viewer-core';
import { createChartFromRange, generateId, imageContentType, pixelsToEmu } from 'xlsx-viewer-core';

import { currentRange } from '../../engine/selection/selection';
import type { ViewerStore } from '../state/viewer-store';

/** Delete the selected floating object. */
export function deleteSelectedObject(store: ViewerStore): void {
	const id = store.state.selectedObject;
	if (!id) {
		return;
	}
	const sheetIndex = store.state.activeSheet;
	const sheet = store.activeSheet;
	store.transact(store.t('undo.deleteObject'), (tx) => {
		tx.touchSheet(sheetIndex);
		sheet.drawings = sheet.drawings.filter((d) => d.id !== id);
		sheet.drawingsDirty = true;
	});
	store.setUi({ selectedObject: null });
}

/** Insert a picture at the active cell, at its natural size (max 600 px wide). */
export async function insertPicture(store: ViewerStore, file: File): Promise<void> {
	const data = new Uint8Array(await file.arrayBuffer());
	const ext = (file.name.split('.').pop() ?? 'png').toLowerCase();
	const size = await new Promise<{ w: number; h: number }>((resolve) => {
		const url = URL.createObjectURL(file);
		const img = new Image();
		img.onload = () => {
			resolve({ w: img.naturalWidth, h: img.naturalHeight });
			URL.revokeObjectURL(url);
		};
		img.onerror = () => {
			resolve({ w: 320, h: 240 });
			URL.revokeObjectURL(url);
		};
		img.src = url;
	});
	const scale = Math.min(1, 600 / Math.max(1, size.w));
	const wb = store.workbook;
	let n = wb.media.size + 1;
	while (wb.media.has(`xl/media/image${n}.${ext}`)) {
		n++;
	}
	const path = `xl/media/image${n}.${ext}`;
	const { row, col } = store.selectionOf().active;
	const picture: PictureObject = {
		kind: 'picture',
		id: generateId('pic'),
		name: file.name,
		mediaPath: path,
		anchor: {
			type: 'oneCell',
			from: { row, col, rowOff: 0, colOff: 0 },
			ext: { cx: pixelsToEmu(size.w * scale), cy: pixelsToEmu(size.h * scale) },
		},
		dirty: true,
	};
	const sheetIndex = store.state.activeSheet;
	store.transact(store.t('undo.insertPicture'), (tx) => {
		tx.touchSheet(sheetIndex);
		wb.media.set(path, { path, contentType: imageContentType(ext), data });
		const sheet = wb.sheets[sheetIndex];
		sheet.drawings = [...sheet.drawings, picture];
		sheet.drawingsDirty = true;
	});
	store.setUi({ selectedObject: picture.id });
}

export type ChartKind = Parameters<typeof createChartFromRange>[3];

/** Insert a chart for the selected data, placed to the right of it. */
export function insertChart(store: ViewerStore, type: ChartKind): void {
	const sheetIndex = store.state.activeSheet;
	const wb = store.workbook;
	const range = currentRange(store.selectionOf());
	const chart = createChartFromRange(wb, sheetIndex, range, type);
	const from = { row: range.top, col: range.right + 2, rowOff: 0, colOff: 0 };
	const obj: ChartObject = {
		kind: 'chart',
		id: generateId('chart'),
		name: 'Chart',
		chart,
		anchor: {
			type: 'twoCell',
			from,
			to: { row: from.row + 15, col: from.col + 7, rowOff: 0, colOff: 0 },
		},
		dirty: true,
	};
	store.transact(store.t('undo.insertChart'), (tx) => {
		tx.touchSheet(sheetIndex);
		const sheet = wb.sheets[sheetIndex];
		sheet.drawings = [...sheet.drawings, obj];
		sheet.drawingsDirty = true;
	});
	store.setUi({ selectedObject: obj.id });
}
