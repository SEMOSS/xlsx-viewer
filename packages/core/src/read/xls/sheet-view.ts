/**
 * Sheet view records: WINDOW2 (gridlines, headers, zeros, scroll position,
 * frozen flag), SCL (zoom), PANE (split / frozen panes) and SELECTION (one
 * per pane).
 */
import { createDefaultSheetView } from '../../sheet/worksheet';
import type { PaneName, SheetPane, SheetSelection, SheetView } from '../../types/sheet';
import { BiffCursor } from './biff-cursor';
import { RT } from './biff-opcodes';
import type { BiffRecord } from './biff-records';
import { makeRange } from './sheet-context';
import { icvColor } from './xf-records';

const PANES: readonly PaneName[] = ['bottomRight', 'topRight', 'bottomLeft', 'topLeft'];

interface Window2 {
	flags: number;
	top: number;
	left: number;
	gridColor: number;
	zoomPageBreak: number;
	zoomNormal: number;
}

interface RawPane {
	x: number;
	y: number;
	top: number;
	left: number;
	active: PaneName;
}

/** Record types handled by {@link ViewReader.read}. */
export const VIEW_RECORDS: readonly number[] = [RT.WINDOW2, RT.SCL, RT.PANE, RT.SELECTION];

export class ViewReader {
	private window2: Window2 | undefined;
	private zoom: number | undefined;
	private pane: RawPane | undefined;
	private readonly selections: SheetSelection[] = [];

	/** Handle a view record; returns `false` for other record types. */
	read(record: BiffRecord): boolean {
		const cur = new BiffCursor(record.data, record.breaks);
		switch (record.type) {
			case RT.WINDOW2: {
				const flags = cur.u16();
				const top = cur.u16();
				const left = cur.u16();
				const gridColor = cur.remaining >= 2 ? cur.u16() : 64;
				if (cur.remaining >= 6) {
					cur.skip(2);
				}
				const zoomPageBreak = cur.remaining >= 2 ? cur.u16() : 0;
				const zoomNormal = cur.remaining >= 2 ? cur.u16() : 0;
				this.window2 ??= { flags, top, left, gridColor, zoomPageBreak, zoomNormal };
				return true;
			}
			case RT.SCL: {
				const num = cur.u16();
				const den = cur.u16();
				if (num > 0 && den > 0) {
					this.zoom = Math.round((num * 100) / den);
				}
				return true;
			}
			case RT.PANE: {
				const [x, y, top, left] = [cur.u16(), cur.u16(), cur.u16(), cur.u16()];
				this.pane = { x, y, top, left, active: PANES[cur.u8() & 0x03] };
				return true;
			}
			case RT.SELECTION:
				this.selections.push(readSelection(cur));
				return true;
			default:
				return false;
		}
	}

	/** The sheet view from the records read. */
	view(): SheetView {
		const view = createDefaultSheetView();
		const w = this.window2;
		if (!w) {
			return view;
		}
		view.showFormulas = (w.flags & 0x0001) !== 0;
		view.showGridLines = (w.flags & 0x0002) !== 0;
		view.showRowColHeaders = (w.flags & 0x0004) !== 0;
		view.showZeros = (w.flags & 0x0010) !== 0;
		view.rightToLeft = (w.flags & 0x0040) !== 0;
		if (w.flags & 0x0200) {
			view.tabSelected = true;
		}
		const pageBreak = (w.flags & 0x0800) !== 0;
		if (pageBreak) {
			view.view = 'pageBreakPreview';
		}
		if (w.top > 0 || w.left > 0) {
			view.topLeftCell = { row: w.top, col: w.left };
		}
		const grid = !(w.flags & 0x0020) ? icvColor(w.gridColor) : undefined;
		if (grid) {
			view.gridColor = grid;
		}
		const zoom = this.zoom ?? (pageBreak ? w.zoomPageBreak : w.zoomNormal);
		if (zoom > 0) {
			view.zoomScale = Math.max(10, Math.min(400, zoom));
		}
		const pane = this.paneOf(w.flags);
		if (pane) {
			view.pane = pane;
		}
		const selection =
			this.selections.find((s) => s.pane === (pane?.activePane ?? 'topLeft')) ?? this.selections[0];
		if (selection) {
			view.selection = pane
				? selection
				: { activeCell: selection.activeCell, ranges: selection.ranges };
		}
		return view;
	}

	private paneOf(flags: number): SheetPane | undefined {
		const p = this.pane;
		if (!p || (p.x === 0 && p.y === 0)) {
			return undefined;
		}
		const frozen = (flags & 0x0008) !== 0;
		return {
			xSplit: p.x,
			ySplit: p.y,
			topLeftCell: { row: p.top, col: p.left },
			activePane: p.active,
			state: frozen ? ((flags & 0x0100) !== 0 ? 'frozen' : 'frozenSplit') : 'split',
		};
	}
}

function readSelection(cur: BiffCursor): SheetSelection {
	const pane = PANES[cur.u8() & 0x03];
	const activeCell = { row: cur.u16(), col: cur.u16() };
	cur.skip(2);
	const count = Math.min(cur.u16(), Math.floor(cur.remaining / 6));
	const ranges = [];
	for (let i = 0; i < count; i++) {
		const r = makeRange(cur.u16(), cur.u16(), cur.u8(), cur.u8());
		if (r) {
			ranges.push(r);
		}
	}
	if (ranges.length === 0) {
		ranges.push({
			top: activeCell.row,
			left: activeCell.col,
			bottom: activeCell.row,
			right: activeCell.col,
		});
	}
	return { activeCell, ranges, pane };
}
