/**
 * Pixel geometry of a worksheet grid (96 DPI), used to position note boxes
 * and to compute DrawingML transforms from cell anchors.
 */
import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../address/constants';
import {
	columnWidthToPixels,
	defaultColumnWidthPx,
	EMU_PER_PIXEL,
	pointsToPixels,
} from '../styles/column-width';
import type { AnchorPoint, DrawingAnchor } from '../types/drawing';
import type { Worksheet } from '../types/sheet';
import { normalizeCols } from './sheet-elements';

interface Span {
	/** First index of the span. */
	start: number;
	/** Last index (inclusive). */
	end: number;
	/** Size of each index in pixels. */
	size: number;
}

/** Sorted spans of non-default sizes with prefix sums for O(log n) offsets. */
class Axis {
	private readonly starts: number[];
	private readonly prefix: number[];

	constructor(
		private readonly spans: Span[],
		private readonly defaultSize: number,
	) {
		this.starts = spans.map((s) => s.start);
		this.prefix = [];
		let acc = 0;
		for (const s of spans) {
			this.prefix.push(acc);
			acc += (s.end - s.start + 1) * (s.size - defaultSize);
		}
	}

	private spanIndex(i: number): number {
		let lo = 0;
		let hi = this.starts.length - 1;
		let found = -1;
		while (lo <= hi) {
			const mid = (lo + hi) >> 1;
			if (this.starts[mid] <= i) {
				found = mid;
				lo = mid + 1;
			} else {
				hi = mid - 1;
			}
		}
		return found;
	}

	/** Size of one index in pixels. */
	size(i: number): number {
		const k = this.spanIndex(i);
		return k >= 0 && i <= this.spans[k].end ? this.spans[k].size : this.defaultSize;
	}

	/** Pixel offset of the start of index `i`. */
	offset(i: number): number {
		const k = this.spanIndex(i - 1);
		let extra = 0;
		if (k >= 0) {
			const s = this.spans[k];
			extra = this.prefix[k] + (Math.min(s.end, i - 1) - s.start + 1) * (s.size - this.defaultSize);
		}
		return i * this.defaultSize + extra;
	}
}

export class SheetGeometry {
	private readonly cols: Axis;
	private readonly rows: Axis;

	constructor(sheet: Worksheet) {
		const defaultCol = defaultColumnWidthPx(sheet.defaultColWidth, sheet.baseColWidth);
		const colSpans: Span[] = normalizeCols(sheet.cols).map((c) => ({
			start: c.min,
			end: Math.min(c.max, MAX_COL_INDEX),
			size: c.hidden ? 0 : c.width !== undefined ? columnWidthToPixels(c.width) : defaultCol,
		}));
		this.cols = new Axis(colSpans, defaultCol);
		const defaultRow = sheet.zeroHeight
			? 0
			: Math.round(pointsToPixels(sheet.defaultRowHeight || 15));
		const rowSpans: Span[] = [];
		for (const [r, info] of [...sheet.rows].sort((a, b) => a[0] - b[0])) {
			if (r > MAX_ROW_INDEX) {
				continue;
			}
			const size = info.hidden
				? 0
				: info.height !== undefined
					? Math.round(pointsToPixels(info.height))
					: defaultRow;
			if (size !== defaultRow) {
				rowSpans.push({ start: r, end: r, size });
			}
		}
		this.rows = new Axis(rowSpans, defaultRow);
	}

	colWidth(col: number): number {
		return this.cols.size(col);
	}

	rowHeight(row: number): number {
		return this.rows.size(row);
	}

	colLeft(col: number): number {
		return this.cols.offset(col);
	}

	rowTop(row: number): number {
		return this.rows.offset(row);
	}

	/** Absolute EMU position of a cell anchor point. */
	pointEmu(p: AnchorPoint): { x: number; y: number } {
		return {
			x: this.colLeft(p.col) * EMU_PER_PIXEL + p.colOff,
			y: this.rowTop(p.row) * EMU_PER_PIXEL + p.rowOff,
		};
	}

	/** Bounding box of an anchor in EMU. */
	anchorBox(anchor: DrawingAnchor): { x: number; y: number; cx: number; cy: number } {
		if (anchor.type === 'absolute') {
			return { ...anchor.pos, ...anchor.ext };
		}
		const from = this.pointEmu(anchor.from);
		if (anchor.type === 'oneCell') {
			return { ...from, ...anchor.ext };
		}
		const to = this.pointEmu(anchor.to);
		return { ...from, cx: Math.max(0, to.x - from.x), cy: Math.max(0, to.y - from.y) };
	}

	/**
	 * Walk right from (`col`, `offset` px) by `distance` pixels; returns the
	 * column and in-column offset where it ends.
	 */
	walkCols(col: number, offset: number, distance: number): { index: number; offset: number } {
		return walk((i) => this.colWidth(i), MAX_COL_INDEX, col, offset, distance);
	}

	walkRows(row: number, offset: number, distance: number): { index: number; offset: number } {
		return walk((i) => this.rowHeight(i), MAX_ROW_INDEX, row, offset, distance);
	}
}

function walk(
	size: (i: number) => number,
	max: number,
	start: number,
	offset: number,
	distance: number,
): { index: number; offset: number } {
	let i = start;
	let remaining = distance + offset;
	// Hidden (zero-size) indices are stepped over; the guard bounds pathological input.
	for (let guard = 0; guard < 100000 && i < max; guard++) {
		const s = size(i);
		if (remaining < s) {
			break;
		}
		remaining -= s;
		i++;
	}
	return { index: i, offset: Math.max(0, Math.round(remaining)) };
}
