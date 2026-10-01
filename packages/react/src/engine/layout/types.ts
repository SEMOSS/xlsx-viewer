import type { DifferentialFormat } from 'xlsx-viewer-core';

import type { BorderSegment } from '../render/borders';
import type { CellContent } from '../render/cell-content';
import type { FillCss } from '../render/fills';

/** Inclusive index span. */
export interface Span {
	start: number;
	end: number;
}

/** A painted cell background (sheet pixel coordinates). */
export interface FillBox {
	key: string;
	x: number;
	y: number;
	width: number;
	height: number;
	css: FillCss;
}

/** A painted cell text. `clip` is the area the text may occupy (overflow included). */
export interface TextBox {
	key: string;
	row: number;
	col: number;
	/** The cell (or merged area) box. */
	cell: { x: number; y: number; width: number; height: number };
	/** Clip rectangle, wider than `cell` when text overflows into empty neighbours. */
	clip: { x: number; y: number; width: number; height: number };
	content: CellContent;
	/** Text replaced by `#` characters because the number does not fit. */
	hashes: boolean;
	/** Font scale applied by shrink-to-fit (1 = none). */
	shrinkScale: number;
	/** Conditional-format icon to draw at the left of the cell. */
	icon?: { iconSet: string; index: number };
}

/** Data bar painted inside a cell box (from conditional formatting). */
export interface DataBarBox {
	key: string;
	x: number;
	y: number;
	width: number;
	height: number;
	bar: unknown;
}

/** Overlay a conditional format applies to one cell. */
export interface CfOverlay {
	dxf?: DifferentialFormat;
	/** Color-scale background (CSS color). */
	colorScale?: string;
	dataBar?: unknown;
	icon?: { iconSet: string; index: number };
	hideValue?: boolean;
}

export interface ViewportLayout {
	fills: FillBox[];
	texts: TextBox[];
	borders: BorderSegment[];
	dataBars: DataBarBox[];
	/** Rectangles inside which gridlines must not be drawn (merges, overflow text). */
	gridlineMasks: Array<{ x: number; y: number; width: number; height: number }>;
}
