import type {
	DrawingObject,
	HeaderFooterSections,
	PrintBlockKind,
	PrintLayout,
	PrintPage,
	Worksheet,
} from 'xlsx-viewer-core';
import { columnToLetters } from 'xlsx-viewer-core';

import { anchorRect } from '../geometry/anchor';
import type { Rect, SheetGeometry } from '../geometry/sheet-geometry';
import type { Span } from '../layout/types';
import { PX_PER_PT } from './print-geometry';

/**
 * What to draw for one printed page, in CSS pixels at 96 DPI with the
 * origin at the paper's top-left corner: the cell blocks (each a window on
 * the sheet, like a grid pane), headings, header / footer bands and the
 * floating objects of the body. Pure; the React view maps it onto JSX.
 */

export interface PrintBlockModel {
	key: string;
	kind: PrintBlockKind;
	/** Box on the page. */
	x: number;
	y: number;
	width: number;
	height: number;
	rows: Span;
	cols: Span;
	/** Sheet coordinates (geometry pixels) of the block's top-left corner. */
	originX: number;
	originY: number;
}

/** A row-number or column-letter box (empty text for the corner). */
export interface PrintHeadingCell {
	key: string;
	kind: 'corner' | 'col' | 'row';
	text: string;
	x: number;
	y: number;
	width: number;
	height: number;
}

/** Header or footer: sections between the margins, `offset` from the top (bottom) edge. */
export interface PrintTextBand {
	left: number;
	right: number;
	offset: number;
	sections: HeaderFooterSections;
}

export interface PrintDrawingModel {
	object: DrawingObject;
	/** Rectangle in sheet coordinates (geometry pixels). */
	rect: Rect;
}

export interface PrintPageModel {
	width: number;
	height: number;
	/** Print scale (the geometry zoom). */
	scale: number;
	blocks: PrintBlockModel[];
	headings: PrintHeadingCell[];
	/** Heading font size in pixels. */
	headingFontSize: number;
	gridLines: boolean;
	header: PrintTextBand | null;
	footer: PrintTextBand | null;
	/** Scale of header / footer fonts. */
	headerFooterScale: number;
	/** Floating objects drawn in the body block (in its sheet coordinates). */
	drawings: PrintDrawingModel[];
}

function intersects(a: Rect, b: Rect): boolean {
	return a.x < b.x + b.width && b.x < a.x + a.width && a.y < b.y + b.height && b.y < a.y + a.height;
}

function headingCells(
	page: PrintPage,
	blocks: PrintBlockModel[],
	g: SheetGeometry,
): PrintHeadingCell[] {
	const h = page.headings;
	if (!h) {
		return [];
	}
	const out: PrintHeadingCell[] = [];
	const x0 = page.content.x * PX_PER_PT;
	const y0 = page.content.y * PX_PER_PT;
	const rowWidth = h.rowWidth * PX_PER_PT;
	const colHeight = h.colHeight * PX_PER_PT;
	out.push({
		key: 'corner',
		kind: 'corner',
		text: '',
		x: x0,
		y: y0,
		width: rowWidth,
		height: colHeight,
	});
	const seenCols = new Set<string>();
	const seenRows = new Set<string>();
	for (const b of blocks) {
		const colKey = `${b.x}:${b.cols.start}`;
		if (!seenCols.has(colKey)) {
			seenCols.add(colKey);
			for (let c = b.cols.start; c <= b.cols.end; c++) {
				const w = g.colWidth(c);
				if (w > 0) {
					const x = b.x + g.colLeft(c) - b.originX;
					out.push({
						key: `col:${c}@${b.x}`,
						kind: 'col',
						text: columnToLetters(c),
						x,
						y: y0,
						width: w,
						height: colHeight,
					});
				}
			}
		}
		const rowKey = `${b.y}:${b.rows.start}`;
		if (!seenRows.has(rowKey)) {
			seenRows.add(rowKey);
			for (let r = b.rows.start; r <= b.rows.end; r++) {
				const ht = g.rowHeight(r);
				if (ht > 0) {
					const y = b.y + g.rowTop(r) - b.originY;
					out.push({
						key: `row:${r}@${b.y}`,
						kind: 'row',
						text: String(r + 1),
						x: x0,
						y,
						width: rowWidth,
						height: ht,
					});
				}
			}
		}
	}
	return out;
}

/** The render model of one page of a layout. */
export function buildPrintPageModel(
	layout: PrintLayout,
	page: PrintPage,
	sheet: Worksheet,
	geometry: SheetGeometry,
): PrintPageModel {
	const blocks: PrintBlockModel[] = page.blocks.map((b) => ({
		key: b.kind,
		kind: b.kind,
		x: b.x * PX_PER_PT,
		y: b.y * PX_PER_PT,
		width: b.width * PX_PER_PT,
		height: b.height * PX_PER_PT,
		rows: b.rows,
		cols: b.cols,
		originX: geometry.colLeft(b.cols.start),
		originY: geometry.rowTop(b.rows.start),
	}));
	const body = blocks.find((b) => b.kind === 'body');
	const drawings: PrintDrawingModel[] = [];
	if (body && !layout.draft) {
		const view = { x: body.originX, y: body.originY, width: body.width, height: body.height };
		for (const object of sheet.drawings) {
			if (object.hidden) {
				continue;
			}
			const rect = anchorRect(object.anchor, geometry);
			if (rect.width > 0 && rect.height > 0 && intersects(rect, view)) {
				drawings.push({ object, rect });
			}
		}
	}
	const m = layout.margins;
	const band = (sections: HeaderFooterSections | null, offset: number): PrintTextBand | null =>
		sections
			? {
					left: m.left * PX_PER_PT,
					right: (layout.paper.width - m.right) * PX_PER_PT,
					offset: offset * PX_PER_PT,
					sections,
				}
			: null;
	return {
		width: layout.paper.width * PX_PER_PT,
		height: layout.paper.height * PX_PER_PT,
		scale: layout.scale,
		blocks,
		headings: headingCells(page, blocks, geometry),
		headingFontSize: 11 * PX_PER_PT * layout.scale,
		gridLines: layout.gridLines,
		header: band(page.header, m.header),
		footer: band(page.footer, m.footer),
		headerFooterScale: layout.headerFooterScale,
		drawings,
	};
}
