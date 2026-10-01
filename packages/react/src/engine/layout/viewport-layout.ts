import type {
	BorderSpec,
	Cell,
	DifferentialFormat,
	ResolvedCellStyle,
	StyleResolver,
	Worksheet,
} from 'xlsx-viewer-core';
import { cellKey, formatValue, MAX_COLS } from 'xlsx-viewer-core';

import type { SheetGeometry } from '../geometry/sheet-geometry';
import { BorderCollector } from '../render/borders';
import {
	canOverflow,
	describeCellContent,
	type CellContent,
	type DescribeOptions,
} from '../render/cell-content';
import { canvasFont } from '../render/fonts';
import { fillToCss } from '../render/fills';
import { measureLines, measureText } from '../render/text-measure';
import { applyDxf } from './apply-dxf';
import { applyTableDxf } from './overlays';
import { MergeIndex, styleIdAt } from './style-at';
import type { CfOverlay, FillBox, Span, TextBox, ViewportLayout } from './types';

export interface LayoutParams {
	sheet: Worksheet;
	geometry: SheetGeometry;
	/** Visible rows / columns (inclusive). */
	rows: Span;
	cols: Span;
	styles: StyleResolver;
	describe: DescribeOptions;
	merges: MergeIndex;
	cf?: (row: number, col: number) => CfOverlay | undefined;
	/** Table-style formatting (applied beneath the cell's own formatting). */
	table?: (row: number, col: number) => DifferentialFormat | undefined;
	/** Columns scanned beyond the visible span for text overflowing into view. */
	overflowLookaround?: number;
	/** Fill color used to hide gridlines inside merged cells without a fill. */
	canvasColor?: string;
}

/** Horizontal text padding in px at 100% zoom. */
export const CELL_PADDING_X = 2;

interface Prepared {
	cell: Cell | undefined;
	style: ResolvedCellStyle;
	overlay?: CfOverlay;
}

function isOccupied(cell: Cell | undefined): boolean {
	return cell !== undefined && (cell.value !== null || cell.formula !== undefined);
}

/**
 * Lay out the visible part of a sheet: backgrounds, borders, texts (with
 * overflow into empty neighbours, hashes for numbers that do not fit,
 * shrink-to-fit), merged areas and gridline masks. Pure function of its
 * inputs, in sheet pixel coordinates.
 */
export function layoutViewport(p: LayoutParams): ViewportLayout {
	const { sheet, geometry: g, styles } = p;
	const zoom = p.describe.zoom;
	const padX = Math.max(1, CELL_PADDING_X * zoom);
	const look = p.overflowLookaround ?? 24;
	const fills: FillBox[] = [];
	const texts: TextBox[] = [];
	const masks: ViewportLayout['gridlineMasks'] = [];
	const dataBars: ViewportLayout['dataBars'] = [];
	const borders = new BorderCollector(p.describe.colorCtx, zoom, p.describe.defaultTextColor);
	const doneMerges = new Set<string>();
	const scanStart = Math.max(0, p.cols.start - look);
	const scanEnd = Math.min(MAX_COLS - 1, p.cols.end + look);

	const normal = styles.resolve(0);
	const prepare = (row: number, col: number, cell: Cell | undefined): Prepared => {
		let style = styles.resolve(styleIdAt(sheet, row, col, cell));
		const tableDxf = p.table?.(row, col);
		if (tableDxf) {
			style = applyTableDxf(style, tableDxf, normal);
		}
		const overlay = p.cf?.(row, col);
		return { cell, style: overlay?.dxf ? applyDxf(style, overlay.dxf) : style, overlay };
	};

	const addBorders = (
		border: BorderSpec,
		row: number,
		col: number,
		x: number,
		y: number,
		w: number,
		h: number,
	): void => {
		borders.add('h', row, col, y, x, x + w, border.top);
		borders.add('h', row + 1, col, y + h, x, x + w, border.bottom);
		borders.add('v', col, row, x, y, y + h, border.left);
		borders.add('v', col + 1, row, x + w, y, y + h, border.right);
	};

	const pushFill = (
		key: string,
		x: number,
		y: number,
		w: number,
		h: number,
		prep: Prepared,
		force?: string,
	): void => {
		const css = prep.overlay?.colorScale
			? { backgroundColor: prep.overlay.colorScale }
			: fillToCss(prep.style.fill, p.describe.colorCtx);
		if (css) {
			fills.push({ key, x, y, width: w, height: h, css });
		} else if (force) {
			fills.push({ key, x, y, width: w, height: h, css: { backgroundColor: force } });
		}
		if (prep.overlay?.dataBar) {
			dataBars.push({ key, x, y, width: w, height: h, bar: prep.overlay.dataBar });
		}
	};

	for (let row = p.rows.start; row <= p.rows.end; row++) {
		const h = g.rowHeight(row);
		if (h <= 0) {
			continue;
		}
		const y = g.rowTop(row);
		// Cells of this row across the scan window, plus occupancy for overflow.
		const rowCells = new Map<number, Cell>();
		for (let col = scanStart; col <= scanEnd; col++) {
			const cell = sheet.cells.get(cellKey(row, col));
			if (cell) {
				rowCells.set(col, cell);
			}
		}
		const occupied = (col: number): boolean => {
			if (col < scanStart || col > scanEnd) {
				return isOccupied(sheet.cells.get(cellKey(row, col)));
			}
			return isOccupied(rowCells.get(col)) || p.merges.find(row, col) !== undefined;
		};

		for (let col = scanStart; col <= scanEnd; col++) {
			const w = g.colWidth(col);
			const visibleCol = col >= p.cols.start && col <= p.cols.end;
			if (w <= 0 && !visibleCol) {
				continue;
			}
			const merge = p.merges.find(row, col);
			if (merge) {
				if (!visibleCol) {
					continue;
				}
				const mkey = `${merge.top}:${merge.left}`;
				if (doneMerges.has(mkey)) {
					// Perimeter borders of the part of the merge inside the viewport.
					const cell = sheet.cells.get(cellKey(row, col));
					const prep = prepare(row, col, cell);
					const x = g.colLeft(col);
					const b = prep.style.border;
					if (row === merge.top) borders.add('h', row, col, y, x, x + w, b.top);
					if (row === merge.bottom) borders.add('h', row + 1, col, y + h, x, x + w, b.bottom);
					if (col === merge.left) borders.add('v', col, row, x, y, y + h, b.left);
					if (col === merge.right) borders.add('v', col + 1, row, x + w, y, y + h, b.right);
					continue;
				}
				doneMerges.add(mkey);
				const rect = g.rangeRect(merge);
				const anchor = sheet.cells.get(cellKey(merge.top, merge.left));
				const prep = prepare(merge.top, merge.left, anchor);
				pushFill(
					`m:${mkey}`,
					rect.x,
					rect.y,
					rect.width,
					rect.height,
					prep,
					p.canvasColor ?? '#ffffff',
				);
				masks.push({ x: rect.x, y: rect.y, width: rect.width - 1, height: rect.height - 1 });
				const self = prepare(row, col, sheet.cells.get(cellKey(row, col)));
				const x = g.colLeft(col);
				const b = self.style.border;
				if (row === merge.top) borders.add('h', row, col, y, x, x + w, b.top);
				if (row === merge.bottom) borders.add('h', row + 1, col, y + h, x, x + w, b.bottom);
				if (col === merge.left) borders.add('v', col, row, x, y, y + h, b.left);
				if (col === merge.right) borders.add('v', col + 1, row, x + w, y, y + h, b.right);
				const content = prep.overlay?.hideValue
					? null
					: describeCellContent(anchor, prep.style, p.describe);
				if (content || prep.overlay?.icon) {
					texts.push(
						makeTextBox(
							`t:${merge.top}:${merge.left}`,
							merge.top,
							merge.left,
							rect,
							rect,
							content,
							prep,
							padX,
						),
					);
				}
				continue;
			}

			const cell = rowCells.get(col);
			const x = g.colLeft(col);
			let prep: Prepared | undefined;
			if (visibleCol && w > 0) {
				prep = prepare(row, col, cell);
				pushFill(`f:${row}:${col}`, x, y, w, h, prep);
				addBorders(prep.style.border, row, col, x, y, w, h);
			}
			if (!cell || w <= 0) {
				continue;
			}
			prep ??= prepare(row, col, cell);
			const content = prep.overlay?.hideValue
				? null
				: describeCellContent(cell, prep.style, p.describe);
			if (!content && !prep.overlay?.icon) {
				continue;
			}
			const cellBox = { x, y, width: w, height: h };
			let clip = cellBox;
			if (content && canOverflow(content)) {
				const font = canvasFont(content.font);
				const needed = measureText(content.text, font) + padX * 2 + content.indent;
				if (needed > w) {
					let left = col;
					let right = col;
					let widthLeft = 0;
					let widthRight = 0;
					const extra = needed - w;
					const wantLeft =
						content.hAlign === 'right' ? extra : content.hAlign === 'center' ? extra / 2 : 0;
					const wantRight =
						content.hAlign === 'left' ? extra : content.hAlign === 'center' ? extra / 2 : 0;
					while (widthRight < wantRight && right + 1 < MAX_COLS && !occupied(right + 1)) {
						right++;
						widthRight += g.colWidth(right);
					}
					while (widthLeft < wantLeft && left - 1 >= 0 && !occupied(left - 1)) {
						left--;
						widthLeft += g.colWidth(left);
					}
					if (left !== col || right !== col) {
						const cx = g.colLeft(left);
						clip = { x: cx, y, width: g.colLeft(right) + g.colWidth(right) - cx, height: h };
					}
				}
			}
			const clipVisible =
				clip.x + clip.width > g.colLeft(p.cols.start) && clip.x < g.colLeft(p.cols.end + 1);
			if (!clipVisible) {
				continue;
			}
			if (clip !== cellBox) {
				masks.push({ x: clip.x, y, width: clip.width - 1, height: h - 1 });
			}
			texts.push(makeTextBox(`t:${row}:${col}`, row, col, cellBox, clip, content, prep, padX));
		}
	}

	return { fills, texts, borders: borders.segments(), dataBars, gridlineMasks: masks };
}

function makeTextBox(
	key: string,
	row: number,
	col: number,
	cell: TextBox['cell'],
	clip: TextBox['clip'],
	content: CellContent | null,
	prep: Prepared,
	padX: number,
): TextBox {
	let hashes = false;
	let shrinkScale = 1;
	let finalContent = content;
	if (content) {
		const font = canvasFont(content.font);
		const available = Math.max(0, cell.width - padX * 2 - content.indent);
		if (content.kind === 'number' && content.shrink && !content.invalid) {
			// Shrink to fit wins over "####": Excel scales the number down instead.
			const width = measureText(
				content.repeatFill ? content.repeatFill.before + content.repeatFill.after : content.text,
				font,
			);
			if (width > available && width > 0) {
				shrinkScale = Math.max(0.1, available / width);
			}
		} else if (content.kind === 'number') {
			// Accounting formats fill the gap with a repeated character; only the digits must fit.
			const width = measureText(
				content.repeatFill ? content.repeatFill.before + content.repeatFill.after : content.text,
				font,
			);
			if (content.invalid || width > available) {
				const value = prep.cell?.value;
				const general = /^general$/i.test(prep.style.numFmt);
				if (general && typeof value === 'number' && !content.invalid) {
					const digit = Math.max(1, measureText('0', font));
					const maxChars = Math.max(1, Math.floor(available / digit));
					const reformatted = formatValue(value, 'General', { generalMaxChars: maxChars });
					if (measureText(reformatted.text, font) <= available) {
						finalContent = { ...content, text: reformatted.text };
					} else {
						hashes = true;
					}
				} else {
					hashes = true;
				}
			}
			if (hashes) {
				const hashWidth = Math.max(1, measureText('#', font));
				finalContent = {
					...content,
					text: '#'.repeat(Math.max(1, Math.floor(available / hashWidth))),
				};
			}
		} else if (content.shrink) {
			const width = measureLines(content.text, font);
			if (width > available && width > 0) {
				shrinkScale = Math.max(0.1, available / width);
			}
		}
	}
	return {
		key,
		row,
		col,
		cell,
		clip,
		content: finalContent ?? emptyContent(),
		hashes,
		shrinkScale,
		icon: prep.overlay?.icon,
	};
}

function emptyContent(): CellContent {
	return {
		kind: 'text',
		text: '',
		font: { fontFamily: 'sans-serif', fontSize: 14, fontWeight: 400, fontStyle: 'normal' },
		color: 'inherit',
		hAlign: 'left',
		vAlign: 'bottom',
		wrap: false,
		shrink: false,
		indent: 0,
		rotation: 0,
		vertical: false,
		invalid: false,
		accountingUnderline: false,
	};
}
