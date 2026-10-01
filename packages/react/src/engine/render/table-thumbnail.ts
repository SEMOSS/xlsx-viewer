import type {
	BorderEdge,
	ColorContext,
	Table,
	TableStyleDefinition,
	TableStyleInfo,
} from 'xlsx-viewer-core';
import { resolveColorCss, tableCellFormat } from 'xlsx-viewer-core';

export interface ThumbnailCell {
	x: number;
	y: number;
	width: number;
	height: number;
	fill?: string;
	/** Color of the dash standing in for the cell text. */
	text: string;
	bold: boolean;
}

export interface ThumbnailLine {
	x1: number;
	y1: number;
	x2: number;
	y2: number;
	color: string;
	width: number;
}

export interface TableThumbnail {
	width: number;
	height: number;
	cells: ThumbnailCell[];
	lines: ThumbnailLine[];
}

const ROWS = 5;
const COLS = 5;

function edgeWidth(edge: BorderEdge): number {
	switch (edge.style) {
		case 'medium':
		case 'mediumDashed':
		case 'mediumDashDot':
		case 'mediumDashDotDot':
		case 'double':
			return 1.5;
		case 'thick':
			return 2;
		default:
			return 1;
	}
}

/**
 * Miniature of a table style for the Format as Table gallery: a 5 x 5 table
 * (header row plus four data rows) drawn with the style's fills, borders and
 * text colors, honoring the table style options (stripes, first / last column).
 */
export function buildTableThumbnail(
	style: TableStyleDefinition,
	options: Partial<TableStyleInfo>,
	ctx: ColorContext,
	cell = { width: 11, height: 8 },
): TableThumbnail {
	const table: Table = {
		id: 0,
		name: 'Preview',
		displayName: 'Preview',
		ref: { top: 0, left: 0, bottom: ROWS - 1, right: COLS - 1 },
		headerRowCount: 1,
		totalsRowCount: 0,
		columns: [],
		style: {
			name: style.name,
			showFirstColumn: options.showFirstColumn ?? false,
			showLastColumn: options.showLastColumn ?? false,
			showRowStripes: options.showRowStripes ?? true,
			showColumnStripes: options.showColumnStripes ?? false,
		},
	};
	const cells: ThumbnailCell[] = [];
	const lines: ThumbnailLine[] = [];
	for (let r = 0; r < ROWS; r++) {
		for (let c = 0; c < COLS; c++) {
			const dxf = tableCellFormat(table, style, r, c);
			const x = c * cell.width;
			const y = r * cell.height;
			const fillColor =
				dxf?.fill?.type === 'pattern' ? (dxf.fill.bgColor ?? dxf.fill.fgColor) : undefined;
			cells.push({
				x,
				y,
				width: cell.width,
				height: cell.height,
				fill: fillColor ? resolveColorCss(fillColor, ctx) : undefined,
				text: resolveColorCss(dxf?.font?.color, ctx, '#404040') ?? '#404040',
				bold: dxf?.font?.bold === true,
			});
			const b = dxf?.border;
			const add = (
				edge: BorderEdge | undefined,
				x1: number,
				y1: number,
				x2: number,
				y2: number,
			) => {
				if (edge && edge.style !== 'none') {
					lines.push({
						x1,
						y1,
						x2,
						y2,
						color: resolveColorCss(edge.color, ctx, '#000000') ?? '#000000',
						width: edgeWidth(edge),
					});
				}
			};
			add(b?.top, x, y, x + cell.width, y);
			add(b?.bottom, x, y + cell.height, x + cell.width, y + cell.height);
			add(b?.left, x, y, x, y + cell.height);
			add(b?.right, x + cell.width, y, x + cell.width, y + cell.height);
		}
	}
	return { width: COLS * cell.width, height: ROWS * cell.height, cells, lines };
}
