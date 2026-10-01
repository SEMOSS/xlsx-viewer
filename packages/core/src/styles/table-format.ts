import type { RangeAddress } from '../types/address';
import type { BorderEdge, BorderSpec, DifferentialFormat } from '../types/style';
import type { Table } from '../types/table';
import {
	TABLE_ELEMENT_ORDER,
	type TableStyleDefinition,
	type TableStyleElementType,
} from './table-styles';

interface Region {
	top: number;
	bottom: number;
	left: number;
	right: number;
}

/** Edge of a region's border spec that applies to a cell inside it. */
function edgesFor(border: BorderSpec, region: Region, row: number, col: number): BorderSpec {
	const out: BorderSpec = {};
	const top = row === region.top ? border.top : border.horizontal;
	const bottom = row === region.bottom ? border.bottom : border.horizontal;
	const left = col === region.left ? border.left : border.vertical;
	const right = col === region.right ? border.right : border.vertical;
	if (top) out.top = top;
	if (bottom) out.bottom = bottom;
	if (left) out.left = left;
	if (right) out.right = right;
	return out;
}

function merge(
	into: DifferentialFormat,
	from: DifferentialFormat,
	region: Region,
	row: number,
	col: number,
): void {
	if (from.font) {
		into.font = { ...(into.font ?? {}), ...from.font };
	}
	if (from.fill) {
		into.fill = from.fill;
	}
	if (from.numFmt) {
		into.numFmt = from.numFmt;
	}
	if (from.alignment) {
		into.alignment = { ...(into.alignment ?? {}), ...from.alignment };
	}
	if (from.border) {
		const edges = edgesFor(from.border, region, row, col);
		if (Object.keys(edges).length) {
			into.border = { ...(into.border ?? {}), ...edges };
		}
	}
}

/**
 * Table-style formatting of one cell of a table, merging the style's
 * elements in Excel's precedence order (whole table, stripes, first / last
 * column, header and total rows, corner cells). Returns undefined for
 * cells outside the table or without any table formatting.
 */
export function tableCellFormat(
	table: Table,
	style: TableStyleDefinition,
	row: number,
	col: number,
): DifferentialFormat | undefined {
	const r: RangeAddress = table.ref;
	if (row < r.top || row > r.bottom || col < r.left || col > r.right) {
		return undefined;
	}
	const info = table.style ?? {
		showRowStripes: true,
		showColumnStripes: false,
		showFirstColumn: false,
		showLastColumn: false,
	};
	const headerRows = table.headerRowCount ?? 1;
	const totalRows = table.totalsRowCount ?? 0;
	const dataTop = r.top + headerRows;
	const dataBottom = r.bottom - totalRows;
	const isHeader = headerRows > 0 && row < dataTop;
	const isTotal = totalRows > 0 && row > dataBottom;
	const isData = !isHeader && !isTotal;
	const rowSize = Math.max(1, style.rowStripeSize ?? 1);
	const colSize = Math.max(1, style.colStripeSize ?? 1);

	const regions: Partial<Record<TableStyleElementType, Region>> = {
		wholeTable: r,
	};
	if (isData && info.showColumnStripes) {
		const band = Math.floor((col - r.left) / colSize);
		const left = r.left + band * colSize;
		const region = {
			top: dataTop,
			bottom: dataBottom,
			left,
			right: Math.min(r.right, left + colSize - 1),
		};
		regions[band % 2 === 0 ? 'firstColumnStripe' : 'secondColumnStripe'] = region;
	}
	if (isData && info.showRowStripes) {
		const band = Math.floor((row - dataTop) / rowSize);
		const top = dataTop + band * rowSize;
		const region = {
			top,
			bottom: Math.min(dataBottom, top + rowSize - 1),
			left: r.left,
			right: r.right,
		};
		regions[band % 2 === 0 ? 'firstRowStripe' : 'secondRowStripe'] = region;
	}
	if (info.showLastColumn && col === r.right) {
		regions.lastColumn = { top: r.top, bottom: r.bottom, left: col, right: col };
	}
	if (info.showFirstColumn && col === r.left) {
		regions.firstColumn = { top: r.top, bottom: r.bottom, left: col, right: col };
	}
	if (isHeader) {
		regions.headerRow = { top: r.top, bottom: dataTop - 1, left: r.left, right: r.right };
		if (info.showFirstColumn && col === r.left) {
			regions.firstHeaderCell = { top: row, bottom: row, left: col, right: col };
		}
		if (info.showLastColumn && col === r.right) {
			regions.lastHeaderCell = { top: row, bottom: row, left: col, right: col };
		}
	}
	if (isTotal) {
		regions.totalRow = { top: dataBottom + 1, bottom: r.bottom, left: r.left, right: r.right };
		if (info.showFirstColumn && col === r.left) {
			regions.firstTotalCell = { top: row, bottom: row, left: col, right: col };
		}
		if (info.showLastColumn && col === r.right) {
			regions.lastTotalCell = { top: row, bottom: row, left: col, right: col };
		}
	}

	const out: DifferentialFormat = {};
	let any = false;
	for (const type of TABLE_ELEMENT_ORDER) {
		const region = regions[type];
		const dxf = style.elements[type];
		if (region && dxf) {
			merge(out, dxf, region, row, col);
			any = true;
		}
	}
	return any ? out : undefined;
}

/** True when the edge draws nothing. */
export function isEmptyEdge(edge: BorderEdge | undefined): boolean {
	return !edge || edge.style === 'none';
}
