import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../address/constants';
import type { RangeAddress } from '../types/address';
import type { Worksheet } from '../types/sheet';
import type { Workbook } from '../types/workbook';
import { axisTotal, fitAxisScale, paginateAxis, titlesBefore, type AxisInput } from './axis-breaks';
import { formatHeaderFooter, headerFooterFor } from './header-footer';
import { defaultPrintRowHeight, printColWidth, printRowHeight, spanSize } from './metrics';
import { DEFAULT_MARGINS, paperDimensions, paperSize } from './paper';
import { getPrintArea, getPrintTitles } from './print-names';
import { printableRange } from './print-range';
import type {
	PaginateOptions,
	PrintBlock,
	PrintHeadings,
	PrintLayout,
	PrintMargins,
	PrintPage,
	PrintSpan,
} from './types';

/**
 * Page layout of a sheet as Excel prints it: print area (or the used
 * range), print titles, scale or Fit to N pages, manual and automatic page
 * breaks, page order, centering, headings, headers and footers.
 */

const POINTS_PER_INCH = 72;

function clampArea(area: RangeAddress, used: RangeAddress | null): RangeAddress {
	const out = { ...area };
	if (area.top === 0 && area.bottom === MAX_ROW_INDEX) {
		out.bottom = Math.max(out.top, used?.bottom ?? 0);
	}
	if (area.left === 0 && area.right === MAX_COL_INDEX) {
		out.right = Math.max(out.left, used?.right ?? 0);
	}
	return out;
}

function marginsInPoints(sheet: Worksheet): PrintMargins {
	const m = sheet.pageMargins ?? DEFAULT_MARGINS;
	const pt = (inches: number) => Math.round(Math.max(0, inches) * POINTS_PER_INCH * 1000) / 1000;
	return {
		left: pt(m.left),
		right: pt(m.right),
		top: pt(m.top),
		bottom: pt(m.bottom),
		header: pt(m.header),
		footer: pt(m.footer),
	};
}

/** Headings size at 100%: digits of the last row number plus padding, one default row tall. */
function headingsSize(sheet: Worksheet, areas: RangeAddress[]): PrintHeadings {
	const lastRow = areas.reduce((max, a) => Math.max(max, a.bottom + 1), 1);
	const digits = Math.max(3, String(lastRow).length);
	return {
		rowWidth: (digits * 7 + 10) * 0.75,
		colHeight: defaultPrintRowHeight(sheet) || 15,
	};
}

function areasToPrint(
	workbook: Workbook,
	sheetIndex: number,
	options: PaginateOptions,
): RangeAddress[] {
	const sheet = workbook.sheets[sheetIndex];
	let areas: RangeAddress[] = [];
	if (options.ranges?.length) {
		areas = options.ranges;
	} else if (options.range) {
		areas = [options.range];
	} else if (!options.ignorePrintArea) {
		areas = getPrintArea(workbook, sheetIndex);
	}
	const used = printableRange(sheet, workbook.styles);
	if (areas.length === 0) {
		return used ? [used] : [];
	}
	return areas.map((a) => clampArea(a, used));
}

interface AreaAxes {
	rows: AxisInput;
	cols: AxisInput;
}

/** Lay out the pages of a sheet. */
export function paginateSheet(
	workbook: Workbook,
	sheetIndex: number,
	options: PaginateOptions = {},
): PrintLayout {
	const sheet = workbook.sheets[sheetIndex];
	const setup = sheet.pageSetup ?? {};
	const po = sheet.printOptions ?? {};
	const paper = paperDimensions(setup);
	const margins = marginsInPoints(sheet);
	const printable = {
		x: margins.left,
		y: margins.top,
		width: Math.max(1, paper.width - margins.left - margins.right),
		height: Math.max(1, paper.height - margins.top - margins.bottom),
	};
	const areas = areasToPrint(workbook, sheetIndex, options);
	const titles = getPrintTitles(workbook, sheetIndex);
	const heading = po.headings ? headingsSize(sheet, areas) : undefined;
	const fit = setup.fitToPage === true;
	const rowSize = (r: number) => printRowHeight(sheet, r);
	const colSize = (c: number) => printColWidth(sheet, c);
	const manualRows = fit ? undefined : new Set(sheet.rowBreaks ?? []);
	const manualCols = fit ? undefined : new Set(sheet.colBreaks ?? []);

	const axesOf = (area: RangeAddress): AreaAxes => ({
		rows: {
			span: { start: area.top, end: area.bottom },
			size: rowSize,
			manual: manualRows,
			capacity: (start, scale) => {
				const t = titlesBefore(titles.rows, start);
				const fixed = (t ? spanSize(rowSize, t.start, t.end) : 0) + (heading?.colHeight ?? 0);
				return printable.height - fixed * scale;
			},
		},
		cols: {
			span: { start: area.left, end: area.right },
			size: colSize,
			manual: manualCols,
			capacity: (start, scale) => {
				const t = titlesBefore(titles.cols, start);
				const fixed = (t ? spanSize(colSize, t.start, t.end) : 0) + (heading?.rowWidth ?? 0);
				return printable.width - fixed * scale;
			},
		},
	});
	const axes = areas.map(axesOf);

	let scale = Math.max(10, Math.min(400, setup.scale ?? 100)) / 100;
	if (fit) {
		const wide = setup.fitToWidth ?? 1;
		const tall = setup.fitToHeight ?? 1;
		const titleW = titles.cols ? spanSize(colSize, titles.cols.start, titles.cols.end) : 0;
		const titleH = titles.rows ? spanSize(rowSize, titles.rows.start, titles.rows.end) : 0;
		scale = 1;
		for (const a of axes) {
			if (wide > 0) {
				const fixed = titleW + (heading?.rowWidth ?? 0);
				scale = Math.min(scale, fitAxisScale(a.cols, wide, printable.width, fixed));
			}
			if (tall > 0) {
				const fixed = titleH + (heading?.colHeight ?? 0);
				scale = Math.min(scale, fitAxisScale(a.rows, tall, printable.height, fixed));
			}
		}
	}

	const pages: PrintPage[] = [];
	const bands = axes.map((a) => ({
		rows: axisTotal(a.rows) > 0 ? paginateAxis(a.rows, scale) : [],
		cols: axisTotal(a.cols) > 0 ? paginateAxis(a.cols, scale) : [],
	}));
	bands.forEach((b, area) => {
		const order: Array<[PrintSpan, PrintSpan]> = [];
		if (setup.pageOrder === 'overThenDown') {
			for (const r of b.rows) {
				for (const c of b.cols) {
					order.push([r, c]);
				}
			}
		} else {
			for (const c of b.cols) {
				for (const r of b.rows) {
					order.push([r, c]);
				}
			}
		}
		for (const [rows, cols] of order) {
			pages.push(
				layoutPage({
					index: pages.length,
					pageNumber: pages.length + 1,
					area,
					rows,
					cols,
					titleRows: titlesBefore(titles.rows, rows.start),
					titleCols: titlesBefore(titles.cols, cols.start),
					rowSize,
					colSize,
					scale,
					heading,
					printable,
					centerH: po.horizontalCentered === true,
					centerV: po.verticalCentered === true,
				}),
			);
		}
	});

	const layout: PrintLayout = {
		sheetIndex,
		sheetName: sheet.name,
		paper: {
			width: paper.width,
			height: paper.height,
			paperSize: paperSize(setup.paperSize).code,
			landscape: paper.landscape,
		},
		margins,
		printable,
		scale,
		fitToPage: fit,
		gridLines: po.gridLines === true,
		headings: po.headings === true,
		blackAndWhite: setup.blackAndWhite === true,
		draft: setup.draft === true,
		pageOrder: setup.pageOrder ?? 'downThenOver',
		horizontalCentered: po.horizontalCentered === true,
		verticalCentered: po.verticalCentered === true,
		areas,
		titleRows: titles.rows,
		titleCols: titles.cols,
		pages,
		pageCount: pages.length,
		totalPages: pages.length,
		headerFooterScale: sheet.headerFooter?.scaleWithDoc === false ? 1 : scale,
		rowBreaks: bands[0]?.rows.slice(1).map((s) => s.start) ?? [],
		colBreaks: bands[0]?.cols.slice(1).map((s) => s.start) ?? [],
	};
	numberPages(workbook, layout, options);
	return layout;
}

/**
 * Number the pages and format their headers and footers. A print job over
 * several sheets calls it again to continue page numbers (`firstPageNumber`)
 * and count every page of the job (`totalPages`).
 */
export function numberPages(
	workbook: Workbook,
	layout: PrintLayout,
	options: PaginateOptions = {},
): void {
	const sheet = workbook.sheets[layout.sheetIndex];
	const setup = sheet.pageSetup ?? {};
	const first =
		setup.useFirstPageNumber && setup.firstPageNumber !== undefined
			? setup.firstPageNumber
			: (options.firstPageNumber ?? 1);
	const total = options.totalPages ?? layout.pages.length;
	layout.totalPages = total;
	const date = options.date ?? new Date();
	for (const page of layout.pages) {
		page.pageNumber = first + page.index;
		const codes = headerFooterFor(sheet.headerFooter, page.index, page.pageNumber);
		const ctx = {
			page: page.pageNumber,
			pages: total,
			sheetName: sheet.name,
			fileName: options.fileName,
			filePath: options.filePath,
			date,
		};
		page.header = codes.header ? formatHeaderFooter(codes.header, ctx) : null;
		page.footer = codes.footer ? formatHeaderFooter(codes.footer, ctx) : null;
	}
}

interface PageInput {
	index: number;
	pageNumber: number;
	area: number;
	rows: PrintSpan;
	cols: PrintSpan;
	titleRows?: PrintSpan;
	titleCols?: PrintSpan;
	rowSize: (i: number) => number;
	colSize: (i: number) => number;
	scale: number;
	heading?: PrintHeadings;
	printable: { x: number; y: number; width: number; height: number };
	centerH: boolean;
	centerV: boolean;
}

function block(
	kind: PrintBlock['kind'],
	rows: PrintSpan,
	cols: PrintSpan,
	x: number,
	y: number,
	width: number,
	height: number,
): PrintBlock {
	return {
		kind,
		rows,
		cols,
		range: { top: rows.start, left: cols.start, bottom: rows.end, right: cols.end },
		x,
		y,
		width,
		height,
	};
}

function layoutPage(p: PageInput): PrintPage {
	const s = p.scale;
	const size = (fn: (i: number) => number, span: PrintSpan | undefined) =>
		span ? spanSize(fn, span.start, span.end) * s : 0;
	const tcw = size(p.colSize, p.titleCols);
	const trh = size(p.rowSize, p.titleRows);
	const bw = size(p.colSize, p.cols);
	const bh = size(p.rowSize, p.rows);
	const hw = p.heading ? p.heading.rowWidth * s : 0;
	const hh = p.heading ? p.heading.colHeight * s : 0;
	const width = hw + tcw + bw;
	const height = hh + trh + bh;
	const x0 = p.printable.x + (p.centerH ? Math.max(0, (p.printable.width - width) / 2) : 0);
	const y0 = p.printable.y + (p.centerV ? Math.max(0, (p.printable.height - height) / 2) : 0);
	const blocks: PrintBlock[] = [];
	if (p.titleRows && p.titleCols) {
		blocks.push(block('corner', p.titleRows, p.titleCols, x0 + hw, y0 + hh, tcw, trh));
	}
	if (p.titleRows) {
		blocks.push(block('titleRows', p.titleRows, p.cols, x0 + hw + tcw, y0 + hh, bw, trh));
	}
	if (p.titleCols) {
		blocks.push(block('titleCols', p.rows, p.titleCols, x0 + hw, y0 + hh + trh, tcw, bh));
	}
	blocks.push(block('body', p.rows, p.cols, x0 + hw + tcw, y0 + hh + trh, bw, bh));
	return {
		index: p.index,
		pageNumber: p.pageNumber,
		area: p.area,
		body: { top: p.rows.start, left: p.cols.start, bottom: p.rows.end, right: p.cols.end },
		titleRows: p.titleRows,
		titleCols: p.titleCols,
		blocks,
		headings: p.heading ? { rowWidth: hw, colHeight: hh } : undefined,
		content: { x: x0, y: y0, width, height },
		header: null,
		footer: null,
	};
}
