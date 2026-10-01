import type { RangeAddress } from '../types/address';
import type { ColorSpec } from '../types/style';

/** Inclusive span of 0-based rows or columns. */
export interface PrintSpan {
	start: number;
	end: number;
}

/** A rectangle on the paper, in points from the paper's top-left corner. */
export interface PrintRect {
	x: number;
	y: number;
	width: number;
	height: number;
}

/** Margins in points (converted from the sheet's inches). */
export interface PrintMargins {
	left: number;
	right: number;
	top: number;
	bottom: number;
	header: number;
	footer: number;
}

/** Which part of a page a block of cells belongs to. */
export type PrintBlockKind = 'corner' | 'titleRows' | 'titleCols' | 'body';

/**
 * A rectangle of cells printed on a page: the repeated title rows and
 * columns, their corner, or the body. `x / y / width / height` are paper
 * points (already scaled); hidden rows and columns inside the range have no
 * size.
 */
export interface PrintBlock extends PrintRect {
	kind: PrintBlockKind;
	rows: PrintSpan;
	cols: PrintSpan;
	range: RangeAddress;
}

/** Font of a header / footer run (`undefined` fields use the default font). */
export interface HeaderFooterFont {
	name?: string;
	/** Size in points. */
	size?: number;
	bold?: boolean;
	italic?: boolean;
	underline?: 'single' | 'double';
	strike?: boolean;
	vertAlign?: 'superscript' | 'subscript';
	color?: ColorSpec;
}

export interface HeaderFooterRun {
	text: string;
	font: HeaderFooterFont;
}

/** The three sections of a formatted header or footer. */
export interface HeaderFooterSections {
	left: HeaderFooterRun[];
	center: HeaderFooterRun[];
	right: HeaderFooterRun[];
}

/** Row / column headings printed around the cells (Sheet Options > Headings > Print). */
export interface PrintHeadings {
	/** Width of the row-number strip, points. */
	rowWidth: number;
	/** Height of the column-letter strip, points. */
	colHeight: number;
}

export interface PrintPage {
	/** 0-based position in this layout. */
	index: number;
	/** Printed page number (`&P`). */
	pageNumber: number;
	/** Index of the print area the page belongs to. */
	area: number;
	/** Rows and columns of the body. */
	body: RangeAddress;
	/** Title rows / columns repeated on this page, if any. */
	titleRows?: PrintSpan;
	titleCols?: PrintSpan;
	/** Body, title and corner blocks (body last). */
	blocks: PrintBlock[];
	/** Headings, when printed; drawn above / left of the blocks. */
	headings?: PrintHeadings;
	/** Box of headings plus cells, paper points. */
	content: PrintRect;
	header: HeaderFooterSections | null;
	footer: HeaderFooterSections | null;
}

export interface PrintLayout {
	sheetIndex: number;
	sheetName: string;
	/** Paper size in points, orientation applied. */
	paper: { width: number; height: number; paperSize: number; landscape: boolean };
	margins: PrintMargins;
	/** Area between the margins, paper points. */
	printable: PrintRect;
	/** Print scale (1 = 100%). */
	scale: number;
	/** Scale came from Fit to N pages. */
	fitToPage: boolean;
	gridLines: boolean;
	headings: boolean;
	blackAndWhite: boolean;
	draft: boolean;
	pageOrder: 'downThenOver' | 'overThenDown';
	horizontalCentered: boolean;
	verticalCentered: boolean;
	/** The ranges printed (print area, a requested range, or the used range). */
	areas: RangeAddress[];
	titleRows?: PrintSpan;
	titleCols?: PrintSpan;
	pages: PrintPage[];
	/** Pages of this layout (`pages.length`). */
	pageCount: number;
	/** Pages of the whole print job (`&N`): this layout's, or every sheet's when printed together. */
	totalPages: number;
	/** Header / footer text scales with the document. */
	headerFooterScale: number;
	/**
	 * Every page break of the first area (manual and automatic): rows /
	 * columns that start a new page. Used to draw page breaks on the grid.
	 */
	rowBreaks: number[];
	colBreaks: number[];
}

export interface PaginateOptions {
	/** Print this range instead of the print area (Print Selection). */
	range?: RangeAddress;
	/** Several ranges, each starting on new pages (a multi-area selection). */
	ranges?: RangeAddress[];
	/** Print the used range even when the sheet has a print area. */
	ignorePrintArea?: boolean;
	/** For `&F` / `&Z`. */
	fileName?: string;
	filePath?: string;
	/** For `&D` / `&T` (default: now). */
	date?: Date;
	/** Page number of the first page when the sheet uses "Auto" (continuing numbering). */
	firstPageNumber?: number;
	/** Total pages for `&N` (default: this layout's pages). */
	totalPages?: number;
}
