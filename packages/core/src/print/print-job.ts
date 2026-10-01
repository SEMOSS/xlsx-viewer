import type { Workbook } from '../types/workbook';
import { numberPages, paginateSheet } from './paginate';
import type { PaginateOptions, PrintLayout } from './types';

/**
 * Layouts of several sheets printed as one job (Print Entire Workbook):
 * sheets with an "Auto" first page number continue the page numbers of the
 * sheet before them, and `&N` counts every page of the job.
 */
export function paginateSheets(
	workbook: Workbook,
	sheetIndexes: readonly number[],
	options: Omit<PaginateOptions, 'firstPageNumber' | 'totalPages'> = {},
): PrintLayout[] {
	const layouts = sheetIndexes.map((i) => paginateSheet(workbook, i, options));
	if (layouts.length < 2) {
		return layouts;
	}
	const totalPages = layouts.reduce((n, l) => n + l.pageCount, 0);
	let next = 1;
	for (const layout of layouts) {
		numberPages(workbook, layout, { ...options, firstPageNumber: next, totalPages });
		const last = layout.pages[layout.pages.length - 1];
		if (last) {
			next = last.pageNumber + 1;
		}
	}
	return layouts;
}

/** Total pages of a job. */
export function totalPrintPages(layouts: readonly PrintLayout[]): number {
	return layouts.reduce((n, l) => n + l.pageCount, 0);
}
