import { MARGIN_PRESETS, type MarginPreset } from '../print/paper';
import type { HeaderFooter, PageMargins, PageSetup, PrintOptions, Worksheet } from '../types/sheet';
import type { Workbook } from '../types/workbook';
import type { EditTransaction } from './transaction';

export * from './page-setup-ops-breaks';
export * from './page-setup-ops-names';

/**
 * Page Layout > Page Setup and the Page Setup dialog as undoable edits:
 * orientation, paper size, margins, scaling, print area, print titles,
 * page breaks, gridlines / headings printing and headers / footers.
 * Changes on a protected sheet are allowed, as in Excel.
 */

export type PageOrientation = 'portrait' | 'landscape';

function sheetFor(tx: EditTransaction, workbook: Workbook, sheetIndex: number): Worksheet {
	tx.touchSheet(sheetIndex, false, 'view');
	return workbook.sheets[sheetIndex];
}

function compactObject<T extends object>(obj: T): T | undefined {
	const out = { ...obj };
	for (const key of Object.keys(out) as Array<keyof T>) {
		if (out[key] === undefined) {
			delete out[key];
		}
	}
	return Object.keys(out).length ? out : undefined;
}

/** Merge `patch` into the page setup (`undefined` values remove the setting). */
export function setPageSetup(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	patch: Partial<PageSetup>,
): void {
	const sheet = sheetFor(tx, workbook, sheetIndex);
	sheet.pageSetup = compactObject({ ...sheet.pageSetup, ...patch });
}

/** Merge `patch` into the print options. */
export function setPrintOptions(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	patch: Partial<PrintOptions>,
): void {
	const sheet = sheetFor(tx, workbook, sheetIndex);
	sheet.printOptions = compactObject({ ...sheet.printOptions, ...patch });
}

export function setOrientation(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	orientation: PageOrientation,
): void {
	setPageSetup(tx, workbook, sheetIndex, { orientation });
}

/** Paper size code (`PAPER_SIZES`). */
export function setPaperSize(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	paperSize: number,
): void {
	setPageSetup(tx, workbook, sheetIndex, { paperSize: Math.max(1, Math.trunc(paperSize)) });
}

/** Margins from a preset (Normal, Wide, Narrow) or custom values in inches. */
export function setMargins(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	margins: MarginPreset | PageMargins,
): void {
	const m = typeof margins === 'string' ? MARGIN_PRESETS[margins] : margins;
	const inches = (v: number) => (Number.isFinite(v) ? Math.max(0, v) : 0);
	const sheet = sheetFor(tx, workbook, sheetIndex);
	sheet.pageMargins = {
		left: inches(m.left),
		right: inches(m.right),
		top: inches(m.top),
		bottom: inches(m.bottom),
		header: inches(m.header),
		footer: inches(m.footer),
	};
}

/** Adjust to N% normal size (10 to 400); turns Fit to off. */
export function setScale(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	percent: number,
): void {
	const scale = Math.max(10, Math.min(400, Math.round(percent)));
	setPageSetup(tx, workbook, sheetIndex, { scale, fitToPage: false });
}

/**
 * Fit to `width` pages wide by `height` tall (0 = automatic). Both automatic
 * turns fitting off and returns to 100%.
 */
export function setFitTo(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	width: number,
	height: number,
): void {
	const w = Math.max(0, Math.trunc(width));
	const h = Math.max(0, Math.trunc(height));
	if (w === 0 && h === 0) {
		setPageSetup(tx, workbook, sheetIndex, {
			fitToPage: false,
			fitToWidth: undefined,
			fitToHeight: undefined,
			scale: undefined,
		});
		return;
	}
	setPageSetup(tx, workbook, sheetIndex, { fitToPage: true, fitToWidth: w, fitToHeight: h });
}

export function setPrintGridlines(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	on: boolean,
): void {
	setPrintOptions(tx, workbook, sheetIndex, { gridLines: on || undefined });
}

export function setPrintHeadings(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	on: boolean,
): void {
	setPrintOptions(tx, workbook, sheetIndex, { headings: on || undefined });
}

/** Replace the headers and footers (empty texts are dropped). */
export function setHeaderFooter(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	hf: HeaderFooter | undefined,
): void {
	const sheet = sheetFor(tx, workbook, sheetIndex);
	if (!hf) {
		sheet.headerFooter = undefined;
		return;
	}
	const next: HeaderFooter = { ...hf };
	for (const key of Object.keys(next) as Array<keyof HeaderFooter>) {
		if (next[key] === '' || next[key] === undefined || next[key] === false) {
			delete next[key];
		}
	}
	if (hf.scaleWithDoc === false) {
		next.scaleWithDoc = false;
	}
	if (hf.alignWithMargins === false) {
		next.alignWithMargins = false;
	}
	sheet.headerFooter = Object.keys(next).length ? next : undefined;
}
