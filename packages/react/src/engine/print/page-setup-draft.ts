import type {
	EditTransaction,
	HeaderFooter,
	PageMargins,
	RangeAddress,
	Workbook,
} from 'xlsx-viewer-core';
import {
	columnToLetters,
	DEFAULT_MARGINS,
	formatRange,
	getPrintArea,
	getPrintTitles,
	paperDimensions,
	paperSize,
	parsePrintAreaText,
	parseTitleCols,
	parseTitleRows,
	setHeaderFooter,
	setMargins,
	setPageSetup,
	setPrintArea,
	setPrintOptions,
	setPrintTitles,
} from 'xlsx-viewer-core';

import { createTranslator, type Translate } from '../../i18n/translator';

/**
 * The editable state of the Page Setup dialog (text fields as typed),
 * read from a sheet, validated with Excel's messages and applied through
 * the page setup operations.
 */

export type MarginKey = keyof PageMargins;
export type HeaderFooterKey =
	| 'oddHeader'
	| 'oddFooter'
	| 'evenHeader'
	| 'evenFooter'
	| 'firstHeader'
	| 'firstFooter';

export const HEADER_FOOTER_KEYS: readonly HeaderFooterKey[] = [
	'oddHeader',
	'oddFooter',
	'evenHeader',
	'evenFooter',
	'firstHeader',
	'firstFooter',
];

export interface PageSetupDraft {
	orientation: 'portrait' | 'landscape';
	scaleMode: 'adjust' | 'fit';
	/** Adjust to (percent). */
	scale: string;
	/** Fit to pages wide / tall ('' = automatic). */
	fitWidth: string;
	fitHeight: string;
	paperSize: number;
	/** 'Auto' (or the UI's word for it) or a number. */
	firstPageNumber: string;
	margins: Record<MarginKey, string>;
	centerH: boolean;
	centerV: boolean;
	texts: Record<HeaderFooterKey, string>;
	differentOddEven: boolean;
	differentFirst: boolean;
	scaleWithDoc: boolean;
	alignWithMargins: boolean;
	printArea: string;
	titleRows: string;
	titleCols: string;
	gridLines: boolean;
	headings: boolean;
	blackAndWhite: boolean;
	draft: boolean;
	pageOrder: 'downThenOver' | 'overThenDown';
}

function inches(v: number): string {
	return String(Math.round(v * 100) / 100);
}

/** The dialog state for a sheet (`t` words the automatic first page number). */
export function draftFromSheet(
	workbook: Workbook,
	sheetIndex: number,
	t: Translate = createTranslator(),
): PageSetupDraft {
	const sheet = workbook.sheets[sheetIndex];
	const ps = sheet.pageSetup ?? {};
	const po = sheet.printOptions ?? {};
	const hf = sheet.headerFooter ?? {};
	const m = sheet.pageMargins ?? DEFAULT_MARGINS;
	const titles = getPrintTitles(workbook, sheetIndex);
	const pages = (n: number | undefined) => (n === 0 ? '' : String(n ?? 1));
	const texts = {} as Record<HeaderFooterKey, string>;
	for (const key of HEADER_FOOTER_KEYS) {
		texts[key] = hf[key] ?? '';
	}
	return {
		orientation: ps.orientation === 'landscape' ? 'landscape' : 'portrait',
		scaleMode: ps.fitToPage ? 'fit' : 'adjust',
		scale: String(ps.scale ?? 100),
		fitWidth: pages(ps.fitToWidth),
		fitHeight: pages(ps.fitToHeight),
		paperSize: paperSize(ps.paperSize).code,
		firstPageNumber:
			ps.useFirstPageNumber && ps.firstPageNumber !== undefined
				? String(ps.firstPageNumber)
				: t('dialog.pageSetup.auto'),
		margins: {
			top: inches(m.top),
			bottom: inches(m.bottom),
			left: inches(m.left),
			right: inches(m.right),
			header: inches(m.header),
			footer: inches(m.footer),
		},
		centerH: po.horizontalCentered === true,
		centerV: po.verticalCentered === true,
		texts,
		differentOddEven: hf.differentOddEven === true,
		differentFirst: hf.differentFirst === true,
		scaleWithDoc: hf.scaleWithDoc !== false,
		alignWithMargins: hf.alignWithMargins !== false,
		printArea: getPrintArea(workbook, sheetIndex)
			.map((a) => formatRange(a, { absolute: true }))
			.join(','),
		titleRows: titles.rows ? `$${titles.rows.start + 1}:$${titles.rows.end + 1}` : '',
		titleCols: titles.cols
			? `$${columnToLetters(titles.cols.start)}:$${columnToLetters(titles.cols.end)}`
			: '',
		gridLines: po.gridLines === true,
		headings: po.headings === true,
		blackAndWhite: ps.blackAndWhite === true,
		draft: ps.draft === true,
		pageOrder: ps.pageOrder === 'overThenDown' ? 'overThenDown' : 'downThenOver',
	};
}

function wholeNumber(text: string, min: number, max: number): number | null {
	const t = text.trim();
	if (!/^\d+$/.test(t)) {
		return null;
	}
	const n = Number(t);
	return n >= min && n <= max ? n : null;
}

function marginValues(d: PageSetupDraft): PageMargins | null {
	const out = {} as PageMargins;
	for (const key of Object.keys(d.margins) as MarginKey[]) {
		const n = Number(d.margins[key].trim());
		if (d.margins[key].trim() === '' || !Number.isFinite(n) || n < 0 || n > 100) {
			return null;
		}
		out[key] = n;
	}
	return out;
}

/** A whole first page number (anything else is automatic once validated). */
const PAGE_NUMBER = /^-?\d+$/;

/** The First page number box asks for automatic numbering: `Auto` or the UI's word for it. */
function isAutoPageNumber(text: string, t: Translate): boolean {
	const v = text.trim().toLowerCase();
	return v === 'auto' || v === t('dialog.pageSetup.auto').toLowerCase();
}

/** Excel's message for the first invalid field (in the language of `t`), or null. */
export function validateDraft(
	d: PageSetupDraft,
	sheetName: string,
	t: Translate = createTranslator(),
): string | null {
	if (d.scaleMode === 'adjust' && wholeNumber(d.scale, 10, 400) === null) {
		return t('dialog.pageSetup.invalidScale');
	}
	if (d.scaleMode === 'fit') {
		for (const v of [d.fitWidth, d.fitHeight]) {
			if (v.trim() !== '' && wholeNumber(v, 1, 32767) === null) {
				return t('dialog.pageSetup.invalidPages');
			}
		}
	}
	if (!isAutoPageNumber(d.firstPageNumber, t) && !PAGE_NUMBER.test(d.firstPageNumber.trim())) {
		return t('dialog.pageSetup.invalidFirstPage', { auto: t('dialog.pageSetup.auto') });
	}
	const m = marginValues(d);
	if (!m) {
		return t('dialog.pageSetup.invalidMargins');
	}
	const paper = paperDimensions({ paperSize: d.paperSize, orientation: d.orientation });
	if ((m.left + m.right) * 72 >= paper.width || (m.top + m.bottom) * 72 >= paper.height) {
		return t('dialog.pageSetup.marginsTooLarge');
	}
	if (parsePrintAreaText(d.printArea, sheetName) === null) {
		return t('message.invalidReference');
	}
	if (parseTitleRows(d.titleRows) === null || parseTitleCols(d.titleCols) === null) {
		return t('message.invalidReference');
	}
	return null;
}

function sameAreas(a: readonly RangeAddress[], b: readonly RangeAddress[]): boolean {
	return JSON.stringify(a) === JSON.stringify(b);
}

/** Apply a validated draft to the sheet (call inside one transaction). */
export function applyDraft(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	d: PageSetupDraft,
): void {
	const sheet = workbook.sheets[sheetIndex];
	const prev = sheet.pageSetup ?? {};
	const fit = d.scaleMode === 'fit';
	const pages = (v: string) => (v.trim() === '' ? 0 : Number(v.trim()));
	const auto = !PAGE_NUMBER.test(d.firstPageNumber.trim());
	setPageSetup(tx, workbook, sheetIndex, {
		orientation: d.orientation,
		paperSize: d.paperSize,
		scale: fit ? prev.scale : Number(d.scale.trim()),
		fitToPage: fit ? true : prev.fitToPage !== undefined ? false : undefined,
		fitToWidth: fit ? pages(d.fitWidth) : prev.fitToWidth,
		fitToHeight: fit ? pages(d.fitHeight) : prev.fitToHeight,
		firstPageNumber: auto ? prev.firstPageNumber : Number(d.firstPageNumber.trim()),
		useFirstPageNumber: auto ? undefined : true,
		pageOrder: d.pageOrder === 'overThenDown' ? 'overThenDown' : undefined,
		blackAndWhite: d.blackAndWhite || undefined,
		draft: d.draft || undefined,
	});
	const m = marginValues(d);
	if (m) {
		setMargins(tx, workbook, sheetIndex, m);
	}
	setPrintOptions(tx, workbook, sheetIndex, {
		gridLines: d.gridLines || undefined,
		headings: d.headings || undefined,
		horizontalCentered: d.centerH || undefined,
		verticalCentered: d.centerV || undefined,
	});
	const hf: HeaderFooter = {
		differentOddEven: d.differentOddEven,
		differentFirst: d.differentFirst,
		scaleWithDoc: d.scaleWithDoc ? undefined : false,
		alignWithMargins: d.alignWithMargins ? undefined : false,
	};
	for (const key of HEADER_FOOTER_KEYS) {
		const evenKey = key.startsWith('even');
		const firstKey = key.startsWith('first');
		if ((evenKey && !d.differentOddEven) || (firstKey && !d.differentFirst)) {
			continue;
		}
		hf[key] = d.texts[key];
	}
	setHeaderFooter(tx, workbook, sheetIndex, hf);
	const areas = parsePrintAreaText(d.printArea, sheet.name) ?? [];
	if (!sameAreas(areas, getPrintArea(workbook, sheetIndex))) {
		setPrintArea(tx, workbook, sheetIndex, areas);
	}
	const rows = parseTitleRows(d.titleRows) ?? undefined;
	const cols = parseTitleCols(d.titleCols) ?? undefined;
	const titles = getPrintTitles(workbook, sheetIndex);
	if (JSON.stringify({ rows, cols }) !== JSON.stringify({ rows: titles.rows, cols: titles.cols })) {
		setPrintTitles(tx, workbook, sheetIndex, { rows, cols });
	}
}
