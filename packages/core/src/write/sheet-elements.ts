/**
 * Builders for the simple worksheet children owned by the model. Each takes
 * the original element (when the sheet was loaded) so attributes and
 * children the model does not cover survive.
 */
import { a1 } from '../address/cell-address';
import { MAX_COL_INDEX } from '../address/constants';
import { formatRange, formatSqref } from '../address/range';
import { DEFAULT_COL_WIDTH_CHARS } from '../styles/defaults';
import type { RangeAddress } from '../types/address';
import type {
	ColInfo,
	HeaderFooter,
	PageMargins,
	PageSetup,
	PaneName,
	PrintOptions,
	SheetPane,
	SheetProtection,
	SheetSelection,
	SheetView,
	Worksheet,
} from '../types/sheet';
import { cloneElement, el, textEl } from '../xml';
import type { XmlElement } from '../xml/types';
import { applyFitToPage } from './print-settings';
import {
	colorAttrs,
	findChild,
	findChildren,
	flag,
	mergeAttrs,
	setOrderedChild,
	unflag,
} from './xml-utils';

const SHEET_PR_ORDER = ['tabColor', 'outlinePr', 'pageSetUpPr'] as const;

/** `<sheetPr>`: tab color and outline settings from the model, everything else kept. */
export function buildSheetPr(
	orig: XmlElement | undefined,
	sheet: Worksheet,
	p: string,
): XmlElement | null {
	const out = orig ? cloneElement(orig) : el('sheetPr');
	setOrderedChild(
		out,
		'tabColor',
		sheet.tabColor ? el('tabColor', colorAttrs(sheet.tabColor)) : null,
		SHEET_PR_ORDER,
		p,
	);
	const op = findChild(out, 'outlinePr', p);
	const opAttrs = mergeAttrs(op?.attrs, {
		summaryBelow: unflag(sheet.summaryBelow),
		summaryRight: unflag(sheet.summaryRight),
	});
	setOrderedChild(
		out,
		'outlinePr',
		Object.keys(opAttrs).length ? el('outlinePr', opAttrs) : null,
		SHEET_PR_ORDER,
		p,
	);
	applyFitToPage(out, sheet, !sheet.source, SHEET_PR_ORDER, p);
	return out.children.length || Object.keys(out.attrs).length ? out : null;
}

/** `<dimension>` from the used range. */
export function buildDimension(used: RangeAddress | null): XmlElement {
	return el('dimension', { ref: used ? formatRange(used, { forceCells: true }) : 'A1' });
}

function defaultActivePane(pane: SheetPane): PaneName {
	if (pane.xSplit && pane.ySplit) {
		return 'bottomRight';
	}
	return pane.ySplit ? 'bottomLeft' : 'topRight';
}

function buildPane(pane: SheetPane): XmlElement | null {
	if (!pane.xSplit && !pane.ySplit) {
		return null;
	}
	const frozen = pane.state !== 'split';
	const topLeft =
		pane.topLeftCell ??
		(frozen ? { row: Math.trunc(pane.ySplit), col: Math.trunc(pane.xSplit) } : undefined);
	return el('pane', {
		xSplit: pane.xSplit || undefined,
		ySplit: pane.ySplit || undefined,
		topLeftCell: topLeft ? a1(topLeft.row, topLeft.col) : undefined,
		activePane: pane.activePane ?? defaultActivePane(pane),
		state: pane.state === 'split' ? undefined : pane.state,
	});
}

function buildSelection(sel: SheetSelection, pane: PaneName | undefined): XmlElement {
	const active = a1(sel.activeCell.row, sel.activeCell.col);
	return el('selection', {
		pane: sel.pane ?? pane,
		activeCell: active,
		sqref: sel.ranges.length ? formatSqref(sel.ranges) : active,
	});
}

function buildSheetView(orig: XmlElement | undefined, view: SheetView, p: string): XmlElement {
	const normal = !view.view || view.view === 'normal';
	const zoom =
		view.zoomScale && view.zoomScale !== 100
			? Math.max(10, Math.min(400, Math.round(view.zoomScale)))
			: undefined;
	const grid = view.gridColor?.indexed;
	const attrs = mergeAttrs(orig?.attrs, {
		showFormulas: flag(view.showFormulas),
		showGridLines: unflag(view.showGridLines),
		showRowColHeaders: unflag(view.showRowColHeaders),
		showZeros: unflag(view.showZeros),
		rightToLeft: flag(view.rightToLeft),
		tabSelected: flag(view.tabSelected),
		defaultGridColor: grid !== undefined ? '0' : undefined,
		view: normal ? undefined : view.view,
		topLeftCell: view.topLeftCell ? a1(view.topLeftCell.row, view.topLeftCell.col) : undefined,
		colorId: grid,
		zoomScale: zoom,
		...(normal && orig?.attrs.zoomScaleNormal !== undefined ? { zoomScaleNormal: zoom } : {}),
		workbookViewId: orig?.attrs.workbookViewId ?? '0',
	});
	const pane = view.pane ? buildPane(view.pane) : null;
	const activePane = pane ? (pane.attrs.activePane as PaneName) : undefined;
	const kids: XmlElement[] = [];
	if (pane) {
		kids.push(pane);
	}
	if (view.selection) {
		kids.push(buildSelection(view.selection, activePane));
	}
	// Keep pivot selections and extensions of the original view.
	for (const name of ['pivotSelection', 'extLst']) {
		kids.push(...findChildren(orig, name, p).map(cloneElement));
	}
	return el('sheetView', attrs, kids);
}

/** `<sheetViews>`: the first view from the model, extra views kept. */
export function buildSheetViews(
	orig: XmlElement | undefined,
	view: SheetView,
	p: string,
): XmlElement {
	const out = orig ? cloneElement(orig) : el('sheetViews');
	const first = findChild(out, 'sheetView', p);
	const next = buildSheetView(first, view, p);
	if (first) {
		out.children[out.children.indexOf(first)] = next;
	} else {
		out.children.unshift(next);
	}
	return out;
}

/** `<sheetFormatPr>` (defaults and outline levels). */
export function buildSheetFormatPr(
	orig: XmlElement | undefined,
	sheet: Worksheet,
	outlineRow: number,
	outlineCol: number,
): XmlElement {
	const rowLevel = Math.max(sheet.outlineLevelRow ?? 0, outlineRow);
	const colLevel = Math.max(sheet.outlineLevelCol ?? 0, outlineCol);
	return el(
		'sheetFormatPr',
		mergeAttrs(orig?.attrs, {
			baseColWidth: sheet.baseColWidth,
			defaultColWidth: sheet.defaultColWidth,
			defaultRowHeight: sheet.defaultRowHeight > 0 ? sheet.defaultRowHeight : 15,
			zeroHeight: flag(sheet.zeroHeight),
			outlineLevelRow: rowLevel || undefined,
			outlineLevelCol: colLevel || undefined,
		}),
	);
}

function colKey(c: ColInfo): string {
	return `${c.width}|${c.customWidth}|${c.hidden}|${c.style}|${c.outlineLevel}|${c.collapsed}|${c.bestFit}`;
}

/** Column records: sorted, clipped to the grid, overlap-free, identical neighbours merged. */
export function normalizeCols(cols: readonly ColInfo[]): ColInfo[] {
	const sorted = [...cols]
		.filter((c) => Number.isFinite(c.min) && Number.isFinite(c.max))
		.sort((a, b) => a.min - b.min);
	const out: ColInfo[] = [];
	for (const c of sorted) {
		const prev = out[out.length - 1];
		const min = Math.max(0, c.min, prev ? prev.max + 1 : 0);
		const max = Math.min(MAX_COL_INDEX, c.max);
		if (min > max) {
			continue;
		}
		if (prev && prev.max + 1 === min && colKey(prev) === colKey(c)) {
			prev.max = max;
		} else {
			out.push({ ...c, min, max });
		}
	}
	return out;
}

/** `<cols>` (1-based in the file) plus the highest column outline level. */
export function buildCols(
	sheet: Worksheet,
	styleCount: number,
): { element: XmlElement | null; maxOutline: number } {
	const cols = normalizeCols(sheet.cols);
	let maxOutline = 0;
	const fallbackWidth = sheet.defaultColWidth ?? DEFAULT_COL_WIDTH_CHARS;
	const kids = cols.map((c) => {
		const level = c.outlineLevel ? Math.min(7, Math.trunc(c.outlineLevel)) : 0;
		maxOutline = Math.max(maxOutline, level);
		return el('col', {
			min: c.min + 1,
			max: c.max + 1,
			width: c.width !== undefined && c.width >= 0 ? Math.min(255, c.width) : fallbackWidth,
			style: c.style && c.style > 0 && c.style < styleCount ? c.style : undefined,
			hidden: flag(c.hidden),
			bestFit: flag(c.bestFit),
			customWidth: flag(c.customWidth),
			outlineLevel: level || undefined,
			collapsed: flag(c.collapsed),
		});
	});
	return { element: kids.length ? el('cols', {}, kids) : null, maxOutline };
}

/** `<sheetProtection>`. */
export function buildSheetProtection(prot: SheetProtection | undefined): XmlElement | null {
	if (!prot) {
		return null;
	}
	return el('sheetProtection', {
		password: prot.password,
		algorithmName: prot.algorithmName,
		hashValue: prot.hashValue,
		saltValue: prot.saltValue,
		spinCount: prot.spinCount,
		sheet: prot.sheet ? true : undefined,
		objects: prot.objects,
		scenarios: prot.scenarios,
		formatCells: prot.formatCells,
		formatColumns: prot.formatColumns,
		formatRows: prot.formatRows,
		insertColumns: prot.insertColumns,
		insertRows: prot.insertRows,
		insertHyperlinks: prot.insertHyperlinks,
		deleteColumns: prot.deleteColumns,
		deleteRows: prot.deleteRows,
		selectLockedCells: prot.selectLockedCells,
		sort: prot.sort,
		autoFilter: prot.autoFilter,
		pivotTables: prot.pivotTables,
		selectUnlockedCells: prot.selectUnlockedCells,
	});
}

function intersects(a: RangeAddress, b: RangeAddress): boolean {
	return a.left <= b.right && b.left <= a.right && a.top <= b.bottom && b.top <= a.bottom;
}

/** Merges that Excel accepts: multi-cell, deduplicated, non-overlapping (first one wins). */
export function validMerges(merges: readonly RangeAddress[]): {
	kept: RangeAddress[];
	dropped: number;
} {
	const sorted = merges
		.filter((m) => m.bottom > m.top || m.right > m.left)
		.slice()
		.sort((a, b) => a.top - b.top || a.left - b.left);
	const kept: RangeAddress[] = [];
	let active: RangeAddress[] = [];
	for (const m of sorted) {
		active = active.filter((a) => a.bottom >= m.top);
		if (active.some((a) => intersects(a, m))) {
			continue;
		}
		active.push(m);
		kept.push(m);
	}
	return { kept, dropped: merges.length - kept.length };
}

/** `<mergeCells>`. */
export function buildMergeCells(merges: readonly RangeAddress[]): XmlElement | null {
	if (!merges.length) {
		return null;
	}
	return el(
		'mergeCells',
		{ count: merges.length },
		merges.map((m) => el('mergeCell', { ref: formatRange(m, { forceCells: true }) })),
	);
}

/** `<printOptions>`. */
export function buildPrintOptions(
	orig: XmlElement | undefined,
	po: PrintOptions | undefined,
): XmlElement | null {
	const attrs = mergeAttrs(orig?.attrs, {
		horizontalCentered: flag(po?.horizontalCentered),
		verticalCentered: flag(po?.verticalCentered),
		headings: flag(po?.headings),
		gridLines: flag(po?.gridLines),
	});
	return Object.keys(attrs).length ? el('printOptions', attrs) : null;
}

/** Excel's default margins for a new sheet. */
export const DEFAULT_PAGE_MARGINS: PageMargins = {
	left: 0.7,
	right: 0.7,
	top: 0.75,
	bottom: 0.75,
	header: 0.3,
	footer: 0.3,
};

/** `<pageMargins>` (all six attributes are required). */
export function buildPageMargins(m: PageMargins): XmlElement {
	return el('pageMargins', {
		left: m.left,
		right: m.right,
		top: m.top,
		bottom: m.bottom,
		header: m.header,
		footer: m.footer,
	});
}

/** `<pageSetup>`; the printer-settings relationship and unknown attributes are kept. */
export function buildPageSetup(
	orig: XmlElement | undefined,
	ps: PageSetup | undefined,
): XmlElement | null {
	const attrs = mergeAttrs(orig?.attrs, {
		paperSize: ps?.paperSize,
		scale: ps?.scale,
		firstPageNumber: ps?.firstPageNumber,
		fitToWidth: ps?.fitToWidth,
		fitToHeight: ps?.fitToHeight,
		pageOrder: ps?.pageOrder && ps.pageOrder !== 'downThenOver' ? ps.pageOrder : undefined,
		orientation: ps?.orientation && ps.orientation !== 'default' ? ps.orientation : undefined,
		blackAndWhite: flag(ps?.blackAndWhite),
		draft: flag(ps?.draft),
		useFirstPageNumber: flag(ps?.useFirstPageNumber),
		horizontalDpi: ps?.horizontalDpi,
		verticalDpi: ps?.verticalDpi,
	});
	return Object.keys(attrs).length ? el('pageSetup', attrs) : null;
}

/** `<headerFooter>`. */
export function buildHeaderFooter(
	orig: XmlElement | undefined,
	hf: HeaderFooter | undefined,
): XmlElement | null {
	if (!hf) {
		return null;
	}
	const attrs = mergeAttrs(orig?.attrs, {
		differentOddEven: flag(hf.differentOddEven),
		differentFirst: flag(hf.differentFirst),
		scaleWithDoc: unflag(hf.scaleWithDoc),
		alignWithMargins: unflag(hf.alignWithMargins),
	});
	const kids: XmlElement[] = [];
	const parts: Array<[keyof HeaderFooter, string]> = [
		['oddHeader', 'oddHeader'],
		['oddFooter', 'oddFooter'],
		['evenHeader', 'evenHeader'],
		['evenFooter', 'evenFooter'],
		['firstHeader', 'firstHeader'],
		['firstFooter', 'firstFooter'],
	];
	for (const [key, name] of parts) {
		const text = hf[key];
		if (typeof text === 'string' && text !== '') {
			kids.push(textEl(name, text));
		}
	}
	return kids.length || Object.keys(attrs).length ? el('headerFooter', attrs, kids) : null;
}
