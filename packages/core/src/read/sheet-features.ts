import { MAX_COL_INDEX } from '../address/constants';
import { parseRange } from '../address/range';
import type {
	ColInfo,
	HeaderFooter,
	Hyperlink,
	PageMargins,
	PageSetup,
	PrintOptions,
	SheetProtection,
	Worksheet,
} from '../types/sheet';
import type { XmlElement } from '../xml/types';
import type { RelLookup } from './part-io';
import { parseColor } from './style-parts';
import { bool, compact, int, kid, kids, num, oneOf, ooxmlText, str } from './xml-utils';

/** Readers for the simple worksheet-level elements (everything but cells and rules). */

/** `sheetPr`: tab color and outline summary positions. */
export function applySheetPr(sheet: Worksheet, sheetPr: XmlElement | undefined): void {
	if (!sheetPr) {
		return;
	}
	const tab = parseColor(kid(sheetPr, 'tabColor'));
	if (tab) {
		sheet.tabColor = tab;
	}
	const outline = kid(sheetPr, 'outlinePr');
	const below = bool(outline, 'summaryBelow');
	if (below !== undefined) {
		sheet.summaryBelow = below;
	}
	const right = bool(outline, 'summaryRight');
	if (right !== undefined) {
		sheet.summaryRight = right;
	}
}

/** `sheetFormatPr`: default row height / column width and outline levels. */
export function applySheetFormat(sheet: Worksheet, fmt: XmlElement | undefined): void {
	if (!fmt) {
		return;
	}
	const height = num(fmt, 'defaultRowHeight');
	if (height !== undefined && height > 0) {
		sheet.defaultRowHeight = height;
	}
	const width = num(fmt, 'defaultColWidth');
	if (width !== undefined) {
		sheet.defaultColWidth = width;
	}
	const base = num(fmt, 'baseColWidth');
	if (base !== undefined) {
		sheet.baseColWidth = base;
	}
	if (bool(fmt, 'zeroHeight')) {
		sheet.zeroHeight = true;
	}
	const rowLevel = int(fmt, 'outlineLevelRow');
	if (rowLevel) {
		sheet.outlineLevelRow = rowLevel;
	}
	const colLevel = int(fmt, 'outlineLevelCol');
	if (colLevel) {
		sheet.outlineLevelCol = colLevel;
	}
}

/** All `<cols>/<col>` records, converted to 0-based and sorted. */
export function parseCols(root: XmlElement): ColInfo[] {
	const out: ColInfo[] = [];
	for (const cols of kids(root, 'cols')) {
		for (const c of kids(cols, 'col')) {
			const min = int(c, 'min');
			const max = int(c, 'max') ?? min;
			if (min === undefined || max === undefined || min < 1 || max < min) {
				continue;
			}
			const level = int(c, 'outlineLevel');
			out.push(
				compact<ColInfo>({
					min: min - 1,
					max: Math.min(max - 1, MAX_COL_INDEX),
					width: num(c, 'width'),
					customWidth: bool(c, 'customWidth') || undefined,
					hidden: bool(c, 'hidden') || undefined,
					style: int(c, 'style'),
					outlineLevel: level ? level : undefined,
					collapsed: bool(c, 'collapsed') || undefined,
					bestFit: bool(c, 'bestFit') || undefined,
				}),
			);
		}
	}
	out.sort((a, b) => a.min - b.min);
	// Keep the records non-overlapping: a later record yields to an earlier one.
	const clean: ColInfo[] = [];
	for (const c of out) {
		const prev = clean[clean.length - 1];
		if (prev && c.min <= prev.max) {
			if (c.max <= prev.max) {
				continue;
			}
			c.min = prev.max + 1;
		}
		clean.push(c);
	}
	return clean;
}

/** `<mergeCells>/<mergeCell ref>`. */
export function parseMerges(root: XmlElement): Worksheet['merges'] {
	const out: Worksheet['merges'] = [];
	for (const m of kids(kid(root, 'mergeCells'), 'mergeCell')) {
		const r = parseRange(str(m, 'ref') ?? '');
		if (r && (r.top !== r.bottom || r.left !== r.right)) {
			out.push(r);
		}
	}
	return out;
}

const PROTECTION_FLAGS = [
	'objects',
	'scenarios',
	'formatCells',
	'formatColumns',
	'formatRows',
	'insertColumns',
	'insertRows',
	'insertHyperlinks',
	'deleteColumns',
	'deleteRows',
	'selectLockedCells',
	'sort',
	'autoFilter',
	'pivotTables',
	'selectUnlockedCells',
] as const;

/**
 * `sheetProtection`. Attributes map 1:1 onto the model with OOXML meaning
 * (`true` = the action is locked while the sheet is protected); attributes
 * absent from the file are left `undefined` so the schema defaults apply.
 */
export function parseSheetProtection(el: XmlElement | undefined): SheetProtection | undefined {
	if (!el) {
		return undefined;
	}
	const p: SheetProtection = compact({
		password: str(el, 'password'),
		algorithmName: str(el, 'algorithmName'),
		hashValue: str(el, 'hashValue'),
		saltValue: str(el, 'saltValue'),
		spinCount: int(el, 'spinCount'),
		sheet: bool(el, 'sheet') ?? false,
	});
	for (const key of PROTECTION_FLAGS) {
		const v = bool(el, key);
		if (v !== undefined) {
			p[key] = v;
		}
	}
	return p;
}

/** `<hyperlinks>`: external targets come from the sheet relationships. */
export function parseHyperlinks(root: XmlElement, rels: RelLookup): Hyperlink[] {
	const out: Hyperlink[] = [];
	for (const h of kids(kid(root, 'hyperlinks'), 'hyperlink')) {
		const ref = parseRange(str(h, 'ref') ?? '');
		if (!ref) {
			continue;
		}
		const link: Hyperlink = { ref };
		const rel = rels.byId(str(h, 'r:id'));
		if (rel) {
			link.target = rel.target;
		}
		const location = str(h, 'location');
		if (location) {
			link.location = location;
		}
		const display = str(h, 'display');
		if (display !== undefined) {
			link.display = display;
		}
		const tooltip = str(h, 'tooltip');
		if (tooltip !== undefined) {
			link.tooltip = tooltip;
		}
		out.push(link);
	}
	return out;
}

/**
 * Manual breaks of `<rowBreaks>` / `<colBreaks>`: the `id` of each `<brk>`
 * is the 0-based index of the first row (column) after the break, so it is
 * the model value as is. Automatic breaks (`man` absent or false) are left
 * to the writer, which keeps them from the source element.
 */
function parseBreaks(el: XmlElement | undefined): number[] | undefined {
	if (!el) {
		return undefined;
	}
	const ids = new Set<number>();
	for (const b of kids(el, 'brk')) {
		const id = int(b, 'id');
		if (id !== undefined && id > 0 && bool(b, 'man') === true) {
			ids.add(id);
		}
	}
	return [...ids].sort((a, b) => a - b);
}

/**
 * Print-related elements: `printOptions`, `pageMargins`, `pageSetup` (plus
 * `sheetPr/pageSetUpPr/@fitToPage`), `headerFooter`, `rowBreaks`, `colBreaks`.
 */
export function applyPrintSettings(sheet: Worksheet, root: XmlElement): void {
	const po = kid(root, 'printOptions');
	if (po) {
		const options = compact<PrintOptions>({
			gridLines: bool(po, 'gridLines'),
			headings: bool(po, 'headings'),
			horizontalCentered: bool(po, 'horizontalCentered'),
			verticalCentered: bool(po, 'verticalCentered'),
		});
		sheet.printOptions = options;
	}
	const pm = kid(root, 'pageMargins');
	if (pm) {
		const margins: PageMargins = {
			left: num(pm, 'left') ?? 0.7,
			right: num(pm, 'right') ?? 0.7,
			top: num(pm, 'top') ?? 0.75,
			bottom: num(pm, 'bottom') ?? 0.75,
			header: num(pm, 'header') ?? 0.3,
			footer: num(pm, 'footer') ?? 0.3,
		};
		sheet.pageMargins = margins;
	}
	const ps = kid(root, 'pageSetup');
	if (ps) {
		sheet.pageSetup = compact<PageSetup>({
			paperSize: int(ps, 'paperSize'),
			orientation: oneOf(str(ps, 'orientation'), ['default', 'portrait', 'landscape'] as const),
			scale: int(ps, 'scale'),
			fitToWidth: int(ps, 'fitToWidth'),
			fitToHeight: int(ps, 'fitToHeight'),
			firstPageNumber: int(ps, 'firstPageNumber'),
			useFirstPageNumber: bool(ps, 'useFirstPageNumber'),
			pageOrder: oneOf(str(ps, 'pageOrder'), ['downThenOver', 'overThenDown'] as const),
			blackAndWhite: bool(ps, 'blackAndWhite'),
			draft: bool(ps, 'draft'),
			horizontalDpi: int(ps, 'horizontalDpi'),
			verticalDpi: int(ps, 'verticalDpi'),
		});
	}
	const fitToPage = bool(kid(kid(root, 'sheetPr'), 'pageSetUpPr'), 'fitToPage');
	if (fitToPage !== undefined) {
		sheet.pageSetup = { ...sheet.pageSetup, fitToPage };
	}
	const rowBreaks = parseBreaks(kid(root, 'rowBreaks'));
	if (rowBreaks) {
		sheet.rowBreaks = rowBreaks;
	}
	const colBreaks = parseBreaks(kid(root, 'colBreaks'));
	if (colBreaks) {
		sheet.colBreaks = colBreaks;
	}
	const hf = kid(root, 'headerFooter');
	if (hf) {
		const headerFooter: HeaderFooter = compact({
			differentOddEven: bool(hf, 'differentOddEven'),
			differentFirst: bool(hf, 'differentFirst'),
			scaleWithDoc: bool(hf, 'scaleWithDoc'),
			alignWithMargins: bool(hf, 'alignWithMargins'),
		});
		for (const key of [
			'oddHeader',
			'oddFooter',
			'evenHeader',
			'evenFooter',
			'firstHeader',
			'firstFooter',
		] as const) {
			const el = kid(hf, key);
			if (el) {
				headerFooter[key] = ooxmlText(el);
			}
		}
		sheet.headerFooter = headerFooter;
	}
}
