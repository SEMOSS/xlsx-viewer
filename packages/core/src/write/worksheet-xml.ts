/**
 * Worksheet part assembly. Starts from the original root (minus sheetData)
 * when the sheet was loaded, replaces the elements the model owns in
 * CT_Worksheet order and splices in the stream-built `<sheetData>`.
 */
import { formatRange } from '../address/range';
import type { Hyperlink, Worksheet } from '../types/sheet';
import { cloneElement, el } from '../xml';
import { NS } from '../xml/namespaces';
import type { XmlElement } from '../xml/types';
import { buildAutoFilter, buildSortState } from './auto-filter';
import { buildConditionalFormats, CF_EXT_URI, type CfStyles } from './conditional-formats';
import { buildDataValidations, DV_EXT_URI } from './data-validations';
import { mergeExtLst } from './ext-list';
import { buildBreaks } from './print-settings';
import { buildSheetData, type SheetDataContext, type SheetDataResult } from './sheet-data';
import {
	buildCols,
	buildDimension,
	buildHeaderFooter,
	buildMergeCells,
	buildPageMargins,
	buildPageSetup,
	buildPrintOptions,
	buildSheetFormatPr,
	buildSheetPr,
	buildSheetProtection,
	buildSheetViews,
	DEFAULT_PAGE_MARGINS,
	validMerges,
} from './sheet-elements';
import { buildSparklineExt, SPARKLINE_EXT_URI } from './sparklines';
import {
	ensureNamespace,
	findChild,
	normalizeRoot,
	serializeRootWith,
	setOrderedChild,
} from './xml-utils';

/** CT_Worksheet child order. */
export const WORKSHEET_ORDER = [
	'sheetPr',
	'dimension',
	'sheetViews',
	'sheetFormatPr',
	'cols',
	'sheetData',
	'sheetCalcPr',
	'sheetProtection',
	'protectedRanges',
	'scenarios',
	'autoFilter',
	'sortState',
	'dataConsolidate',
	'customSheetViews',
	'mergeCells',
	'phoneticPr',
	'conditionalFormatting',
	'dataValidations',
	'hyperlinks',
	'printOptions',
	'pageMargins',
	'pageSetup',
	'headerFooter',
	'rowBreaks',
	'colBreaks',
	'customProperties',
	'cellWatches',
	'ignoredErrors',
	'smartTags',
	'drawing',
	'legacyDrawing',
	'legacyDrawingHF',
	'drawingHF',
	'picture',
	'oleObjects',
	'controls',
	'webPublishItems',
	'tableParts',
	'extLst',
] as const;

/**
 * Relationship ids of the parts the sheet references. `undefined` keeps the
 * original element untouched, `null` (or an empty list) removes it.
 */
export interface WorksheetLinks {
	drawing?: string | null;
	legacyDrawing?: string | null;
	tableParts?: string[];
	/** Relationship ids of external hyperlinks. */
	hyperlinks: Map<Hyperlink, string>;
}

export interface WorksheetBuildOptions {
	styleCount: number;
	/** Differential formats (inline rule formats are interned through it). */
	styles: CfStyles;
	sheetData: SheetDataContext;
	/** The sheet was created in memory (Excel defaults such as page margins are added). */
	isNew: boolean;
}

export interface WorksheetBuildResult {
	xml: string;
	data: SheetDataResult;
	droppedMerges: number;
	skippedRules: number;
}

function hyperlinksElement(
	links: readonly Hyperlink[],
	relIds: Map<Hyperlink, string>,
	r: string,
): XmlElement | null {
	const kids: XmlElement[] = [];
	for (const h of links) {
		const id = relIds.get(h);
		// Internal targets (`#Sheet2!A1`) are stored as a location, not a relationship.
		const location = h.location ?? (h.target?.startsWith('#') ? h.target.slice(1) : undefined);
		if (!id && !location) {
			continue;
		}
		kids.push(
			el('hyperlink', {
				ref: formatRange(h.ref, { forceCells: true }),
				[`${r}:id`]: id,
				location,
				tooltip: h.tooltip,
				display: h.display,
			}),
		);
	}
	return kids.length ? el('hyperlinks', {}, kids) : null;
}

/** Build the worksheet part XML. */
export function buildWorksheetXml(
	sheet: Worksheet,
	sourceRoot: XmlElement | undefined,
	links: WorksheetLinks,
	opts: WorksheetBuildOptions,
): WorksheetBuildResult {
	const root = sourceRoot
		? cloneElement(sourceRoot)
		: el('worksheet', { xmlns: NS.main, 'xmlns:r': NS.r });
	const p = normalizeRoot(root, 'worksheet', NS.main);
	const r = ensureNamespace(root, 'r', NS.r);
	const set = (name: string, value: XmlElement | XmlElement[] | null): void =>
		setOrderedChild(root, name, value, WORKSHEET_ORDER, p);
	const find = (name: string): XmlElement | undefined => findChild(root, name, p);

	const data = buildSheetData(sheet, opts.sheetData);
	const cols = buildCols(sheet, opts.styleCount);
	const merges = validMerges(sheet.merges);
	const cf = buildConditionalFormats(sheet.conditionalFormats, opts.styles);
	const dv = buildDataValidations(sheet.dataValidations);

	set('sheetPr', buildSheetPr(find('sheetPr'), sheet, p));
	set('dimension', buildDimension(data.used));
	set('sheetViews', buildSheetViews(find('sheetViews'), sheet.view, p));
	set(
		'sheetFormatPr',
		buildSheetFormatPr(find('sheetFormatPr'), sheet, data.maxRowOutline, cols.maxOutline),
	);
	set('cols', cols.element);
	const sheetData = el('sheetData');
	set('sheetData', sheetData);
	set('sheetProtection', buildSheetProtection(sheet.protection));
	set('autoFilter', sheet.autoFilter ? buildAutoFilter(sheet.autoFilter) : null);
	set('sortState', sheet.sortState ? buildSortState(sheet.sortState) : null);
	set('mergeCells', buildMergeCells(merges.kept));
	set('conditionalFormatting', cf.elements);
	set('dataValidations', dv.element);
	set('hyperlinks', hyperlinksElement(sheet.hyperlinks, links.hyperlinks, r));
	set('printOptions', buildPrintOptions(find('printOptions'), sheet.printOptions));
	if (sheet.pageMargins) {
		set('pageMargins', buildPageMargins(sheet.pageMargins));
	} else if (opts.isNew || !find('pageMargins')) {
		set('pageMargins', buildPageMargins(DEFAULT_PAGE_MARGINS));
	}
	set('pageSetup', buildPageSetup(find('pageSetup'), sheet.pageSetup));
	set('headerFooter', buildHeaderFooter(find('headerFooter'), sheet.headerFooter));
	for (const kind of ['rowBreaks', 'colBreaks'] as const) {
		const breaks = buildBreaks(kind, find(kind), sheet[kind], p);
		if (breaks !== undefined) {
			set(kind, breaks);
		}
	}
	if (links.drawing !== undefined) {
		set('drawing', links.drawing ? el('drawing', { [`${r}:id`]: links.drawing }) : null);
	}
	if (links.legacyDrawing !== undefined) {
		set(
			'legacyDrawing',
			links.legacyDrawing ? el('legacyDrawing', { [`${r}:id`]: links.legacyDrawing }) : null,
		);
	}
	if (links.tableParts !== undefined) {
		const ids = links.tableParts;
		set(
			'tableParts',
			ids.length
				? el(
						'tableParts',
						{ count: ids.length },
						ids.map((id) => el('tablePart', { [`${r}:id`]: id })),
					)
				: null,
		);
	}
	set(
		'extLst',
		mergeExtLst(find('extLst'), [
			[CF_EXT_URI, cf.ext],
			[DV_EXT_URI, dv.ext],
			[SPARKLINE_EXT_URI, buildSparklineExt(sheet.sparklineGroups)],
		]),
	);
	return {
		xml: serializeRootWith(root, new Map([[sheetData, data.xml]])),
		data,
		droppedMerges: merges.dropped,
		skippedRules: cf.skipped,
	};
}
