import { createWorksheet } from '../sheet/worksheet';
import type { SheetKind, Worksheet } from '../types/sheet';
import { findDescendantLocal } from '../xml/dom';
import type { XmlElement } from '../xml/types';
import { parseAutoFilter, parseSortState } from './auto-filter';
import {
	findWorksheetExtensions,
	mergeCfExtensions,
	type WorksheetExtensions,
} from './cf-extensions';
import { readChartsheet } from './chartsheet';
import { parseConditionalFormats } from './conditional-formatting';
import { attemptSync, type ReadContext } from './context';
import { parseDataValidations } from './data-validation';
import { readPartText, relLookup, relTypeSuffix, type RelLookup } from './part-io';
import { parseSheetData } from './sheet-data';
import {
	applyPrintSettings,
	applySheetFormat,
	applySheetPr,
	parseCols,
	parseHyperlinks,
	parseMerges,
	parseSheetProtection,
} from './sheet-features';
import { readSheetParts } from './sheet-parts';
import { parseSheetView } from './sheet-view';
import { parseSparklineGroups } from './sparklines';
import type { SheetEntry } from './workbook-part';
import { parseSpreadsheetXml, scanRootTag, spreadsheetPrefix, tagEnd } from './xml-normalize';
import { kid } from './xml-utils';

/** A loaded sheet before its part is read: model defaults plus clean change-tracking flags. */
export function createSheetShell(entry: SheetEntry, kind: SheetKind): Worksheet {
	const sheet = createWorksheet(entry.name, entry.sheetId);
	sheet.kind = kind;
	sheet.state = entry.state;
	sheet.dirty = false;
	sheet.drawingsDirty = false;
	sheet.commentsDirty = false;
	return sheet;
}

/** Content bounds of `<sheetData>` (`undefined` when absent or empty). */
export function locateSheetData(
	xml: string,
	from: number,
	prefix: string,
): { contentStart: number; contentEnd: number } | undefined {
	const p = prefix ? `${prefix}:` : '';
	const open = `<${p}sheetData`;
	let i = xml.indexOf(open, from);
	while (i !== -1) {
		const c = xml.charCodeAt(i + open.length);
		if (c === 62 || c === 47 || c === 32 || c === 9 || c === 10 || c === 13) {
			break;
		}
		i = xml.indexOf(open, i + 1);
	}
	if (i === -1) {
		return undefined;
	}
	const gt = tagEnd(xml, i);
	if (gt === -1 || xml.charCodeAt(gt - 2) === 47) {
		return undefined;
	}
	const close = xml.lastIndexOf(`</${p}sheetData>`);
	return close >= gt ? { contentStart: gt, contentEnd: close } : undefined;
}

function featureWarnings(
	ctx: ReadContext,
	sheet: Worksheet,
	root: XmlElement,
	rels: RelLookup,
	ext: WorksheetExtensions,
): void {
	const at = { part: sheet.source?.partPath, sheet: sheet.name };
	const relTypes = new Set(rels.rels.map((r) => relTypeSuffix(r.type)));
	if (relTypes.has('pivotTable')) {
		ctx.warn('pivot-table', `Pivot tables on "${sheet.name}" are shown as static cells`, at);
	}
	if (relTypes.has('slicer') || ext.names.has('slicerList')) {
		ctx.warn(
			'slicer',
			`Slicers on "${sheet.name}" are shown as placeholders and cannot be used`,
			at,
		);
	}
	if (relTypes.has('timeline') || ext.names.has('timelineRefs')) {
		ctx.warn('slicer', `Timelines on "${sheet.name}" are shown as placeholders`, at);
	}
	if (relTypes.has('queryTable')) {
		ctx.warn('external-data', `"${sheet.name}" contains query tables; data is not refreshed`, at);
	}
	if (findDescendantLocal(root, 'oleObjects')) {
		ctx.warn(
			'ole-object',
			`"${sheet.name}" contains embedded OLE objects, which are not rendered`,
			at,
		);
	}
	if (findDescendantLocal(root, 'controls')) {
		ctx.warn(
			'form-control',
			`"${sheet.name}" contains form or ActiveX controls, which are not interactive`,
			at,
		);
	}
}

function applyFeatures(
	ctx: ReadContext,
	sheet: Worksheet,
	root: XmlElement,
	rels: RelLookup,
): WorksheetExtensions {
	const part = sheet.source!.partPath;
	applySheetPr(sheet, kid(root, 'sheetPr'));
	sheet.view = parseSheetView(kid(root, 'sheetViews'));
	applySheetFormat(sheet, kid(root, 'sheetFormatPr'));
	sheet.cols = parseCols(root);
	sheet.merges = parseMerges(root);
	const protection = parseSheetProtection(kid(root, 'sheetProtection'));
	if (protection) {
		sheet.protection = protection;
	}
	const autoFilter = parseAutoFilter(kid(root, 'autoFilter'));
	if (autoFilter) {
		sheet.autoFilter = autoFilter;
	}
	const sortState = parseSortState(kid(root, 'sortState'));
	if (sortState) {
		sheet.sortState = sortState;
	}
	sheet.hyperlinks = parseHyperlinks(root, rels);
	applyPrintSettings(sheet, root);
	const ext = findWorksheetExtensions(root);
	sheet.conditionalFormats = attemptSync(
		ctx,
		part,
		'conditional formatting in',
		() => {
			const formats = parseConditionalFormats(root);
			if (ext.conditionalFormattings) {
				mergeCfExtensions(formats, ext.conditionalFormattings);
			}
			return formats;
		},
		[],
		sheet.name,
	);
	sheet.dataValidations = attemptSync(
		ctx,
		part,
		'data validation in',
		() => parseDataValidations(root, ext.dataValidations),
		[],
		sheet.name,
	);
	sheet.sparklineGroups = attemptSync(
		ctx,
		part,
		'sparklines in',
		() => parseSparklineGroups(ext.sparklineGroups),
		[],
		sheet.name,
	);
	return ext;
}

function readCells(
	ctx: ReadContext,
	sheet: Worksheet,
	xml: string,
	start: number,
	end: number,
	prefix: string,
): void {
	const part = sheet.source!.partPath;
	const data = parseSheetData(xml, start, end, {
		prefix,
		sharedStrings: ctx.sharedStrings,
		date1904: ctx.date1904,
		dynamicArrayCm: ctx.dynamicArrayCm,
	});
	sheet.cells = data.cells;
	sheet.rows = data.rows;
	for (const issue of data.issues) {
		ctx.warn(issue.code, `${sheet.name}: ${issue.message}`, {
			part,
			sheet: sheet.name,
			cell: issue.cell,
		});
	}
	if (data.dataTableCells > 0) {
		ctx.warn(
			'data-table',
			`"${sheet.name}" contains a what-if data table; its cached results are shown but not recalculated`,
			{ part, sheet: sheet.name },
		);
	}
}

/**
 * Read a sheet part into `sheet`. Cells come from the fast `<sheetData>`
 * scanner; everything else is parsed as XML with the sheetData content cut
 * out, and that tree becomes `sheet.source.root`.
 */
export async function readSheetPart(
	ctx: ReadContext,
	sheet: Worksheet,
	partPath: string,
): Promise<void> {
	const xml = await readPartText(ctx.pkg, partPath);
	if (xml === undefined) {
		ctx.warn('missing-part', `The part of sheet "${sheet.name}" (${partPath}) is missing`, {
			part: partPath,
			sheet: sheet.name,
		});
		return;
	}
	const rels = relLookup(partPath, await ctx.pkg.rels(partPath), ctx.pkg);
	// The model gets its own copies: the package caches the parsed relationships.
	sheet.source = { partPath, rels: rels.rels.map((r) => ({ ...r })) };
	const rootTag = scanRootTag(xml);
	if (!rootTag) {
		throw new Error('the part has no root element');
	}
	const prefix = spreadsheetPrefix(rootTag.attrs);
	let rest = xml;
	if (sheet.kind !== 'chartsheet') {
		const bounds = locateSheetData(xml, rootTag.end, prefix);
		if (bounds) {
			attemptSync(
				ctx,
				partPath,
				'cells of',
				() => readCells(ctx, sheet, xml, bounds.contentStart, bounds.contentEnd, prefix),
				undefined,
				sheet.name,
			);
			rest = xml.slice(0, bounds.contentStart) + xml.slice(bounds.contentEnd);
		}
	}
	const root = parseSpreadsheetXml(rest).root;
	sheet.source.root = root;
	if (sheet.kind === 'chartsheet') {
		await readChartsheet(ctx, sheet, root, rels);
		return;
	}
	const ext = applyFeatures(ctx, sheet, root, rels);
	featureWarnings(ctx, sheet, root, rels, ext);
	await readSheetParts(ctx, sheet, root, rels);
}
