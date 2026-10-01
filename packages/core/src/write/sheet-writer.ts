/**
 * Saving one worksheet. A loaded sheet with `dirty === false` keeps its part
 * and relationships unless a linked part had to be added or dropped (then
 * it is regenerated so the references stay consistent). Everything else is
 * regenerated from the model on top of the original root.
 */
import { relativeTarget, resolveRelTarget } from '../package';
import type { Relationship, Worksheet } from '../types/sheet';
import type { Workbook } from '../types/workbook';
import { isElement, localName, parseXml } from '../xml';
import { CT, REL } from '../xml/namespaces';
import type { XmlElement } from '../xml/types';
import { addRel, sameRels } from './context';
import type { CellMetadata } from './metadata';
import type { SharedStringTable } from './shared-strings';
import { type LinkState, planSheetLinks } from './sheet-links';
import { activeTab } from './workbook-xml';
import { buildWorksheetXml } from './worksheet-xml';
import { findChild } from './xml-utils';

export interface SheetTarget {
	sheet: Worksheet;
	/** Output part path. */
	partPath: string;
	/** Relationship id in the workbook rels. */
	relId: string;
	/** The part is the sheet's own source part. */
	hasSource: boolean;
	/** The sheet is a copy of another sheet's source (fresh parts, shared root markup). */
	fresh: boolean;
}

export interface SheetWriteState extends LinkState {
	sst(): Promise<SharedStringTable>;
	metadata: CellMetadata;
}

export interface SheetWriteResult {
	regenerated: boolean;
	/** Stale or never-calculated formulas were written (Excel should recalculate on load). */
	needsCalc: boolean;
}

const MANAGED = new Set(['drawing', 'comments', 'threadedComment', 'table', 'hyperlink']);

function suffix(type: string): string {
	return type.slice(type.lastIndexOf('/') + 1);
}

function rootRelId(root: XmlElement | undefined, name: string): string | undefined {
	const el = findChild(root, name);
	return el
		? (el.attrs['r:id'] ?? Object.entries(el.attrs).find(([k]) => k.endsWith(':id'))?.[1])
		: undefined;
}

/** Relationships a sheet copy inherits: everything except the parts it gets its own copy of. */
function inheritedRels(sheet: Worksheet, partPath: string): Relationship[] {
	const src = sheet.source;
	if (!src) {
		return [];
	}
	const legacy = rootRelId(src.root, 'legacyDrawing');
	return src.rels
		.filter(
			(r) => !MANAGED.has(suffix(r.type)) && !(suffix(r.type) === 'vmlDrawing' && r.id === legacy),
		)
		.map((r) =>
			r.targetMode === 'External'
				? { ...r }
				: { ...r, target: relativeTarget(partPath, resolveRelTarget(src.partPath, r.target)) },
		);
}

/** The original root (sheetData emptied): the reader's copy, else parsed from the part. */
async function sourceRoot(
	state: SheetWriteState,
	target: SheetTarget,
): Promise<XmlElement | undefined> {
	const src = target.sheet.source;
	if (src?.root) {
		return src.root;
	}
	if (!target.hasSource) {
		return undefined;
	}
	const xml = await state.ctx.sourceText(target.partPath);
	if (!xml) {
		return undefined;
	}
	const root = parseXml(xml).root;
	for (const c of root.children) {
		if (isElement(c) && localName(c.name) === 'sheetData') {
			c.children = [];
		}
	}
	return root;
}

function hasStale(sheet: Worksheet): boolean {
	for (const c of sheet.cells.values()) {
		if (c.stale) {
			return true;
		}
	}
	return false;
}

function uncalculated(sheet: Worksheet): boolean {
	for (const c of sheet.cells.values()) {
		if (c.formula && c.arrayAnchor === undefined && (c.value === null || c.value === undefined)) {
			return true;
		}
	}
	return false;
}

/**
 * Excel expects the active sheet's view to be selected; when no sheet of
 * the model is marked, the active tab is.
 */
function selectedView(workbook: Workbook, sheet: Worksheet): Worksheet['view'] {
	if (sheet.view.tabSelected || workbook.sheets.some((s) => s.view.tabSelected)) {
		return sheet.view;
	}
	const active = workbook.sheets[activeTab(workbook)];
	return active === sheet ? { ...sheet.view, tabSelected: true } : sheet.view;
}

/** Write one worksheet (and the parts it links to). */
export async function writeWorksheet(
	state: SheetWriteState,
	target: SheetTarget,
): Promise<SheetWriteResult> {
	const { ctx } = state;
	const { sheet, partPath } = target;
	const clean = target.hasSource && sheet.dirty === false && !ctx.options.regenerateAll;
	const srcRels = target.hasSource
		? await ctx.rels(partPath)
		: target.fresh
			? inheritedRels(sheet, partPath)
			: [];
	const rels = srcRels.map((r) => ({ ...r }));
	let root = clean ? sheet.source?.root : await sourceRoot(state, target);
	const links = await planSheetLinks(
		state,
		{ sheet, partPath, hasSource: target.hasSource, fresh: target.fresh, root, rels },
		clean,
	);
	if (clean && sameRels(srcRels, rels)) {
		return { regenerated: false, needsCalc: hasStale(sheet) };
	}
	root ??= await sourceRoot(state, target);
	for (const r of [...rels]) {
		if (suffix(r.type) === 'hyperlink') {
			rels.splice(rels.indexOf(r), 1);
		}
	}
	const hyperlinks = new Map<Worksheet['hyperlinks'][number], string>();
	for (const h of sheet.hyperlinks) {
		if (h.target && !h.target.startsWith('#')) {
			hyperlinks.set(h, addRel(rels, REL.hyperlink, h.target, 'External'));
		}
	}
	const styles = ctx.styles;
	const sst = await state.sst();
	const view = selectedView(ctx.workbook, sheet);
	const model = view === sheet.view ? sheet : { ...sheet, view };
	const result = buildWorksheetXml(
		model,
		root,
		{ ...links, hyperlinks },
		{
			styleCount: styles.cellXfs.length,
			styles: { dxfs: () => ctx.styles.dxfs, internDxf: (d) => ctx.internDxf(d) },
			sheetData: { sst, styleCount: styles.cellXfs.length, metadata: state.metadata },
			isNew: !target.hasSource && !target.fresh,
		},
	);
	ctx.writeText(partPath, result.xml, CT.worksheet);
	await ctx.writeRelsIfChanged(partPath, rels);
	const at = { sheet: sheet.name, part: partPath };
	if (result.droppedMerges) {
		ctx.warn(
			'merge-dropped',
			`${result.droppedMerges} invalid or overlapping merged range(s) were not saved`,
			at,
		);
	}
	if (result.skippedRules) {
		ctx.warn(
			'cf-rule-dropped',
			`${result.skippedRules} incomplete conditional formatting rule(s) were not saved`,
			at,
		);
	}
	if (result.data.invalidStyles) {
		ctx.warn(
			'invalid-style',
			`${result.data.invalidStyles} cell(s) referenced a missing style`,
			at,
		);
	}
	if (result.data.invalidNumbers) {
		ctx.warn(
			'invalid-number',
			`${result.data.invalidNumbers} non-finite number(s) were saved as #NUM!`,
			at,
		);
	}
	return { regenerated: true, needsCalc: result.data.stale || uncalculated(sheet) };
}
