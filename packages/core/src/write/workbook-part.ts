/**
 * The workbook part and its relationships: locating the original, giving
 * every sheet a part (new sheets get fresh `sheetN.xml` names), validating
 * sheet names / ids, writing `workbook.xml` and dropping the calc chain.
 */
import { isValidSheetName } from '../address/sheet-name';
import { findRelByType, relativeTarget, resolveRelTarget } from '../package';
import type { Relationship } from '../types/sheet';
import { parseXml, serializeElement, serializeXml } from '../xml';
import { CT, REL } from '../xml/namespaces';
import type { XmlElement } from '../xml/types';
import { addRel, type WriteContext } from './context';
import type { SheetTarget } from './sheet-writer';
import { buildWorkbookElement, type WorkbookSheetEntry } from './workbook-xml';
import { findChild, findChildren } from './xml-utils';

export interface OriginalSheet {
	name: string;
	sheetId: number;
	relId: string;
	partPath?: string;
	worksheet: boolean;
}

export interface WorkbookPart {
	path: string;
	root?: XmlElement;
	/** Working copy of the workbook relationships (updated during the save). */
	rels: Relationship[];
	originalSheets: OriginalSheet[];
}

const SHEET_TYPES = new Set([
	'worksheet',
	'chartsheet',
	'dialogsheet',
	'xlMacrosheet',
	'xlIntlMacrosheet',
]);

export function relSuffix(type: string): string {
	return type.slice(type.lastIndexOf('/') + 1);
}

function relIdAttr(el: XmlElement): string | undefined {
	return el.attrs['r:id'] ?? Object.entries(el.attrs).find(([k]) => k.endsWith(':id'))?.[1];
}

/** Locate and parse the source workbook part (or describe a new one). */
export async function loadWorkbookPart(ctx: WriteContext): Promise<WorkbookPart> {
	if (!ctx.source) {
		return { path: 'xl/workbook.xml', rels: [], originalSheets: [] };
	}
	const main = findRelByType(await ctx.rels(''), REL.officeDocument);
	const path = main ? resolveRelTarget('', main.target) : 'xl/workbook.xml';
	const xml = await ctx.sourceText(path);
	const root = xml ? parseXml(xml).root : undefined;
	const rels = (await ctx.rels(path)).map((r) => ({ ...r }));
	const byId = new Map(rels.map((r) => [r.id, r]));
	const originalSheets: OriginalSheet[] = findChildren(
		findChild(root, 'sheets', ''),
		'sheet',
		'',
	).map((s) => {
		const relId = relIdAttr(s) ?? '';
		const rel = byId.get(relId);
		return {
			name: s.attrs.name ?? '',
			sheetId: Number(s.attrs.sheetId),
			relId,
			partPath: rel ? resolveRelTarget(path, rel.target) : undefined,
			worksheet: !!rel && relSuffix(rel.type) === 'worksheet',
		};
	});
	return { path, root, rels, originalSheets };
}

function sanitizeName(name: string): string {
	let out = (name ?? '').replace(/[\\/?*[\]:]/g, '_').replace(/^'+|'+$/g, '');
	if (out.length > 31) {
		out = out.slice(0, 31);
	}
	return out || 'Sheet';
}

/** Sheet names and ids as written: valid, case-insensitively unique names and unique ids. */
export function sheetIdentities(ctx: WriteContext): Array<{ name: string; sheetId: number }> {
	const sheets = ctx.workbook.sheets;
	const names = new Set<string>();
	const ids = new Set<number>();
	let maxId = sheets.reduce(
		(m, s) => (Number.isInteger(s.sheetId) && s.sheetId > 0 ? Math.max(m, s.sheetId) : m),
		0,
	);
	return sheets.map((s) => {
		let name = isValidSheetName(s.name) ? s.name : sanitizeName(s.name);
		const base = name;
		for (let n = 2; names.has(name.toLowerCase()); n++) {
			const tail = ` (${n})`;
			name = base.slice(0, 31 - tail.length) + tail;
		}
		names.add(name.toLowerCase());
		if (name !== s.name) {
			ctx.warn(
				'sheet-renamed',
				`Sheet "${s.name}" was saved as "${name}" (invalid or duplicate name)`,
				{ sheet: s.name },
			);
		}
		let sheetId = s.sheetId;
		if (!Number.isInteger(sheetId) || sheetId <= 0 || ids.has(sheetId)) {
			sheetId = ++maxId;
		}
		ids.add(sheetId);
		return { name, sheetId };
	});
}

/** Give every sheet an output part and a workbook relationship; drop rels of removed sheets. */
export function assignSheetParts(ctx: WriteContext, wb: WorkbookPart): SheetTarget[] {
	const claimed = new Set<string>();
	const targets: SheetTarget[] = [];
	for (const sheet of ctx.workbook.sheets) {
		const src = sheet.source?.partPath;
		const key = src?.toLowerCase();
		if (src && key && ctx.sourceHas(src) && !claimed.has(key)) {
			claimed.add(key);
			const rel = wb.rels.find(
				(r) =>
					SHEET_TYPES.has(relSuffix(r.type)) &&
					resolveRelTarget(wb.path, r.target).toLowerCase() === key,
			);
			const relId =
				rel?.id ??
				addRel(
					wb.rels,
					sheet.kind === 'chartsheet' ? REL.chartsheet : REL.worksheet,
					relativeTarget(wb.path, src),
				);
			targets.push({ sheet, partPath: src, relId, hasSource: true, fresh: false });
			continue;
		}
		if (sheet.kind !== 'worksheet') {
			ctx.warn('sheet-kind', `Sheet "${sheet.name}" (${sheet.kind}) is saved as a worksheet`, {
				sheet: sheet.name,
			});
		}
		const partPath = ctx.allocate('xl/worksheets', 'sheet', 'xml');
		const relId = addRel(wb.rels, REL.worksheet, relativeTarget(wb.path, partPath));
		targets.push({ sheet, partPath, relId, hasSource: false, fresh: !!src });
	}
	const used = new Set(targets.map((t) => t.relId));
	wb.rels = wb.rels.filter((r) => !SHEET_TYPES.has(relSuffix(r.type)) || used.has(r.id));
	return targets;
}

/** Remove the calculation chain (Excel rebuilds it); needed whenever cells were rewritten. */
export function removeCalcChain(ctx: WriteContext, wb: WorkbookPart): void {
	for (const r of wb.rels.filter((x) => relSuffix(x.type) === 'calcChain')) {
		ctx.remove(resolveRelTarget(wb.path, r.target));
	}
	wb.rels = wb.rels.filter((x) => relSuffix(x.type) !== 'calcChain');
	if (ctx.has('xl/calcChain.xml')) {
		ctx.remove('xl/calcChain.xml');
	}
}

/** Ensure a workbook relationship of `type` exists; returns the target part path. */
export function ensureWorkbookRel(wb: WorkbookPart, type: string, defaultPath: string): string {
	const rel = findRelByType(wb.rels, type);
	if (rel) {
		return resolveRelTarget(wb.path, rel.target);
	}
	addRel(wb.rels, type, relativeTarget(wb.path, defaultPath));
	return defaultPath;
}

/** Path of an existing workbook relationship target, if any. */
export function workbookRelPath(wb: WorkbookPart, type: string): string | undefined {
	const rel = findRelByType(wb.rels, type);
	return rel ? resolveRelTarget(wb.path, rel.target) : undefined;
}

/** Write `workbook.xml` (only when it changed) and the workbook relationships. */
export async function writeWorkbookPart(
	ctx: WriteContext,
	wb: WorkbookPart,
	targets: readonly SheetTarget[],
	fullCalcOnLoad: boolean,
): Promise<WorkbookSheetEntry[]> {
	const identities = sheetIdentities(ctx);
	const visible = ctx.workbook.sheets.some((s) => s.state === 'visible');
	const entries: WorkbookSheetEntry[] = targets.map((t, i) => ({
		sheet: !visible && i === 0 ? { ...t.sheet, state: 'visible' } : t.sheet,
		relId: t.relId,
		name: identities[i].name,
		sheetId: identities[i].sheetId,
	}));
	const root = buildWorkbookElement({
		workbook: ctx.workbook,
		original: wb.root,
		sheets: entries,
		originalSheetNames: wb.originalSheets.map((s) => s.name),
		fullCalcOnLoad,
	});
	const changed =
		!wb.root || !!ctx.options.regenerateAll || serializeElement(root) !== serializeElement(wb.root);
	if (changed) {
		ctx.writeText(wb.path, serializeXml(root), wb.root ? undefined : CT.workbook);
	}
	await ctx.writeRelsIfChanged(wb.path, wb.rels);
	return entries;
}
