/**
 * Workbook-wide parts written after the sheets: shared strings, styles,
 * cell metadata, the person list and the theme.
 */
import type { Person } from '../types/comment';
import type { Table } from '../types/table';
import { CT, REL } from '../xml/namespaces';
import type { WriteContext } from './context';
import { CellMetadata, CT_SHEET_METADATA } from './metadata';
import { SharedStringTable } from './shared-strings';
import type { SheetTarget, SheetWriteState } from './sheet-writer';
import { buildStylesXml, parseStylesRoot, stylesChanged } from './styles-xml';
import { buildThemeXml } from './theme-xml';
import { buildPersonsXml, CT_PERSONS, GuidMap } from './threaded-comments';
import { VmlBlockAllocator, vmlBlocks } from './vml';
import { ensureWorkbookRel, type WorkbookPart, workbookRelPath } from './workbook-part';

export interface WriterState extends SheetWriteState {
	/** The shared string table was loaded (some sheet was regenerated). */
	sstLoaded(): boolean;
}

/** Workbook-unique table ids and names (untouched tables keep theirs). */
export function assignTables(ctx: WriteContext): Map<Table, { id: number; name: string }> {
	const out = new Map<Table, { id: number; name: string }>();
	const all = ctx.workbook.sheets.flatMap((s) => s.tables);
	const ids = new Set<number>();
	const names = new Set<string>();
	const keep = all.filter((t) => t.dirty === false && t.partPath && ctx.sourceHas(t.partPath));
	const rest = all.filter((t) => !keep.includes(t));
	let maxId = all.reduce((m, t) => (Number.isInteger(t.id) ? Math.max(m, t.id) : m), 0);
	const kept = new Set(keep);
	for (const t of [...keep, ...rest]) {
		let id = t.id;
		if (!Number.isInteger(id) || id <= 0 || ids.has(id)) {
			id = ++maxId;
		}
		ids.add(id);
		const valid =
			kept.has(t) ||
			(!!t.name &&
				t.name.length <= 255 &&
				/^[\p{L}_\\][\p{L}\p{N}_.\\]*$/u.test(t.name) &&
				!/^[A-Za-z]{1,3}\d+$/.test(t.name) &&
				!/^[RrCc]$/.test(t.name));
		let name = valid ? t.name : `Table${id}`;
		const base = name;
		for (let n = 2; names.has(name.toLowerCase()); n++) {
			name = `${base}_${n}`;
		}
		names.add(name.toLowerCase());
		out.set(t, { id, name });
	}
	return out;
}

/** State shared by every sheet writer during one save. */
export async function createWriterState(ctx: WriteContext, wb: WorkbookPart): Promise<WriterState> {
	let sst: Promise<SharedStringTable> | undefined;
	let blocks: Promise<VmlBlockAllocator> | undefined;
	const metadataPath = workbookRelPath(wb, REL.sheetMetadata);
	let metadata = new CellMetadata();
	if (metadataPath) {
		const xml = await ctx.sourceText(metadataPath);
		if (xml) {
			try {
				metadata = CellMetadata.parse(xml, metadataPath);
			} catch {
				ctx.warn('unreadable-part', `Cell metadata ${metadataPath} could not be read`, {
					part: metadataPath,
				});
			}
		}
	}
	return {
		ctx,
		metadata,
		tables: assignTables(ctx),
		threadIds: new GuidMap(),
		personIds: new GuidMap(),
		referencedPersons: new Set(),
		threadedWritten: false,
		sstLoaded: () => sst !== undefined,
		sst: () => {
			sst ??= (async () => {
				const path = workbookRelPath(wb, REL.sharedStrings);
				const xml = path ? await ctx.sourceText(path) : undefined;
				return xml ? SharedStringTable.fromXml(xml) : SharedStringTable.empty();
			})();
			return sst;
		},
		vmlBlocks: () => {
			blocks ??= (async () => {
				const alloc = new VmlBlockAllocator();
				for (const p of ctx.source?.partPaths() ?? []) {
					if (p.toLowerCase().endsWith('.vml')) {
						alloc.reserve(vmlBlocks((await ctx.sourceText(p)) ?? ''));
					}
				}
				return alloc;
			})();
			return blocks;
		},
	};
}

/** Shared strings: written for new workbooks, or when strings were appended. */
async function writeSharedStrings(
	ctx: WriteContext,
	wb: WorkbookPart,
	state: WriterState,
	allRegenerated: boolean,
): Promise<void> {
	if (!state.sstLoaded() && ctx.source) {
		return;
	}
	const sst = await state.sst();
	const existing = workbookRelPath(wb, REL.sharedStrings);
	const present = !!existing && ctx.has(existing);
	if (ctx.source && present && !sst.grew && !ctx.options.regenerateAll) {
		return;
	}
	if (ctx.source && !present && sst.size === 0) {
		return;
	}
	const path = ensureWorkbookRel(wb, REL.sharedStrings, existing ?? 'xl/sharedStrings.xml');
	ctx.writeText(path, sst.serialize(allRegenerated), CT.sharedStrings);
	if (sst.truncated) {
		ctx.warn(
			'text-truncated',
			`${sst.truncated} string(s) longer than 32767 characters were truncated`,
		);
	}
}

async function writeStyles(ctx: WriteContext, wb: WorkbookPart): Promise<void> {
	const styles = ctx.styles;
	const existing = workbookRelPath(wb, REL.styles);
	const present = !!existing && ctx.sourceHas(existing);
	let root = styles.raw?.root;
	if (!root && present) {
		root = parseStylesRoot((await ctx.sourceText(existing!)) ?? '');
	}
	if (present && !ctx.options.regenerateAll && !stylesChanged(styles, root)) {
		return;
	}
	const path = ensureWorkbookRel(wb, REL.styles, existing ?? 'xl/styles.xml');
	ctx.writeText(path, buildStylesXml(styles, present ? root : undefined), CT.styles);
}

function writeMetadata(ctx: WriteContext, wb: WorkbookPart, state: WriterState): void {
	if (!state.metadata.changed) {
		return;
	}
	const path = ensureWorkbookRel(
		wb,
		REL.sheetMetadata,
		state.metadata.sourcePath ?? 'xl/metadata.xml',
	);
	ctx.writeText(path, state.metadata.serialize(), CT_SHEET_METADATA);
}

function writePersons(ctx: WriteContext, wb: WorkbookPart, state: WriterState): void {
	if (!state.threadedWritten) {
		return;
	}
	const persons: Person[] = [...ctx.workbook.persons];
	const known = new Set(persons.map((p) => p.id));
	for (const id of state.referencedPersons) {
		if (!known.has(id)) {
			persons.push({ id, displayName: 'Unknown' });
			known.add(id);
		}
	}
	const path = ensureWorkbookRel(wb, REL.person, 'xl/persons/person.xml');
	ctx.writeText(path, buildPersonsXml(persons, state.personIds), CT_PERSONS);
}

async function writeTheme(ctx: WriteContext, wb: WorkbookPart): Promise<void> {
	const theme = ctx.workbook.theme;
	const existing = workbookRelPath(wb, REL.theme);
	if (existing && ctx.has(existing)) {
		// The theme is read-only in the model; a missing rawXml marks an edited theme.
		if (theme.rawXml === undefined && (ctx.options.regenerateAll || ctx.workbook.dirty !== false)) {
			ctx.writeText(existing, buildThemeXml(theme), CT.theme);
		}
		return;
	}
	const path = ensureWorkbookRel(wb, REL.theme, existing ?? 'xl/theme/theme1.xml');
	ctx.writeText(path, theme.rawXml ?? buildThemeXml(theme), CT.theme);
}

/** Write every workbook-level part that depends on what the sheets wrote. */
export async function writeSharedParts(
	ctx: WriteContext,
	wb: WorkbookPart,
	state: WriterState,
	targets: readonly SheetTarget[],
	regenerated: ReadonlySet<SheetTarget>,
): Promise<void> {
	const allRegenerated = targets.every((t) => regenerated.has(t) || t.sheet.kind !== 'worksheet');
	await writeSharedStrings(ctx, wb, state, allRegenerated);
	await writeStyles(ctx, wb);
	writeMetadata(ctx, wb, state);
	writePersons(ctx, wb, state);
	await writeTheme(ctx, wb);
}
