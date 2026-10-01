/**
 * Parts a worksheet links to (tables, drawing, notes + VML, threaded
 * comments): reuse the original part and relationship id when there is one,
 * create new parts otherwise, and drop links that are no longer needed.
 */
import { relativeTarget, resolveRelTarget } from '../package';
import type { Relationship, Worksheet } from '../types/sheet';
import type { Table } from '../types/table';
import { CT, REL } from '../xml/namespaces';
import type { XmlElement } from '../xml/types';
import { buildCommentsXml, legacyComments } from './comments';
import { addRel, type WriteContext } from './context';
import { writeDrawingPart } from './drawings';
import { SheetGeometry } from './sheet-geometry';
import { buildTableXml } from './tables';
import { buildThreadedCommentsXml, CT_THREADED_COMMENTS, type GuidMap } from './threaded-comments';
import { buildVmlXml, hasKeptShapes, parseOriginalVml, type VmlBlockAllocator } from './vml';
import { findChild } from './xml-utils';

/** Workbook-wide state the sheet link planner needs. */
export interface LinkState {
	ctx: WriteContext;
	tables: Map<Table, { id: number; name: string }>;
	vmlBlocks(): Promise<VmlBlockAllocator>;
	threadIds: GuidMap;
	personIds: GuidMap;
	/** Person ids referenced by threaded comments written in this save. */
	referencedPersons: Set<string>;
	threadedWritten: boolean;
}

export interface SheetLinkInput {
	sheet: Worksheet;
	partPath: string;
	/** The sheet part comes from the source package (not a copy, not new). */
	hasSource: boolean;
	/** Sheet copy: links of the original must not be shared. */
	fresh: boolean;
	root?: XmlElement;
	rels: Relationship[];
}

export interface SheetLinks {
	/** `undefined` keeps the root's element, `null` removes it. */
	drawing?: string | null;
	legacyDrawing?: string | null;
	tableParts: string[];
}

function suffix(type: string): string {
	return type.slice(type.lastIndexOf('/') + 1);
}

function isType(r: Relationship, type: string): boolean {
	return r.targetMode !== 'External' && suffix(r.type) === suffix(type);
}

function removeRel(rels: Relationship[], id: string | undefined): void {
	const i = rels.findIndex((r) => r.id === id);
	if (i >= 0) {
		rels.splice(i, 1);
	}
}

function target(input: SheetLinkInput, r: Relationship): string {
	return resolveRelTarget(input.partPath, r.target);
}

function rootRelId(root: XmlElement | undefined, name: string): string | undefined {
	const el = findChild(root, name);
	if (!el) {
		return undefined;
	}
	for (const k in el.attrs) {
		if (k === 'r:id' || k.endsWith(':id')) {
			return el.attrs[k];
		}
	}
	return undefined;
}

async function planTables(
	state: LinkState,
	input: SheetLinkInput,
	clean: boolean,
): Promise<string[]> {
	const { ctx } = state;
	const { sheet } = input;
	const ids: string[] = [];
	const kept = new Set<string>();
	for (const table of sheet.tables) {
		const assigned = state.tables.get(table)!;
		const rel =
			input.hasSource && table.partPath
				? input.rels.find((r) => isType(r, REL.table) && target(input, r) === table.partPath)
				: undefined;
		let path = table.partPath;
		let relId = rel?.id;
		if (!rel || !path) {
			path = ctx.allocate('xl/tables', 'table', 'xml');
			relId = addRel(input.rels, REL.table, relativeTarget(input.partPath, path));
		}
		kept.add(relId!);
		ids.push(relId!);
		const unchanged =
			!!rel &&
			table.dirty === false &&
			!ctx.options.regenerateAll &&
			assigned.id === table.id &&
			assigned.name === table.name;
		if (!unchanged) {
			const renamed =
				assigned.name === table.name
					? table
					: { ...table, name: assigned.name, displayName: assigned.name };
			ctx.writeText(path, buildTableXml(renamed, assigned.id), CT.table);
		}
	}
	for (const r of [...input.rels]) {
		if (isType(r, REL.table) && !kept.has(r.id)) {
			if (clean) {
				ids.push(r.id);
			} else {
				removeRel(input.rels, r.id);
			}
		}
	}
	return ids;
}

async function planDrawing(
	state: LinkState,
	input: SheetLinkInput,
): Promise<string | null | undefined> {
	const { ctx } = state;
	const { sheet } = input;
	if (input.hasSource && sheet.drawingsDirty === false && !ctx.options.regenerateAll) {
		return undefined;
	}
	const rid = rootRelId(input.root, 'drawing');
	const srcRel = input.hasSource
		? (input.rels.find((r) => r.id === rid && isType(r, REL.drawing)) ??
			input.rels.find((r) => isType(r, REL.drawing)))
		: undefined;
	if (!sheet.drawings.length) {
		removeRel(input.rels, srcRel?.id);
		return null;
	}
	let path: string;
	let relId: string;
	if (srcRel) {
		path = target(input, srcRel);
		relId = srcRel.id;
	} else {
		path = ctx.allocate('xl/drawings', 'drawing', 'xml');
		relId = addRel(input.rels, REL.drawing, relativeTarget(input.partPath, path));
	}
	const inPlace = !!srcRel && ctx.sourceHas(path);
	const written = await writeDrawingPart(ctx, sheet, { path, inPlace, fresh: input.fresh });
	if (!written) {
		removeRel(input.rels, relId);
		return null;
	}
	return relId;
}

async function planComments(
	state: LinkState,
	input: SheetLinkInput,
): Promise<string | null | undefined> {
	const { ctx } = state;
	const { sheet } = input;
	if (input.hasSource && sheet.commentsDirty === false && !ctx.options.regenerateAll) {
		return undefined;
	}
	const ids = new Map(sheet.threadedComments.map((c) => [c.id, state.threadIds.get(c.id)]));
	const { comments, droppedNotes } = legacyComments(sheet, ids);
	if (droppedNotes) {
		ctx.warn(
			'note-dropped',
			`${droppedNotes} note(s) on "${sheet.name}" share a cell with another comment and were not saved`,
			{ sheet: sheet.name },
		);
	}
	const byType = (type: string): Relationship | undefined =>
		input.hasSource ? input.rels.find((r) => isType(r, type)) : undefined;
	const legacyId = rootRelId(input.root, 'legacyDrawing');
	const vmlRel =
		input.hasSource && legacyId
			? input.rels.find((r) => r.id === legacyId && isType(r, REL.vmlDrawing))
			: undefined;
	const commentsRel = byType(REL.comments);
	const threadedRel = byType(REL.threadedComment);
	const vmlPath = vmlRel ? target(input, vmlRel) : undefined;
	const vmlText = vmlPath ? await ctx.sourceText(vmlPath) : undefined;
	const original = vmlText ? parseOriginalVml(vmlText) : undefined;
	if (vmlText && !original && /ObjectType="(?!Note")/.test(vmlText)) {
		ctx.warn('vml-controls-lost', `Form controls in ${vmlPath} could not be preserved`, {
			sheet: sheet.name,
			part: vmlPath,
		});
	}
	let legacy: string | null;
	if (comments.length || hasKeptShapes(original)) {
		const path = vmlPath ?? ctx.allocate('xl/drawings', 'vmlDrawing', 'vml');
		const relId =
			vmlRel?.id ?? addRel(input.rels, REL.vmlDrawing, relativeTarget(input.partPath, path));
		const blocks = comments.length ? (await state.vmlBlocks()).allocate(comments.length) : [];
		const notes = comments.map((c) => ({
			row: c.ref.row,
			col: c.ref.col,
			visible: c.visible,
			width: c.width,
			height: c.height,
		}));
		ctx.writeText(path, buildVmlXml(notes, new SheetGeometry(sheet), blocks, original));
		ctx.writer.contentTypes.ensureDefault('vml', CT.vmlDrawing);
		legacy = relId;
	} else {
		removeRel(input.rels, vmlRel?.id);
		legacy = null;
	}
	if (comments.length) {
		const path = commentsRel ? target(input, commentsRel) : ctx.allocate('xl', 'comments', 'xml');
		if (!commentsRel) {
			addRel(input.rels, REL.comments, relativeTarget(input.partPath, path));
		}
		ctx.writeText(path, buildCommentsXml(comments), CT.comments);
	} else {
		removeRel(input.rels, commentsRel?.id);
	}
	if (sheet.threadedComments.length) {
		const path = threadedRel
			? target(input, threadedRel)
			: ctx.allocate('xl/threadedComments', 'threadedComment', 'xml');
		if (!threadedRel) {
			addRel(input.rels, REL.threadedComment, relativeTarget(input.partPath, path));
		}
		const xml = buildThreadedCommentsXml(
			sheet.threadedComments,
			state.threadIds,
			state.personIds,
			ctx.now,
		);
		ctx.writeText(path, xml, CT_THREADED_COMMENTS);
		for (const c of sheet.threadedComments) {
			state.referencedPersons.add(c.personId || 'unknown');
		}
		state.threadedWritten = true;
	} else {
		removeRel(input.rels, threadedRel?.id);
	}
	return legacy;
}

/** Plan (and write) every linked part; `input.rels` is updated in place. */
export async function planSheetLinks(
	state: LinkState,
	input: SheetLinkInput,
	clean: boolean,
): Promise<SheetLinks> {
	const tableParts = await planTables(state, input, clean);
	const drawing = await planDrawing(state, input);
	const legacyDrawing = await planComments(state, input);
	return { drawing, legacyDrawing, tableParts };
}
