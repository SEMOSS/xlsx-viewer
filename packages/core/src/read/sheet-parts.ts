import { REL } from '../xml/namespaces';
import type { Table } from '../types/table';
import type { Worksheet } from '../types/sheet';
import type { XmlElement } from '../xml/types';
import { parseNotes, parseThreadedComments } from './comments';
import { attempt, type ReadContext } from './context';
import { readDrawingPart } from './drawing';
import { readPartText, relTypeSuffix, type RelLookup } from './part-io';
import { parseTable } from './tables';
import { parseVmlNoteShapes, type VmlNoteShape } from './vml';
import { kid, kids, str } from './xml-utils';

/** Parts related to a worksheet: tables, notes (+ VML), threaded comments and the drawing. */

function relPathByType(rels: RelLookup, type: string): string | undefined {
	const suffix = relTypeSuffix(type);
	const rel = rels.rels.find((r) => r.type === type || relTypeSuffix(r.type) === suffix);
	return rel ? rels.path(rel.id) : undefined;
}

async function readTables(
	ctx: ReadContext,
	sheet: Worksheet,
	root: XmlElement,
	rels: RelLookup,
): Promise<Table[]> {
	const out: Table[] = [];
	for (const part of kids(kid(root, 'tableParts'), 'tablePart')) {
		const path = rels.path(str(part, 'r:id'));
		if (!path) {
			continue;
		}
		const table = await attempt(
			ctx,
			path,
			'table',
			async () => {
				const xml = await readPartText(ctx.pkg, path);
				return xml === undefined ? undefined : parseTable(xml, path);
			},
			undefined,
			sheet.name,
		);
		if (table) {
			out.push(table);
			if (table.raw?.attrs.tableType === 'queryTable') {
				ctx.warn(
					'external-data',
					`Table "${table.name}" is linked to an external query; data is not refreshed`,
					{
						part: path,
						sheet: sheet.name,
					},
				);
			}
		}
	}
	return out;
}

/** Read every related part of a worksheet and record its path in `sheet.source`. */
export async function readSheetParts(
	ctx: ReadContext,
	sheet: Worksheet,
	root: XmlElement,
	rels: RelLookup,
): Promise<void> {
	const source = sheet.source!;
	sheet.tables = await readTables(ctx, sheet, root, rels);

	const vmlPath = rels.path(str(kid(root, 'legacyDrawing'), 'r:id'));
	if (vmlPath) {
		source.vmlPath = vmlPath;
	}
	const commentsPath = relPathByType(rels, REL.comments);
	if (commentsPath) {
		source.commentsPath = commentsPath;
		sheet.notes = await attempt(
			ctx,
			commentsPath,
			'comments',
			async () => {
				const xml = await readPartText(ctx.pkg, commentsPath);
				if (xml === undefined) {
					return [];
				}
				let shapes: VmlNoteShape[] = [];
				if (vmlPath) {
					const vml = await readPartText(ctx.pkg, vmlPath);
					shapes = vml ? parseVmlNoteShapes(vml) : [];
				}
				return parseNotes(xml, shapes);
			},
			[],
			sheet.name,
		);
	}

	const threadedPath = relPathByType(rels, REL.threadedComment);
	if (threadedPath) {
		source.threadedCommentsPath = threadedPath;
		sheet.threadedComments = await attempt(
			ctx,
			threadedPath,
			'threaded comments',
			async () => {
				const xml = await readPartText(ctx.pkg, threadedPath);
				return xml === undefined ? [] : parseThreadedComments(xml);
			},
			[],
			sheet.name,
		);
	}

	const drawingPath =
		rels.path(str(kid(root, 'drawing'), 'r:id')) ?? relPathByType(rels, REL.drawing);
	if (drawingPath) {
		source.drawingPath = drawingPath;
		if (!ctx.options.skipDrawings) {
			sheet.drawings = await attempt(
				ctx,
				drawingPath,
				'drawing',
				() => readDrawingPart(ctx, drawingPath, sheet.name),
				[],
				sheet.name,
			);
		}
	}
}
