/**
 * Legacy comments part (`xl/commentsN.xml`). Notes are written as they are;
 * every threaded-comment thread also gets the read-only mirror note Excel
 * writes for older versions (author `tc={threadId}`).
 */
import { a1 } from '../address/cell-address';
import type { CellAddress } from '../types/address';
import type { RichTextRun } from '../types/cell';
import type { ThreadedComment } from '../types/comment';
import type { Worksheet } from '../types/sheet';
import { XML_DECLARATION } from '../xml';
import { NS } from '../xml/namespaces';
import { richRunsXml } from './rich-text';
import { attrString, escapeCellText, textElement } from './xml-utils';

/** One `<comment>` of the legacy part. */
export interface LegacyComment {
	ref: CellAddress;
	author: string;
	text: string;
	runs?: RichTextRun[];
	/** `xr:uid` (the thread id for threaded-comment mirrors). */
	uid?: string;
	visible?: boolean;
	/** Note box size in pixels. */
	width?: number;
	height?: number;
}

const MIRROR_HEADER =
	'[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, any ' +
	'edits to it will get removed if the file is opened in a newer version of Excel. Learn more: ' +
	'https://go.microsoft.com/fwlink/?linkid=870924\n\nComment:\n    ';

/** Threads in document order: each root comment with its replies. */
export function threadsOf(
	comments: readonly ThreadedComment[],
): Array<{ root: ThreadedComment; replies: ThreadedComment[] }> {
	const byId = new Map<string, { root: ThreadedComment; replies: ThreadedComment[] }>();
	const out: Array<{ root: ThreadedComment; replies: ThreadedComment[] }> = [];
	for (const c of comments) {
		if (!c.parentId) {
			const thread = { root: c, replies: [] as ThreadedComment[] };
			byId.set(c.id, thread);
			out.push(thread);
		}
	}
	for (const c of comments) {
		if (c.parentId) {
			byId.get(c.parentId)?.replies.push(c);
		}
	}
	return out;
}

/** The mirror text Excel stores for a thread. */
export function mirrorText(root: ThreadedComment, replies: readonly ThreadedComment[]): string {
	return MIRROR_HEADER + root.text + replies.map((r) => `\nReply:\n    ${r.text}`).join('');
}

/**
 * Comments for the legacy part: the sheet's notes plus thread mirrors.
 * `ids` maps thread ids to the ids written in the threaded part. A note on a
 * cell that also has a thread is dropped (Excel allows only one).
 */
export function legacyComments(
	sheet: Worksheet,
	ids: ReadonlyMap<string, string>,
): { comments: LegacyComment[]; droppedNotes: number } {
	const threads = threadsOf(sheet.threadedComments);
	const threadCells = new Set(threads.map((t) => `${t.root.ref.row},${t.root.ref.col}`));
	const comments: LegacyComment[] = [];
	let droppedNotes = 0;
	const noteCells = new Set<string>();
	for (const n of sheet.notes) {
		const key = `${n.ref.row},${n.ref.col}`;
		if (threadCells.has(key) || noteCells.has(key)) {
			droppedNotes++;
			continue;
		}
		noteCells.add(key);
		comments.push({
			ref: n.ref,
			author: n.author,
			text: n.text,
			runs: n.runs,
			visible: n.visible,
			width: n.width,
			height: n.height,
		});
	}
	const seen = new Set<string>();
	for (const t of threads) {
		const key = `${t.root.ref.row},${t.root.ref.col}`;
		if (seen.has(key)) {
			continue;
		}
		seen.add(key);
		const id = ids.get(t.root.id) ?? t.root.id;
		comments.push({
			ref: t.root.ref,
			author: `tc=${id}`,
			text: mirrorText(t.root, t.replies),
			uid: id,
		});
	}
	comments.sort((a, b) => a.ref.row - b.ref.row || a.ref.col - b.ref.col);
	return { comments, droppedNotes };
}

/** Serialize the legacy comments part. */
export function buildCommentsXml(comments: readonly LegacyComment[]): string {
	const authors: string[] = [];
	const authorIndex = new Map<string, number>();
	for (const c of comments) {
		if (!authorIndex.has(c.author)) {
			authorIndex.set(c.author, authors.length);
			authors.push(c.author);
		}
	}
	const withUid = comments.some((c) => c.uid);
	const rootAttrs: Record<string, string> = { xmlns: NS.main };
	if (withUid) {
		rootAttrs['xmlns:mc'] = NS.mc;
		rootAttrs['mc:Ignorable'] = 'xr';
		rootAttrs['xmlns:xr'] = NS.xr;
	}
	const authorXml = authors.map((a) => `<author>${escapeCellText(a)}</author>`).join('');
	const list = comments
		.map((c) => {
			const attrs = attrString({
				ref: a1(c.ref.row, c.ref.col),
				authorId: authorIndex.get(c.author),
				shapeId: 0,
				'xr:uid': c.uid,
			});
			const text = c.runs?.length ? richRunsXml(c.runs) : textElement(c.text);
			return `<comment${attrs}><text>${text}</text></comment>`;
		})
		.join('');
	return (
		XML_DECLARATION +
		`<comments${attrString(rootAttrs)}><authors>${authorXml}</authors><commentList>${list}</commentList></comments>`
	);
}
