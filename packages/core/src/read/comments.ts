import { parseCellAddress } from '../address/cell-address';
import type { Note, Person, ThreadedComment } from '../types/comment';
import { parseXml } from '../xml/parser';
import { parseRichElement } from './rich-text';
import type { VmlNoteShape } from './vml';
import { parseSpreadsheetXml } from './xml-normalize';
import { bool, kid, kids, ooxmlText, str } from './xml-utils';

/**
 * Legacy notes (`xl/commentsN.xml`). Notes whose author starts with `tc=`
 * are the compatibility copies Excel writes for threaded comments and are
 * skipped (the threaded comments are read separately). Visibility and box
 * size come from the matching VML shape, when one is given.
 */
export function parseNotes(xml: string, shapes: VmlNoteShape[] = []): Note[] {
	const root = parseSpreadsheetXml(xml).root;
	const authors = kids(kid(root, 'authors'), 'author').map((a) => ooxmlText(a));
	const shapeAt = new Map<string, VmlNoteShape>();
	for (const s of shapes) {
		shapeAt.set(`${s.row},${s.col}`, s);
	}
	const notes: Note[] = [];
	for (const c of kids(kid(root, 'commentList'), 'comment')) {
		const ref = parseCellAddress(str(c, 'ref') ?? '');
		if (!ref) {
			continue;
		}
		const author = authors[Number(str(c, 'authorId') ?? 0)] ?? '';
		if (author.startsWith('tc=')) {
			continue;
		}
		const textEl = kid(c, 'text');
		const item = textEl ? parseRichElement(textEl) : { text: '' };
		const note: Note = { ref, author, text: item.text };
		if (item.runs) {
			note.runs = item.runs;
		}
		const shape = shapeAt.get(`${ref.row},${ref.col}`);
		if (shape) {
			if (shape.visible) {
				note.visible = true;
			}
			if (shape.width !== undefined) {
				note.width = shape.width;
			}
			if (shape.height !== undefined) {
				note.height = shape.height;
			}
		}
		notes.push(note);
	}
	return notes;
}

/** Threaded comments (`xl/threadedComments/threadedCommentN.xml`), matched by local name. */
export function parseThreadedComments(xml: string): ThreadedComment[] {
	const root = parseXml(xml).root;
	const out: ThreadedComment[] = [];
	for (const c of kids(root, 'threadedComment')) {
		const ref = parseCellAddress(str(c, 'ref') ?? '');
		const id = str(c, 'id');
		if (!ref || !id) {
			continue;
		}
		const comment: ThreadedComment = {
			id,
			ref,
			personId: str(c, 'personId') ?? '',
			text: ooxmlText(kid(c, 'text')),
		};
		const created = str(c, 'dT');
		if (created) {
			comment.created = created;
		}
		const parentId = str(c, 'parentId');
		if (parentId) {
			comment.parentId = parentId;
		}
		const done = bool(c, 'done');
		if (done !== undefined) {
			comment.done = done;
		}
		out.push(comment);
	}
	return out;
}

/** The person list (`xl/persons/person.xml`) of threaded-comment authors. */
export function parsePersons(xml: string): Person[] {
	const root = parseXml(xml).root;
	const out: Person[] = [];
	for (const p of kids(root, 'person')) {
		const id = str(p, 'id');
		if (!id) {
			continue;
		}
		const person: Person = { id, displayName: str(p, 'displayName') ?? '' };
		const userId = str(p, 'userId');
		if (userId) {
			person.userId = userId;
		}
		const providerId = str(p, 'providerId');
		if (providerId) {
			person.providerId = providerId;
		}
		out.push(person);
	}
	return out;
}
