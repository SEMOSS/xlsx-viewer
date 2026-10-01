/**
 * Threaded comments (`xl/threadedComments/threadedCommentN.xml`) and the
 * workbook person list (`xl/persons/person.xml`). Ids must be braced GUIDs;
 * other ids are mapped to fresh GUIDs consistently.
 */
import { a1 } from '../address/cell-address';
import type { Person, ThreadedComment } from '../types/comment';
import { XML_DECLARATION } from '../xml';
import { NS } from '../xml/namespaces';
import { attrString, escapeCellText, isGuid, newGuid } from './xml-utils';

export const CT_THREADED_COMMENTS = 'application/vnd.ms-excel.threadedcomments+xml';
export const CT_PERSONS = 'application/vnd.ms-excel.person+xml';

/** Maps model ids to the GUIDs written to the file (stable within one save). */
export class GuidMap {
	private readonly map = new Map<string, string>();

	get(id: string): string {
		let out = this.map.get(id);
		if (!out) {
			out = isGuid(id) ? id : newGuid();
			this.map.set(id, out);
		}
		return out;
	}

	/** Read-only view of the ids assigned so far. */
	entries(): ReadonlyMap<string, string> {
		return this.map;
	}
}

/** `dT` timestamp: the model value when it is a valid date-time, else `now`. */
function timestamp(created: string | undefined, now: Date): string {
	if (created && !Number.isNaN(Date.parse(created))) {
		return created;
	}
	return now
		.toISOString()
		.replace(/Z$/, '')
		.replace(/(\.\d{2})\d*$/, '$1');
}

const ROOT_ATTRS = { xmlns: NS.threadedComments, 'xmlns:x': NS.main };

/** Serialize one sheet's threaded comments. */
export function buildThreadedCommentsXml(
	comments: readonly ThreadedComment[],
	ids: GuidMap,
	persons: GuidMap,
	now: Date,
): string {
	const known = new Set(comments.map((c) => c.id));
	const body = comments
		.map((c) => {
			const attrs = attrString({
				ref: a1(c.ref.row, c.ref.col),
				dT: timestamp(c.created, now),
				personId: persons.get(c.personId || 'unknown'),
				id: ids.get(c.id),
				parentId: c.parentId && known.has(c.parentId) ? ids.get(c.parentId) : undefined,
				done: !c.parentId && c.done ? 1 : undefined,
			});
			return `<threadedComment${attrs}><text>${escapeCellText(c.text)}</text></threadedComment>`;
		})
		.join('');
	return `${XML_DECLARATION}<ThreadedComments${attrString(ROOT_ATTRS)}>${body}</ThreadedComments>`;
}

/** Serialize the person list. */
export function buildPersonsXml(persons: readonly Person[], ids: GuidMap): string {
	const body = persons
		.map(
			(p) =>
				`<person${attrString({
					displayName: p.displayName,
					id: ids.get(p.id),
					userId: p.userId,
					providerId: p.providerId ?? 'None',
				})}/>`,
		)
		.join('');
	return `${XML_DECLARATION}<personList${attrString(ROOT_ATTRS)}>${body}</personList>`;
}
