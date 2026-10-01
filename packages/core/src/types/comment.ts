import type { CellAddress } from './address';
import type { RichTextRun } from './cell';

/** A legacy note (what Excel 365 calls a "Note"; stored in `commentsN.xml` + VML). */
export interface Note {
	ref: CellAddress;
	author: string;
	/** Plain text of the note. */
	text: string;
	/** Rich-text runs, when the note text is formatted. */
	runs?: RichTextRun[];
	/** The note box is always shown (not just on hover). */
	visible?: boolean;
	/** Note box size in pixels, from the VML shape, when known. */
	width?: number;
	height?: number;
}

/** A modern threaded comment (Excel 365). */
export interface ThreadedComment {
	id: string;
	ref: CellAddress;
	personId: string;
	text: string;
	/** ISO 8601 timestamp. */
	created?: string;
	/** Id of the thread's first comment, for replies. */
	parentId?: string;
	done?: boolean;
}

export interface Person {
	id: string;
	displayName: string;
	userId?: string;
	providerId?: string;
}
