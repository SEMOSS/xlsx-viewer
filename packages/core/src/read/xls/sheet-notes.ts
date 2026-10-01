/**
 * Legacy cell notes: a NOTE record names the cell, author and visibility;
 * its text lives in the TXO record that follows the note's OBJ record (the
 * characters and formatting runs sit in the TXO's CONTINUE records). OBJ
 * records of other kinds (pictures, shapes, controls) are counted so the
 * reader can report what was skipped.
 */
import type { Note } from '../../types/comment';
import type { RichTextRun } from '../../types/cell';
import { BiffCursor } from './biff-cursor';
import { RT } from './biff-opcodes';
import type { BiffRecord } from './biff-records';
import { type SheetContext, validCell } from './sheet-context';
import { type BiffRun, richTextRuns } from './sst';

const OT_CHART = 0x05;
const OT_PICTURE = 0x08;
const OT_DROPDOWN = 0x14;
const OT_NOTE = 0x19;

interface NoteRecord {
	row: number;
	col: number;
	visible: boolean;
	objId: number;
	author: string;
}

interface TextBox {
	text: string;
	runs?: BiffRun[];
}

/** TXO: text in the first CONTINUE (option byte, characters), then 8-byte formatting runs. */
function parseTxo(record: BiffRecord): TextBox {
	const cur = new BiffCursor(record.data, record.breaks);
	cur.skip(10);
	const cch = cur.u16();
	const cbRuns = cur.u16();
	if (cch === 0 || record.breaks.length === 0) {
		return { text: '' };
	}
	cur.pos = record.breaks[0];
	const text = cur.chars(cch, false).replace(/\r\n?/g, '\n');
	const runs: BiffRun[] = [];
	cur.pos = cur.nextBreak();
	const count = Math.min(Math.floor(cbRuns / 8), Math.floor(cur.remaining / 8));
	for (let i = 0; i < count; i++) {
		runs.push({ ich: cur.u16(), ifnt: cur.u16() });
		cur.skip(4);
	}
	return runs.length ? { text, runs } : { text };
}

/** Record types handled by {@link NoteReader.read}. */
export const NOTE_RECORDS: readonly number[] = [RT.OBJ, RT.TXO, RT.NOTE];

export class NoteReader {
	private readonly notes: NoteRecord[] = [];
	private readonly texts = new Map<number, TextBox>();
	private lastObj: number | undefined;
	/** Skipped drawing objects by kind. */
	readonly skipped = { charts: 0, pictures: 0, other: 0 };

	constructor(private readonly ctx: SheetContext) {}

	/** Handle NOTE / OBJ / TXO records; returns `false` for other record types. */
	read(record: BiffRecord): boolean {
		switch (record.type) {
			case RT.OBJ: {
				const cur = new BiffCursor(record.data, record.breaks);
				cur.skip(4);
				const type = cur.u16();
				this.lastObj = cur.u16();
				if (type === OT_CHART) {
					this.skipped.charts++;
				} else if (type === OT_PICTURE) {
					this.skipped.pictures++;
				} else if (type !== OT_NOTE && type !== OT_DROPDOWN) {
					this.skipped.other++;
				}
				return true;
			}
			case RT.TXO:
				if (this.lastObj !== undefined) {
					this.texts.set(this.lastObj, parseTxo(record));
				}
				return true;
			case RT.NOTE: {
				const cur = new BiffCursor(record.data, record.breaks);
				const [row, col, flags, objId] = [cur.u16(), cur.u16(), cur.u16(), cur.u16()];
				const author = cur.remaining >= 3 ? cur.unicodeString() : '';
				this.notes.push({ row, col, visible: (flags & 0x02) !== 0, objId, author });
				return true;
			}
			default:
				return false;
		}
	}

	/** The notes of the sheet, with their text. */
	finish(): Note[] {
		const out: Note[] = [];
		for (const n of this.notes) {
			if (!validCell(n.row, n.col)) {
				continue;
			}
			const box = this.texts.get(n.objId) ?? { text: '' };
			const note: Note = { ref: { row: n.row, col: n.col }, author: n.author, text: box.text };
			const runs: RichTextRun[] | undefined = box.runs
				? richTextRuns(box.text, box.runs, (f) => this.ctx.styles.font(f))
				: undefined;
			if (runs && runs.length > 1) {
				note.runs = runs;
			}
			if (n.visible) {
				note.visible = true;
			}
			out.push(note);
		}
		return out;
	}
}
