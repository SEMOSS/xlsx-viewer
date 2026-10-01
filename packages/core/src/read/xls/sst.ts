/**
 * Shared string table (SST): `XLUnicodeRichExtendedString` entries packed
 * into the SST record and its CONTINUE records. A string's characters may
 * cross a record boundary, where a fresh option byte switches between
 * 8-bit and 16-bit storage; formatting runs and phonetic data (ExtRst)
 * follow the characters and read straight across boundaries.
 */
import type { RichTextRun } from '../../types/cell';
import type { FontSpec } from '../../types/style';
import { BiffCursor, BiffTruncatedError } from './biff-cursor';
import type { BiffRecord } from './biff-records';

/** A formatting run of a BIFF rich string: from character `ich` on, BIFF font `ifnt`. */
export interface BiffRun {
	ich: number;
	ifnt: number;
}

export interface XlsSharedStrings {
	strings: string[];
	/** Formatting runs by string index, for rich strings. */
	runs: Map<number, BiffRun[]>;
}

interface BiffRichString {
	text: string;
	runs?: BiffRun[];
}

/** Read one `XLUnicodeRichExtendedString` (phonetic data is skipped). */
function readRichExtendedString(cur: BiffCursor): BiffRichString {
	const cch = cur.u16();
	const flags = cur.u8();
	const rich = (flags & 0x08) !== 0;
	const hasExt = (flags & 0x04) !== 0;
	const cRun = rich ? cur.u16() : 0;
	const cbExtRst = hasExt ? cur.u32() : 0;
	const text = cur.chars(cch, (flags & 0x01) !== 0);
	const out: BiffRichString = { text };
	if (cRun > 0) {
		const runs: BiffRun[] = [];
		for (let i = 0; i < cRun; i++) {
			runs.push({ ich: cur.u16(), ifnt: cur.u16() });
		}
		out.runs = runs;
	}
	if (cbExtRst > 0) {
		cur.skip(Math.min(cbExtRst, cur.remaining));
	}
	return out;
}

/**
 * Parse the SST record (CONTINUE data merged, with breaks). A truncated
 * table keeps the strings read so far and reports `truncated`.
 */
export function parseSst(record: BiffRecord): XlsSharedStrings & { truncated: boolean } {
	const cur = new BiffCursor(record.data, record.breaks);
	const strings: string[] = [];
	const runs = new Map<number, BiffRun[]>();
	let truncated = false;
	try {
		cur.u32(); // total number of references in the workbook
		const unique = cur.u32();
		// Every entry takes at least three bytes, which caps a corrupt count.
		const count = Math.min(unique, Math.floor(cur.remaining / 3));
		truncated = count < unique;
		for (let i = 0; i < count; i++) {
			const s = readRichExtendedString(cur);
			if (s.runs) {
				runs.set(strings.length, s.runs);
			}
			strings.push(s.text);
		}
	} catch (e) {
		if (!(e instanceof BiffTruncatedError)) {
			throw e;
		}
		truncated = true;
	}
	return { strings, runs, truncated };
}

/**
 * Rich-text runs for a string with BIFF formatting runs. Text before the
 * first run uses the cell font (no run font); runs at or past the end of the
 * text (TXO terminators) are dropped. Returns `undefined` when no run applies.
 */
export function richTextRuns(
	text: string,
	runs: readonly BiffRun[],
	fontAt: (ifnt: number) => FontSpec | undefined,
): RichTextRun[] | undefined {
	const starts = new Map<number, number>();
	for (const r of runs) {
		if (r.ich < text.length) {
			starts.set(r.ich, r.ifnt);
		}
	}
	if (starts.size === 0) {
		return undefined;
	}
	const sorted = [...starts.entries()].sort((a, b) => a[0] - b[0]);
	const out: RichTextRun[] = [];
	if (sorted[0][0] > 0) {
		out.push({ text: text.slice(0, sorted[0][0]) });
	}
	sorted.forEach(([ich, ifnt], i) => {
		const end = i + 1 < sorted.length ? sorted[i + 1][0] : text.length;
		const font = fontAt(ifnt);
		out.push(
			font ? { text: text.slice(ich, end), font: { ...font } } : { text: text.slice(ich, end) },
		);
	});
	return out;
}
