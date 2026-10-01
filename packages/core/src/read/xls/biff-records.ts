/**
 * BIFF record framing: `[type: u16][size: u16][data]`, with records longer
 * than 8224 bytes split into CONTINUE records. A workbook stream holds the
 * globals substream followed by one substream per sheet, each framed by
 * BOF ... EOF; worksheets may nest further substreams (embedded charts).
 */
import { BOF_TYPE, RT } from './biff-opcodes';

export interface BiffRecord {
	type: number;
	/** Stream offset of the record header. */
	offset: number;
	/** Record data, with the data of the following CONTINUE records appended. */
	data: Uint8Array;
	/** Offsets in `data` where the data of each CONTINUE record starts. */
	breaks: number[];
}

export interface Substream {
	/** Records of the substream, BOF first; nested substreams are left out. */
	records: BiffRecord[];
	/** BOF types of the nested substreams that were skipped (embedded charts). */
	nested: number[];
	/** The stream ended (or a header was cut) before the substream's EOF. */
	truncated: boolean;
}

/** Records that carry a CONTINUE of their own kind and must not absorb CONTINUE records. */
const NO_MERGE = new Set<number>([RT.EOF, RT.BOF]);

interface RawHeader {
	type: number;
	size: number;
}

function header(stream: Uint8Array, offset: number): RawHeader | undefined {
	if (offset + 4 > stream.length) {
		return undefined;
	}
	return {
		type: stream[offset] | (stream[offset + 1] << 8),
		size: stream[offset + 2] | (stream[offset + 3] << 8),
	};
}

/**
 * The record at `offset` with its CONTINUE records merged. `truncated` is
 * set when the declared size runs past the end of the stream (the data is
 * cut to what is there).
 */
export function readRecord(
	stream: Uint8Array,
	offset: number,
): { record: BiffRecord; next: number; truncated: boolean } | undefined {
	const h = header(stream, offset);
	if (!h) {
		return undefined;
	}
	let truncated = offset + 4 + h.size > stream.length;
	let next = Math.min(stream.length, offset + 4 + h.size);
	const parts: Uint8Array[] = [stream.subarray(offset + 4, next)];
	if (!NO_MERGE.has(h.type)) {
		for (let c = header(stream, next); c && c.type === RT.CONTINUE; c = header(stream, next)) {
			const end = Math.min(stream.length, next + 4 + c.size);
			truncated ||= next + 4 + c.size > stream.length;
			parts.push(stream.subarray(next + 4, end));
			next = end;
		}
	}
	const breaks: number[] = [];
	let data = parts[0];
	if (parts.length > 1) {
		const total = parts.reduce((n, p) => n + p.length, 0);
		data = new Uint8Array(total);
		let at = 0;
		parts.forEach((p, i) => {
			if (i > 0) {
				breaks.push(at);
			}
			data.set(p, at);
			at += p.length;
		});
	}
	return { record: { type: h.type, offset, data, breaks }, next, truncated };
}

/** `vers` and `dt` of a BOF record. */
export function bofInfo(record: BiffRecord): { version: number; type: number } {
	const d = record.data;
	return {
		version: d.length >= 2 ? d[0] | (d[1] << 8) : 0,
		type: d.length >= 4 ? d[2] | (d[3] << 8) : 0,
	};
}

/**
 * The substream starting with the BOF at `start`, up to its EOF. Nested
 * chart substreams are skipped; another kind of BOF at the top level means
 * the EOF is missing and ends the substream.
 */
export function readSubstream(stream: Uint8Array, start: number): Substream {
	const records: BiffRecord[] = [];
	const nested: number[] = [];
	let offset = start;
	let depth = 0;
	let truncated = false;
	// Every record is at least 4 bytes, which bounds the loop.
	while (offset < stream.length) {
		const read = readRecord(stream, offset);
		if (!read) {
			truncated = true;
			break;
		}
		const { record, next } = read;
		truncated ||= read.truncated;
		if (record.type === RT.BOF) {
			if (records.length === 0 && depth === 0) {
				records.push(record);
				depth = 1;
			} else if (depth === 1 && bofInfo(record).type !== BOF_TYPE.chart) {
				// A sheet BOF where an EOF was expected: the previous substream is cut short.
				truncated = true;
				break;
			} else {
				if (depth === 1) {
					nested.push(bofInfo(record).type);
				}
				depth++;
			}
		} else if (record.type === RT.EOF) {
			depth--;
			if (depth <= 0) {
				return { records, nested, truncated };
			}
		} else if (depth === 1) {
			if (record.type !== 0 || record.data.length > 0) {
				records.push(record);
			}
		} else if (depth === 0) {
			// Not a substream at all.
			truncated = true;
			break;
		}
		if (read.truncated) {
			break;
		}
		offset = next;
	}
	return { records, nested, truncated: true };
}
