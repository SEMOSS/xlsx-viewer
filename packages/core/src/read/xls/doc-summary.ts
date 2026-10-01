/**
 * Document properties from the OLE property-set streams
 * (`\u0005SummaryInformation`, `\u0005DocumentSummaryInformation`,
 * [MS-OLEPS]): title, subject, author, keywords, comments, last author,
 * creation / save times, category, manager and company. Anything
 * unreadable is skipped; properties are informational only.
 */
import type { DocumentProperties } from '../../types/workbook';
import { BiffCursor, decodeChars } from './biff-cursor';

const VT_I2 = 0x02;
const VT_LPSTR = 0x1e;
const VT_LPWSTR = 0x1f;
const VT_FILETIME = 0x40;

/** FILETIME epoch (1601-01-01) to the Unix epoch, in milliseconds. */
const FILETIME_UNIX_MS = 11644473600000;

const CODEPAGES: Readonly<Record<number, string>> = {
	874: 'windows-874',
	932: 'shift_jis',
	936: 'gbk',
	949: 'euc-kr',
	950: 'big5',
	1200: 'utf-16le',
	10000: 'macintosh',
	65001: 'utf-8',
};

type PropertyValue = string | number | Date;

function decodeText(bytes: Uint8Array, codepage: number): string {
	const label =
		CODEPAGES[codepage] ?? (codepage >= 1250 && codepage <= 1258 ? `windows-${codepage}` : '');
	if (label) {
		try {
			return new TextDecoder(label).decode(bytes);
		} catch {
			// Unknown label in this runtime: fall back to Latin-1 below.
		}
	}
	return decodeChars(bytes, 0, bytes.length, false);
}

/** Properties of the first property set in a stream, by property id. */
function readPropertySet(data: Uint8Array): Map<number, PropertyValue> {
	const out = new Map<number, PropertyValue>();
	const cur = new BiffCursor(data);
	if (cur.u16() !== 0xfffe) {
		return out;
	}
	cur.skip(2 + 4 + 16 + 4 + 16);
	const base = cur.u32();
	cur.pos = base;
	cur.skip(4);
	const count = Math.min(cur.u32(), Math.floor(cur.remaining / 8), 1000);
	const entries: Array<[number, number]> = [];
	for (let i = 0; i < count; i++) {
		entries.push([cur.u32(), cur.u32()]);
	}
	let codepage = 1252;
	const codepageEntry = entries.find(([id]) => id === 1);
	if (codepageEntry) {
		cur.pos = base + codepageEntry[1];
		if (cur.u16() === VT_I2) {
			cur.skip(2);
			codepage = cur.u16();
		}
	}
	for (const [id, offset] of entries) {
		try {
			cur.pos = base + offset;
			const type = cur.u16();
			cur.skip(2);
			if (type === VT_LPSTR) {
				const size = cur.u32();
				const text = decodeText(cur.bytes(size), codepage);
				out.set(id, text.replace(/\u0000+$/, ''));
			} else if (type === VT_LPWSTR) {
				const length = cur.u32();
				out.set(id, decodeChars(cur.bytes(length * 2), 0, length, true).replace(/\u0000+$/, ''));
			} else if (type === VT_FILETIME) {
				const low = cur.u32();
				const high = cur.u32();
				const ms = (high * 2 ** 32 + low) / 10000 - FILETIME_UNIX_MS;
				if (high > 0 && Number.isFinite(ms)) {
					out.set(id, new Date(ms));
				}
			}
		} catch {
			// A damaged property: skip it.
		}
	}
	return out;
}

function iso(d: PropertyValue | undefined): string | undefined {
	return d instanceof Date && !Number.isNaN(d.getTime())
		? d.toISOString().replace(/\.\d{3}Z$/, 'Z')
		: undefined;
}

function text(v: PropertyValue | undefined): string | undefined {
	return typeof v === 'string' && v.trim() !== '' ? v : undefined;
}

/** Read the document properties; `stream` looks up compound-file streams by path. */
export function readDocumentSummary(
	stream: (path: string) => Uint8Array | undefined,
): DocumentProperties {
	const props: DocumentProperties = {};
	const set = <K extends keyof DocumentProperties>(key: K, value: DocumentProperties[K]): void => {
		if (value !== undefined) {
			props[key] = value;
		}
	};
	try {
		const summary = stream('\u0005SummaryInformation');
		if (summary) {
			const p = readPropertySet(summary);
			set('title', text(p.get(2)));
			set('subject', text(p.get(3)));
			set('creator', text(p.get(4)));
			set('keywords', text(p.get(5)));
			set('description', text(p.get(6)));
			set('lastModifiedBy', text(p.get(8)));
			set('created', iso(p.get(12)));
			set('modified', iso(p.get(13)));
		}
		const doc = stream('\u0005DocumentSummaryInformation');
		if (doc) {
			const p = readPropertySet(doc);
			set('category', text(p.get(2)));
			set('manager', text(p.get(14)));
			set('company', text(p.get(15)));
		}
	} catch {
		// Damaged property streams: keep what was read.
	}
	return props;
}
