/**
 * Hyperlinks: HLINK records hold an [MS-OSHARED] hyperlink object (URL or
 * file moniker, location inside the workbook, display text); the tooltip
 * comes from the HLINKTOOLTIP record that follows.
 */
import type { RangeAddress } from '../../types/address';
import type { Hyperlink } from '../../types/sheet';
import { BiffCursor, BiffTruncatedError, decodeChars } from './biff-cursor';
import { RT } from './biff-opcodes';
import type { BiffRecord } from './biff-records';
import { makeRange } from './sheet-context';

const HAS_MONIKER = 0x01;
const HAS_LOCATION = 0x08;
const HAS_DISPLAY = 0x10;
const HAS_GUID = 0x20;
const HAS_TIME = 0x40;
const HAS_FRAME = 0x80;
const MONIKER_AS_STRING = 0x100;

/** CLSIDs in their on-disk byte order. */
const URL_MONIKER = 'e0c9ea79f9bace118c8200aa004ba90b';
const FILE_MONIKER = '0303000000000000c000000000000046';

function hex(bytes: Uint8Array): string {
	return [...bytes].map((b) => b.toString(16).padStart(2, '0')).join('');
}

/** `HyperlinkString`: character count (including the terminating null) and UTF-16 characters. */
function hyperlinkString(cur: BiffCursor): string {
	const count = cur.u32();
	if (count * 2 > cur.remaining) {
		throw new BiffTruncatedError('hyperlink');
	}
	const text = decodeChars(cur.data, cur.pos, count, true);
	cur.skip(count * 2);
	return text.replace(/\u0000+$/, '');
}

function urlMoniker(cur: BiffCursor): string {
	const length = cur.u32();
	const end = cur.pos + Math.min(length, cur.remaining);
	const body = new BiffCursor(cur.data.subarray(0, end), [], cur.pos);
	const url = body.utf16z();
	cur.pos = end;
	return url;
}

function fileMoniker(cur: BiffCursor): string {
	const up = cur.u16();
	const ansiLength = cur.u32();
	const ansi = cur.bytes(ansiLength);
	let path = decodeChars(ansi, 0, ansi.length, false).replace(/\u0000+$/, '');
	cur.skip(2 + 2 + 16 + 4);
	const unicodeSize = cur.u32();
	if (unicodeSize > 0) {
		const bytes = cur.u32();
		cur.skip(2);
		const unicode = cur.bytes(bytes);
		path = decodeChars(unicode, 0, bytes >> 1, true);
	}
	return fileTarget('..\\'.repeat(up) + path);
}

/**
 * Hyperlink target for a file path, as Excel writes it in an `.xlsx`:
 * absolute and UNC paths become `file:///` URLs, relative paths use `/`.
 */
export function fileTarget(path: string): string {
	if (/^[A-Za-z]:[\\/]/.test(path) || path.startsWith('\\\\')) {
		return `file:///${path}`;
	}
	return path.replace(/\\/g, '/');
}

/** Parse an HLINK record; `undefined` when it has nothing to link to. */
function parseHyperlink(record: BiffRecord): Hyperlink | undefined {
	const cur = new BiffCursor(record.data, record.breaks);
	const ref = makeRange(cur.u16(), cur.u16(), cur.u16(), cur.u16());
	cur.skip(16 + 4);
	const flags = cur.u32();
	const link: Partial<Hyperlink> = {};
	if (flags & HAS_DISPLAY) {
		const display = hyperlinkString(cur);
		if (display) {
			link.display = display;
		}
	}
	if (flags & HAS_FRAME) {
		hyperlinkString(cur);
	}
	if (flags & HAS_MONIKER) {
		if (flags & MONIKER_AS_STRING) {
			link.target = hyperlinkString(cur);
		} else {
			const clsid = hex(cur.bytes(16));
			if (clsid === URL_MONIKER) {
				link.target = urlMoniker(cur);
			} else if (clsid === FILE_MONIKER) {
				link.target = fileMoniker(cur);
			} else {
				// Composite and other monikers have no length prefix: the target cannot be read.
				throw new Error('unsupported hyperlink moniker');
			}
		}
	}
	if (flags & HAS_LOCATION) {
		link.location = hyperlinkString(cur);
	}
	if (flags & HAS_GUID) {
		cur.skip(16);
	}
	if (flags & HAS_TIME) {
		cur.skip(8);
	}
	if (!ref || (!link.target && !link.location)) {
		return undefined;
	}
	const out: Hyperlink = { ref };
	if (link.target) {
		out.target = link.target;
	}
	if (link.location) {
		out.location = link.location;
	}
	if (link.display) {
		out.display = link.display;
	}
	return out;
}

/** HLINKTOOLTIP: the cell range and the tooltip text. */
function parseTooltip(record: BiffRecord): { ref?: RangeAddress; text: string } {
	const cur = new BiffCursor(record.data, record.breaks);
	cur.skip(2);
	const ref = makeRange(cur.u16(), cur.u16(), cur.u16(), cur.u16());
	return { ref, text: cur.utf16z() };
}

/** Record types handled by {@link readLink}. */
export const LINK_RECORDS: readonly number[] = [RT.HLINK, RT.HLINKTOOLTIP];

/** Handle hyperlink records; returns `false` for other record types. */
export function readLink(links: Hyperlink[], record: BiffRecord): boolean {
	if (record.type === RT.HLINK) {
		const link = parseHyperlink(record);
		if (link) {
			links.push(link);
		}
		return true;
	}
	if (record.type === RT.HLINKTOOLTIP) {
		const { ref, text } = parseTooltip(record);
		const same = (r: RangeAddress): boolean =>
			!!ref &&
			r.top === ref.top &&
			r.left === ref.left &&
			r.bottom === ref.bottom &&
			r.right === ref.right;
		for (let i = links.length - 1; i >= 0; i--) {
			if (same(links[i].ref)) {
				if (text) {
					links[i].tooltip = text;
				}
				break;
			}
		}
		return true;
	}
	return false;
}
