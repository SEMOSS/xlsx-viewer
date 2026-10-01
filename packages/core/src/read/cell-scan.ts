import { MAX_COLS, MAX_ROWS } from '../address/constants';
import type { RowInfo } from '../types/sheet';
import { parseTagAttributes, tagEnd } from './xml-normalize';

/**
 * Character-level helpers for the `<sheetData>` scanner. They read straight
 * from the document string and write into reusable records, so scanning a
 * cell allocates nothing but the cell itself.
 */

export const T_N = 0;
export const T_S = 1;
export const T_STR = 2;
export const T_B = 3;
export const T_E = 4;
export const T_INLINE = 5;
export const T_D = 6;

export function isWs(c: number): boolean {
	return c === 32 || c === 10 || c === 13 || c === 9;
}

/** Character that ends an element name (`>`, `/` or whitespace). */
export function isNameEnd(c: number): boolean {
	return c === 32 || c === 62 || c === 47 || c === 10 || c === 13 || c === 9;
}

/** Unsigned integer in `xml[a, b)`, or -1. */
export function parseUint(xml: string, a: number, b: number): number {
	let n = 0;
	let digits = 0;
	for (let i = a; i < b; i++) {
		const c = xml.charCodeAt(i);
		if (c >= 48 && c <= 57) {
			n = n * 10 + (c - 48);
			digits++;
		} else if (!isWs(c)) {
			return -1;
		}
	}
	return digits > 0 ? n : -1;
}

/** `1` / `true` (any case) in `xml[a, b)`. */
export function isTrue(xml: string, a: number, b: number): boolean {
	const c = xml.charCodeAt(a);
	return b > a && (c === 49 || c === 116 || c === 84);
}

function cellType(xml: string, a: number, b: number): number {
	const len = b - a;
	const c = xml.charCodeAt(a);
	if (len === 1) {
		switch (c) {
			case 115:
				return T_S;
			case 98:
				return T_B;
			case 101:
				return T_E;
			case 100:
				return T_D;
			default:
				return T_N;
		}
	}
	if (len === 3 && c === 115) {
		return T_STR;
	}
	if (len === 9 && c === 105) {
		return T_INLINE;
	}
	return T_N;
}

/** Attributes of one `<c>` start tag. */
export interface CellTag {
	/** 0-based position from `r`, or -1 when absent / invalid. */
	row: number;
	col: number;
	style: number;
	type: number;
	cm: number;
	vm: number;
	ph: boolean;
	selfClosing: boolean;
}

export function newCellTag(): CellTag {
	return { row: -1, col: -1, style: 0, type: T_N, cm: -1, vm: -1, ph: false, selfClosing: false };
}

/**
 * Read the attributes of a `<c>` tag starting at `k` (just past the element
 * name) into `out`. Returns the offset after the tag.
 */
export function scanCellTag(xml: string, k: number, end: number, out: CellTag): number {
	out.row = -1;
	out.col = -1;
	out.style = 0;
	out.type = T_N;
	out.cm = -1;
	out.vm = -1;
	out.ph = false;
	out.selfClosing = false;
	for (;;) {
		let ch = xml.charCodeAt(k);
		while (isWs(ch)) {
			ch = xml.charCodeAt(++k);
		}
		if (ch === 62 /* > */) {
			return k + 1;
		}
		if (ch === 47 /* / */) {
			out.selfClosing = true;
			return k + 2;
		}
		if (k >= end) {
			throw new Error('Unterminated <c> tag');
		}
		const nameStart = k;
		while (ch !== 61 /* = */ && k < end) {
			ch = xml.charCodeAt(++k);
		}
		let nameEnd = k;
		while (nameEnd > nameStart && isWs(xml.charCodeAt(nameEnd - 1))) {
			nameEnd--;
		}
		ch = xml.charCodeAt(++k);
		while (isWs(ch)) {
			ch = xml.charCodeAt(++k);
		}
		const valStart = k + 1;
		const valEnd = xml.indexOf(ch === 39 ? "'" : '"', valStart);
		if (valEnd === -1) {
			throw new Error('Unterminated attribute in <c>');
		}
		k = valEnd + 1;
		const nameLen = nameEnd - nameStart;
		const n0 = xml.charCodeAt(nameStart);
		if (nameLen === 1) {
			if (n0 === 114 /* r */) {
				readRef(xml, valStart, valEnd, out);
			} else if (n0 === 115 /* s */) {
				out.style = Math.max(0, parseUint(xml, valStart, valEnd));
			} else if (n0 === 116 /* t */) {
				out.type = cellType(xml, valStart, valEnd);
			}
		} else if (nameLen === 2) {
			const n1 = xml.charCodeAt(nameStart + 1);
			if (n1 === 109 /* m */ && n0 === 99 /* c */) {
				out.cm = parseUint(xml, valStart, valEnd);
			} else if (n1 === 109 && n0 === 118 /* v */) {
				out.vm = parseUint(xml, valStart, valEnd);
			} else if (n0 === 112 /* p */ && n1 === 104 /* h */) {
				out.ph = isTrue(xml, valStart, valEnd);
			}
		}
	}
}

/** A1 reference (`$` ignored) into `out.row` / `out.col`; leaves -1 when invalid. */
function readRef(xml: string, a: number, b: number, out: CellTag): void {
	let c = 0;
	let r = 0;
	let j = a;
	for (; j < b; j++) {
		const x = xml.charCodeAt(j);
		if (x >= 65 && x <= 90) {
			c = c * 26 + (x - 64);
		} else if (x >= 97 && x <= 122) {
			c = c * 26 + (x - 96);
		} else if (x !== 36) {
			break;
		}
	}
	for (; j < b; j++) {
		const x = xml.charCodeAt(j);
		if (x >= 48 && x <= 57) {
			r = r * 10 + (x - 48);
		} else if (x !== 36) {
			break;
		}
	}
	if (j === b && c >= 1 && c <= MAX_COLS && r >= 1 && r <= MAX_ROWS) {
		out.row = r - 1;
		out.col = c - 1;
	}
}

/** Row attributes worth keeping (`undefined` when the row only has `r` / `spans`). */
export function rowInfo(attrs: Record<string, string>): RowInfo | undefined {
	const info: RowInfo = {};
	let any = false;
	const flag = (v: string | undefined): boolean => v === '1' || v === 'true';
	if (attrs.ht !== undefined) {
		const h = Number(attrs.ht);
		if (Number.isFinite(h)) {
			info.height = h;
			any = true;
		}
	}
	if (flag(attrs.customHeight)) {
		info.customHeight = any = true;
	}
	if (flag(attrs.hidden)) {
		info.hidden = any = true;
	}
	if (attrs.s !== undefined) {
		const s = Number(attrs.s);
		if (Number.isInteger(s) && s >= 0) {
			info.style = s;
			any = true;
		}
	}
	if (flag(attrs.customFormat)) {
		info.customFormat = any = true;
	}
	const level = Number(attrs.outlineLevel);
	if (level > 0) {
		info.outlineLevel = level;
		any = true;
	}
	if (flag(attrs.collapsed)) {
		info.collapsed = any = true;
	}
	if (flag(attrs.thickTop)) {
		info.thickTop = any = true;
	}
	if (flag(attrs.thickBot)) {
		info.thickBot = any = true;
	}
	return any ? info : undefined;
}

/** Attributes of a `<row>` tag at `lt`; `attrStart` is just past the name. */
export function scanRowTag(
	xml: string,
	lt: number,
	attrStart: number,
): { attrs: Record<string, string>; end: number } {
	const gt = tagEnd(xml, lt);
	if (gt === -1) {
		throw new Error('Unterminated <row> tag');
	}
	const selfClosing = xml.charCodeAt(gt - 2) === 47;
	return {
		attrs: parseTagAttributes(xml.slice(attrStart, selfClosing ? gt - 2 : gt - 1)),
		end: gt,
	};
}

/** Skip a comment, CDATA section or processing instruction at `lt`. */
export function skipSpecial(xml: string, lt: number, end: number): number {
	let e: number;
	if (xml.startsWith('<!--', lt)) {
		e = xml.indexOf('-->', lt + 4);
		return e === -1 ? end : e + 3;
	}
	if (xml.startsWith('<![CDATA[', lt)) {
		e = xml.indexOf(']]>', lt + 9);
		return e === -1 ? end : e + 3;
	}
	e = xml.indexOf('>', lt + 1);
	return e === -1 ? end : e + 1;
}

/** Skip an element we do not read (`extLst`, foreign markup) including its content. */
export function skipElement(xml: string, lt: number, end: number): number {
	const gt = tagEnd(xml, lt);
	if (gt === -1) {
		return end;
	}
	if (xml.charCodeAt(gt - 2) === 47) {
		return gt;
	}
	let q = lt + 1;
	while (q < gt && !isNameEnd(xml.charCodeAt(q))) {
		q++;
	}
	const close = `</${xml.slice(lt + 1, q)}>`;
	const e = xml.indexOf(close, gt);
	return e === -1 ? end : e + close.length;
}
