/**
 * Shared string table for saving. Original `<si>` entries are kept verbatim
 * and in order (cells in untouched sheets keep valid indices); new plain and
 * rich strings are appended and deduplicated.
 */
import type { RichTextRun } from '../types/cell';
import type { FontSpec } from '../types/style';
import { canonicalKey } from '../styles/style-pool';
import {
	decodeOoxmlEscapes,
	elements,
	localName,
	namePrefix,
	parseXml,
	serializeElement,
	textContent,
} from '../xml';
import { NS } from '../xml/namespaces';
import type { XmlElement } from '../xml/types';
import { isEmptyFont, parseFontProps, richRunsXml } from './rich-text';
import { attrString, mergeAttrs, textElement } from './xml-utils';

/** Excel's cell text limit. */
export const MAX_CELL_TEXT = 32767;

function stripFont(font: FontSpec | undefined): FontSpec | undefined {
	if (isEmptyFont(font)) {
		return undefined;
	}
	const out: Record<string, unknown> = {};
	for (const [k, v] of Object.entries(font as FontSpec)) {
		if (v !== undefined) {
			out[k] = v;
		}
	}
	return out as FontSpec;
}

function richKey(runs: readonly RichTextRun[]): string {
	return canonicalKey(runs.map((r) => [r.text, stripFont(r.font) ?? null]));
}

/** The runs carry no formatting at all (they are just a split plain string). */
export function isPlainRuns(runs: readonly RichTextRun[]): boolean {
	return runs.every((r) => isEmptyFont(r.font));
}

function truncate(text: string): string {
	return text.length > MAX_CELL_TEXT ? text.slice(0, MAX_CELL_TEXT) : text;
}

export class SharedStringTable {
	private readonly entries: Array<XmlElement | string> = [];
	private readonly plainIndex = new Map<string, number>();
	private readonly richIndex = new Map<string, number>();
	private readonly originalLength: number;
	private readonly rootName: string;
	private readonly rootAttrs: Record<string, string>;
	private readonly p: string;
	/** Cell references made through this table during the save. */
	refs = 0;
	/** Strings that had to be cut to Excel's 32767-character limit. */
	truncated = 0;

	private constructor(root?: XmlElement) {
		this.rootName = root?.name ?? 'sst';
		this.rootAttrs = root ? { ...root.attrs } : { xmlns: NS.main };
		const prefix = root ? namePrefix(root.name) : '';
		this.p = prefix ? `${prefix}:` : '';
		for (const si of elements(root)) {
			if (localName(si.name) !== 'si') {
				continue;
			}
			const index = this.entries.length;
			this.entries.push(si);
			this.indexOriginal(si, index);
		}
		this.originalLength = this.entries.length;
	}

	/** An empty table (new workbook, or a source without shared strings). */
	static empty(): SharedStringTable {
		return new SharedStringTable();
	}

	/** Seed from the original `sharedStrings.xml`. */
	static fromXml(xml: string): SharedStringTable {
		return new SharedStringTable(parseXml(xml).root);
	}

	private indexOriginal(si: XmlElement, index: number): void {
		const kids = elements(si);
		const runs = kids.filter((k) => localName(k.name) === 'r');
		if (runs.length === 0) {
			const t = kids.find((k) => localName(k.name) === 't');
			const text = decodeOoxmlEscapes(textContent(t));
			if (!this.plainIndex.has(text)) {
				this.plainIndex.set(text, index);
			}
			return;
		}
		const parsed: RichTextRun[] = runs.map((r) => {
			const rPr = elements(r).find((k) => localName(k.name) === 'rPr');
			const t = elements(r).find((k) => localName(k.name) === 't');
			const font = rPr ? parseFontProps(rPr) : undefined;
			return { text: decodeOoxmlEscapes(textContent(t)), ...(font ? { font } : {}) };
		});
		const key = richKey(parsed);
		if (!this.richIndex.has(key)) {
			this.richIndex.set(key, index);
		}
	}

	/** Number of entries (`uniqueCount`). */
	get size(): number {
		return this.entries.length;
	}

	/** Entries were appended to the original table. */
	get grew(): boolean {
		return this.entries.length > this.originalLength;
	}

	/** Index of a plain string, appending it when new. */
	internPlain(text: string): number {
		this.refs++;
		const hit = this.plainIndex.get(text);
		if (hit !== undefined) {
			return hit;
		}
		const index = this.entries.length;
		if (text.length > MAX_CELL_TEXT) {
			this.truncated++;
		}
		this.entries.push(`<${this.p}si>${textElement(truncate(text), this.p)}</${this.p}si>`);
		this.plainIndex.set(text, index);
		return index;
	}

	/** Index of a rich string (runs without formatting are stored as plain text). */
	internRich(runs: readonly RichTextRun[]): number {
		if (isPlainRuns(runs)) {
			return this.internPlain(runs.map((r) => r.text).join(''));
		}
		this.refs++;
		const key = richKey(runs);
		const hit = this.richIndex.get(key);
		if (hit !== undefined) {
			return hit;
		}
		const index = this.entries.length;
		let remaining = MAX_CELL_TEXT;
		const clipped: RichTextRun[] = [];
		for (const r of runs) {
			if (remaining <= 0) {
				this.truncated++;
				break;
			}
			const text = r.text.length > remaining ? r.text.slice(0, remaining) : r.text;
			remaining -= text.length;
			clipped.push({ ...r, text });
		}
		this.entries.push(`<${this.p}si>${richRunsXml(clipped, this.p)}</${this.p}si>`);
		this.richIndex.set(key, index);
		return index;
	}

	/**
	 * Serialize the part. `count` (total references) is only written when it
	 * is known exactly, i.e. every sheet referencing the table was written.
	 */
	serialize(exactCount: boolean): string {
		const attrs = mergeAttrs(this.rootAttrs, {
			count: exactCount ? this.refs : undefined,
			uniqueCount: this.entries.length,
		});
		const parts: string[] = [
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n',
			`<${this.rootName}${attrString(attrs)}>`,
		];
		for (const e of this.entries) {
			parts.push(typeof e === 'string' ? e : serializeElement(e));
		}
		parts.push(`</${this.rootName}>`);
		return parts.join('');
	}
}
