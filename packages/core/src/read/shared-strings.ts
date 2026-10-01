import type { RichTextRun } from '../types/cell';
import { emptySharedStrings, type SharedStrings } from './context';
import { readStringItem } from './rich-text';
import { scanRootTag, spreadsheetPrefix, tagEnd } from './xml-normalize';

/**
 * `xl/sharedStrings.xml` -> {@link SharedStrings}. Shared-string tables can
 * hold hundreds of thousands of items, so `<si>` elements are located with a
 * string scanner; only formatted items are handed to the XML parser.
 */
export function parseSharedStrings(xml: string): SharedStrings {
	const root = scanRootTag(xml);
	if (!root) {
		return emptySharedStrings();
	}
	const prefix = spreadsheetPrefix(root.attrs);
	const p = prefix ? `${prefix}:` : '';
	const open = `<${p}si`;
	const close = `</${p}si>`;
	const strings: string[] = [];
	const rich = new Map<number, RichTextRun[]>();
	let pos = root.end;
	for (;;) {
		const lt = xml.indexOf(open, pos);
		if (lt === -1) {
			break;
		}
		const after = xml.charCodeAt(lt + open.length);
		let gt: number;
		if (after === 62 /* > */) {
			gt = lt + open.length + 1;
		} else if (after === 47 || after === 32 || after === 9 || after === 10 || after === 13) {
			gt = tagEnd(xml, lt);
			if (gt === -1) {
				throw new Error('Unterminated <si> tag');
			}
			if (xml.charCodeAt(gt - 2) === 47 /* / */) {
				strings.push('');
				pos = gt;
				continue;
			}
		} else {
			// Some other element whose name starts with "si".
			pos = lt + open.length;
			continue;
		}
		const end = xml.indexOf(close, gt);
		if (end === -1) {
			throw new Error('Unterminated <si> element');
		}
		const item = readStringItem(xml, lt, gt, end, end + close.length, prefix);
		if (item.runs) {
			rich.set(strings.length, item.runs);
		}
		strings.push(item.text);
		pos = end + close.length;
	}
	return { strings, rich };
}
