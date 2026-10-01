import type { CellAddress, Workbook } from 'xlsx-viewer-core';
import {
	formatValue,
	isCellError,
	keyCol,
	keyRow,
	StyleResolver,
	wildcardToRegExp,
} from 'xlsx-viewer-core';

export interface FindOptions {
	query: string;
	matchCase?: boolean;
	/** Match the entire cell contents. */
	wholeCell?: boolean;
	/** Search formulas (default) or displayed values. */
	lookIn?: 'formulas' | 'values';
	/** Search the active sheet or the whole workbook. */
	within?: 'sheet' | 'workbook';
	/** Search by rows (default) or by columns. */
	order?: 'rows' | 'columns';
}

export interface FindHit extends CellAddress {
	sheet: number;
	text: string;
}

function cellText(
	workbook: Workbook,
	styles: StyleResolver,
	sheet: number,
	key: number,
	lookIn: 'formulas' | 'values',
): string {
	const cell = workbook.sheets[sheet].cells.get(key);
	if (!cell) {
		return '';
	}
	if (lookIn === 'formulas' && cell.formula !== undefined) {
		return `=${cell.formula}`;
	}
	const v = cell.value;
	if (v === null) {
		return '';
	}
	if (typeof v === 'boolean') {
		return v ? 'TRUE' : 'FALSE';
	}
	if (isCellError(v)) {
		return v.code;
	}
	if (lookIn === 'formulas' && typeof v === 'number') {
		return String(v);
	}
	return formatValue(v, styles.resolve(cell.style).numFmt, { date1904: workbook.date1904 }).text;
}

/** Build a matcher honouring Excel's wildcards (`*`, `?`, `~`) and options. */
export function makeMatcher(opts: FindOptions): (text: string) => boolean {
	const q = opts.query;
	const hasWild = /[*?~]/.test(q);
	if (opts.wholeCell) {
		const re = wildcardToRegExp(q);
		const flags = opts.matchCase ? '' : 'i';
		const exact = new RegExp(re.source, flags);
		return (text) => exact.test(text);
	}
	if (hasWild) {
		const inner = wildcardToRegExp(q).source.replace(/^\^/, '').replace(/\$$/, '');
		const re = new RegExp(inner, opts.matchCase ? '' : 'i');
		return (text) => re.test(text);
	}
	if (opts.matchCase) {
		return (text) => text.includes(q);
	}
	const lower = q.toLowerCase();
	return (text) => text.toLowerCase().includes(lower);
}

/** All matching cells, in search order. */
export function findAll(workbook: Workbook, activeSheet: number, opts: FindOptions): FindHit[] {
	if (!opts.query) {
		return [];
	}
	const match = makeMatcher(opts);
	const styles = new StyleResolver(workbook.styles, workbook.theme);
	const sheets = opts.within === 'workbook' ? workbook.sheets.map((_, i) => i) : [activeSheet];
	const hits: FindHit[] = [];
	for (const s of sheets) {
		const keys = [...workbook.sheets[s].cells.keys()];
		if (opts.order === 'columns') {
			keys.sort((a, b) => keyCol(a) - keyCol(b) || keyRow(a) - keyRow(b));
		} else {
			keys.sort((a, b) => a - b);
		}
		for (const key of keys) {
			const text = cellText(workbook, styles, s, key, opts.lookIn ?? 'formulas');
			if (text && match(text)) {
				hits.push({ sheet: s, row: keyRow(key), col: keyCol(key), text });
			}
		}
	}
	return hits;
}

/** Replace occurrences in a text (wildcards supported, like Excel). */
export function replaceInText(text: string, opts: FindOptions, replacement: string): string {
	if (opts.wholeCell) {
		return replacement;
	}
	const q = opts.query;
	const source = /[*?~]/.test(q)
		? wildcardToRegExp(q)
				.source.replace(/^\^/, '')
				.replace(/\$$/, '')
				.replace(/\[\\s\\S\]\*/g, '[\\s\\S]*?')
		: q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
	return text.replace(new RegExp(source, opts.matchCase ? 'g' : 'gi'), () => replacement);
}
