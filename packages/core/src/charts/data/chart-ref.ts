/**
 * Chart series formulas: `Sheet1!$B$2:$B$6`, `'My Sheet'!$A$2:$A$6`,
 * multi-area unions `(Sheet1!$A$1,Sheet1!$A$3)` and defined names.
 */
import { formatRange, parseRange } from '../../address/range';
import { quoteSheetName, unquoteSheetName } from '../../address/sheet-name';
import type { RangeAddress } from '../../types/address';
import type { Workbook } from '../../types/workbook';

/** One rectangular block of cells on a sheet. */
export interface ChartRefArea {
	sheetIndex: number;
	range: RangeAddress;
}

const MAX_NAME_DEPTH = 8;

/** Split on commas that are outside quotes and parentheses; `null` for unbalanced input. */
function splitTopLevel(text: string): string[] | null {
	const parts: string[] = [];
	let depth = 0;
	let quoted = false;
	let start = 0;
	for (let i = 0; i < text.length; i++) {
		const ch = text[i];
		if (ch === "'") {
			quoted = !quoted;
		} else if (!quoted && ch === '(') {
			depth++;
		} else if (!quoted && ch === ')') {
			depth--;
			if (depth < 0) {
				return null;
			}
		} else if (!quoted && depth === 0 && ch === ',') {
			parts.push(text.slice(start, i));
			start = i + 1;
		}
	}
	if (depth !== 0 || quoted) {
		return null;
	}
	parts.push(text.slice(start));
	return parts.map((p) => p.trim()).filter((p) => p.length > 0);
}

/** Strip one pair of parentheses that wraps the whole text. */
function stripOuterParens(text: string): string {
	let t = text.trim();
	while (t.startsWith('(') && t.endsWith(')')) {
		let depth = 0;
		let wraps = true;
		for (let i = 0; i < t.length; i++) {
			if (t[i] === '(') {
				depth++;
			} else if (t[i] === ')') {
				depth--;
				if (depth === 0 && i < t.length - 1) {
					wraps = false;
					break;
				}
			}
		}
		if (!wraps) {
			break;
		}
		t = t.slice(1, -1).trim();
	}
	return t;
}

/** Index of the sheet separator `!` outside quotes, or -1. */
function bangIndex(text: string): number {
	let quoted = false;
	let found = -1;
	for (let i = 0; i < text.length; i++) {
		if (text[i] === "'") {
			quoted = !quoted;
		} else if (!quoted && text[i] === '!') {
			found = i;
		}
	}
	return found;
}

function sheetIndexByName(workbook: Workbook, name: string): number {
	const lower = name.toLowerCase();
	return workbook.sheets.findIndex((s) => s.name.toLowerCase() === lower);
}

function findName(workbook: Workbook, name: string, scope: number | undefined) {
	const lower = name.toLowerCase();
	return workbook.definedNames.find(
		(n) =>
			n.name.toLowerCase() === lower &&
			(scope === undefined ? n.localSheet === undefined : n.localSheet === scope),
	);
}

function resolveName(
	workbook: Workbook,
	name: string,
	scope: number | undefined,
	contextSheet: number,
	depth: number,
): ChartRefArea[] | null {
	if (depth >= MAX_NAME_DEPTH || !/^[A-Za-z_\\\u00C0-\uFFFF][\w.\\\u00C0-\uFFFF]*$/.test(name)) {
		return null;
	}
	// `Sheet1!Name` prefers the sheet-level name and, like Excel, falls back to the workbook-level one.
	const def =
		findName(workbook, name, scope ?? contextSheet) ?? findName(workbook, name, undefined);
	if (!def) {
		return null;
	}
	const home = def.localSheet ?? contextSheet;
	return resolveChartRef(def.formula, workbook, home, depth + 1);
}

/**
 * Resolve a chart reference to cell areas. Returns `null` when any part cannot
 * be resolved (external workbooks, 3-D references, functions, unknown names);
 * callers then fall back to the values cached in the chart part.
 */
export function resolveChartRef(
	ref: string,
	workbook: Workbook,
	sheetIndex: number,
	depth = 0,
): ChartRefArea[] | null {
	const text = stripOuterParens(ref.trim().replace(/^=/, ''));
	if (!text) {
		return null;
	}
	const parts = splitTopLevel(text);
	if (!parts || parts.length === 0) {
		return null;
	}
	const out: ChartRefArea[] = [];
	for (const part of parts) {
		const areas = resolvePart(stripOuterParens(part), workbook, sheetIndex, depth);
		if (!areas) {
			return null;
		}
		out.push(...areas);
	}
	return out;
}

function resolvePart(
	part: string,
	workbook: Workbook,
	sheetIndex: number,
	depth: number,
): ChartRefArea[] | null {
	const bang = bangIndex(part);
	if (bang < 0) {
		const range = parseRange(part);
		if (range) {
			return [{ sheetIndex, range }];
		}
		return resolveName(workbook, part, undefined, sheetIndex, depth);
	}
	const sheetText = part.slice(0, bang).trim();
	const refText = part.slice(bang + 1).trim();
	if (sheetText === '[0]') {
		return resolveName(workbook, refText, undefined, sheetIndex, depth);
	}
	const sheetName = unquoteSheetName(sheetText);
	// External workbooks (`[1]Sheet1`) and 3-D references (`Sheet1:Sheet3`) cannot be read.
	if (/^\[.*\]/.test(sheetName) || (!sheetText.startsWith("'") && sheetName.includes(':'))) {
		return null;
	}
	const target = sheetIndexByName(workbook, sheetName);
	if (target < 0) {
		return null;
	}
	const range = parseRange(refText);
	if (range) {
		return [{ sheetIndex: target, range }];
	}
	return resolveName(workbook, refText, target, target, depth);
}

/** Format an absolute reference to a range, e.g. `'My Sheet'!$A$2:$A$6`. */
export function formatChartRef(sheetName: string, range: RangeAddress): string {
	return `${quoteSheetName(sheetName)}!${formatRange(range, { absolute: true, forceCells: true })}`;
}
