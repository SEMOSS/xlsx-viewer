import { formatRange } from '../address/range';
import { quoteSheetName } from '../address/sheet-name';
import { keyCol, keyRow } from '../address/cell-address';
import type { StructuredRefNode } from '../formula/ast';
import { structuredRange, tableDataRange } from '../formula/evaluator/structured';
import { parseFormula } from '../formula/language/parser';
import { tokenizeFormula } from '../formula/language/tokenizer';
import type { RangeAddress } from '../types/address';
import type { Cell } from '../types/cell';
import type { Workbook } from '../types/workbook';
import { bakeTableStyle } from './table-bake';
import { replaceTable, tableOf } from './table-internal';
import { TableError } from './table-ops';
import type { EditTransaction } from './transaction';

/** Table names: a letter, `_` or `\` first, then letters, digits, `.` and `_`; not a cell reference. */
export function isValidTableName(name: string): boolean {
	if (name.length === 0 || name.length > 255 || /^[CcRr]$/.test(name)) {
		return false;
	}
	if (!/^[A-Za-z_\\\u00C0-\uFFFF][A-Za-z0-9_.\\\u00C0-\uFFFF]*$/.test(name)) {
		return false;
	}
	return !/^[A-Za-z]{1,3}\d+$/.test(name) && !/^[Rr]\d*[Cc]\d*$/.test(name);
}

/** Rewrite every formula of the workbook token by token (cells and defined names). */
function rewriteFormulas(
	tx: EditTransaction,
	workbook: Workbook,
	rewrite: (formula: string, sheetIndex: number, row: number, col: number) => string,
): void {
	workbook.sheets.forEach((sheet, si) => {
		const changes: Array<[number, Cell]> = [];
		for (const [key, cell] of sheet.cells) {
			if (cell.formula) {
				const next = rewrite(cell.formula, si, keyRow(key), keyCol(key));
				if (next !== cell.formula) {
					changes.push([key, { ...cell, formula: next }]);
				}
			}
		}
		for (const [key, cell] of changes) {
			tx.touchCell(si, keyRow(key), keyCol(key));
			sheet.cells.set(key, cell);
		}
	});
	const names = workbook.definedNames.map((n) => rewrite(n.formula, n.localSheet ?? -1, -1, -1));
	if (names.some((f, i) => f !== workbook.definedNames[i].formula)) {
		tx.touchWorkbook();
		workbook.definedNames.forEach((n, i) => {
			n.formula = names[i];
		});
	}
}

function mapTokens(formula: string, map: (type: string, text: string) => string): string {
	let changed = false;
	const out = tokenizeFormula(formula)
		.map((t) => {
			const next = t.type === 'structured' || t.type === 'name' ? map(t.type, t.text) : t.text;
			changed ||= next !== t.text;
			return next;
		})
		.join('');
	return changed ? out : formula;
}

/** Table Design > Table Name: rename a table and the structured references to it. */
export function renameTable(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	index: number,
	name: string,
): void {
	const table = tableOf(workbook, sheetIndex, index);
	const next = name.trim();
	if (next === table.name) {
		return;
	}
	if (!isValidTableName(next)) {
		throw new TableError(`"${next}" isn't a valid table name.`, 'invalid-name');
	}
	const lower = next.toLowerCase();
	const taken =
		workbook.sheets.some((s) =>
			s.tables.some((t) => t !== table && t.name.toLowerCase() === lower),
		) || workbook.definedNames.some((n) => n.name.toLowerCase() === lower);
	if (taken) {
		throw new TableError(`The name "${next}" is already in use.`, 'duplicate');
	}
	const old = table.name.toLowerCase();
	rewriteFormulas(tx, workbook, (f) =>
		mapTokens(f, (type, text) => {
			if (type === 'name') {
				return text.toLowerCase() === old ? next : text;
			}
			return text.toLowerCase().startsWith(`${old}[`) ? next + text.slice(old.length) : text;
		}),
	);
	replaceTable(tx, workbook, sheetIndex, index, { ...table, name: next, displayName: next });
}

/** A1 text of a range seen from a formula on `fromSheet` (absolute, sheet-qualified across sheets). */
function absoluteRef(
	workbook: Workbook,
	sheetIndex: number,
	fromSheet: number,
	range: RangeAddress,
): string {
	const text = formatRange(range, { absolute: true });
	return fromSheet === sheetIndex
		? text
		: `${quoteSheetName(workbook.sheets[sheetIndex].name)}!${text}`;
}

/**
 * Table Design > Convert to Range: the table becomes plain cells. Its
 * style is applied as cell formatting and structured references to it are
 * replaced by absolute A1 references, as Excel does.
 */
export function convertTableToRange(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	index: number,
): void {
	const table = tableOf(workbook, sheetIndex, index);
	const lower = table.name.toLowerCase();
	rewriteFormulas(tx, workbook, (f, fs, row, col) =>
		mapTokens(f, (type, text) => {
			if (type === 'name') {
				const range = text.toLowerCase() === lower ? tableDataRange(table) : undefined;
				return range && !('code' in range) ? absoluteRef(workbook, sheetIndex, fs, range) : text;
			}
			let node: StructuredRefNode;
			try {
				const parsed = parseFormula(text);
				if (parsed.type !== 'structured') {
					return text;
				}
				node = parsed;
			} catch {
				return text;
			}
			const inside =
				fs === sheetIndex &&
				row >= table.ref.top &&
				row <= table.ref.bottom &&
				col >= table.ref.left &&
				col <= table.ref.right;
			if (node.table ? node.table.toLowerCase() !== lower : !inside) {
				return text;
			}
			const range = structuredRange(node, table, row);
			return 'code' in range ? range.code : absoluteRef(workbook, sheetIndex, fs, range);
		}),
	);
	bakeTableStyle(tx, workbook, sheetIndex, table);
	tx.touchSheet(sheetIndex);
	const sheet = workbook.sheets[sheetIndex];
	sheet.tables = sheet.tables.filter((_, i) => i !== index);
	sheet.dirty = true;
}
