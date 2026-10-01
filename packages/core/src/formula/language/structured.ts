import type { StructuredRefNode, StructuredSpecifier } from '../ast';

/**
 * Structured (table) references: `Table1[Col]`, `Table1[#All]`,
 * `Table1[[#Headers],[A]:[C]]`, `[@Qty]`, `Table1[@[Unit Price]]`.
 * Inside brackets `'` escapes the next character (`[`, `]`, `#`, `'`).
 */

const SPECIFIERS: ReadonlyMap<string, StructuredSpecifier> = new Map([
	['#ALL', '#All'],
	['#DATA', '#Data'],
	['#HEADERS', '#Headers'],
	['#TOTALS', '#Totals'],
	['#THIS ROW', '#This Row'],
]);

function unescapeColumn(text: string): string {
	return text.indexOf("'") < 0 ? text : text.replace(/'(.)/g, '$1');
}

/** Escape `[`, `]`, `#` and `'` in a column name. */
export function escapeStructuredColumn(name: string): string {
	return name.replace(/[[\]#']/g, "'$&");
}

/** Index of the `]` closing the bracket item at `open` (honouring `'` escapes), or -1. */
function closingBracket(text: string, open: number): number {
	for (let i = open + 1; i < text.length; i++) {
		const c = text[i];
		if (c === "'") {
			i++;
		} else if (c === ']') {
			return i;
		} else if (c === '[') {
			return -1;
		}
	}
	return -1;
}

function hasUnescapedBracket(inner: string): boolean {
	for (let i = 0; i < inner.length; i++) {
		const c = inner[i];
		if (c === "'") {
			i++;
		} else if (c === '[') {
			return true;
		}
	}
	return false;
}

interface Parts {
	specifiers: StructuredSpecifier[];
	columnStart?: string;
	columnEnd?: string;
}

function parseSimple(inner: string): Parts | null {
	if (inner === '') {
		return { specifiers: [] };
	}
	if (inner[0] === '#') {
		const spec = SPECIFIERS.get(inner.toUpperCase());
		return spec ? { specifiers: [spec] } : null;
	}
	if (inner[0] === '@') {
		const column = unescapeColumn(inner.slice(1));
		return column === ''
			? { specifiers: ['#This Row'] }
			: { specifiers: ['#This Row'], columnStart: column };
	}
	return { specifiers: [], columnStart: unescapeColumn(inner) };
}

function skipSpaces(text: string, i: number): number {
	while (i < text.length && text[i] === ' ') {
		i++;
	}
	return i;
}

/** `[[#Headers],[A]:[C]]`, `[@[Unit Price]]`, `[[Col]]`: items separated by commas. */
function parseComplex(inner: string): Parts | null {
	const parts: Parts = { specifiers: [] };
	let i = skipSpaces(inner, 0);
	if (inner[i] === '@') {
		parts.specifiers.push('#This Row');
		i = skipSpaces(inner, i + 1);
	}
	for (;;) {
		if (inner[i] !== '[') {
			return null;
		}
		const close = closingBracket(inner, i);
		if (close < 0) {
			return null;
		}
		const item = inner.slice(i + 1, close);
		i = skipSpaces(inner, close + 1);
		if (item[0] === '#') {
			const spec = SPECIFIERS.get(item.toUpperCase());
			if (!spec || parts.specifiers.includes(spec)) {
				return null;
			}
			parts.specifiers.push(spec);
		} else {
			if (parts.columnStart !== undefined) {
				return null;
			}
			parts.columnStart = unescapeColumn(item);
			if (inner[i] === ':') {
				i = skipSpaces(inner, i + 1);
				const close2 = inner[i] === '[' ? closingBracket(inner, i) : -1;
				if (close2 < 0) {
					return null;
				}
				parts.columnEnd = unescapeColumn(inner.slice(i + 1, close2));
				i = skipSpaces(inner, close2 + 1);
			}
		}
		if (i >= inner.length) {
			return parts;
		}
		if (inner[i] !== ',') {
			return null;
		}
		i = skipSpaces(inner, i + 1);
	}
}

/**
 * Parse the text of a structured reference token into an AST node, or
 * `null` when the bracket content is malformed.
 */
export function parseStructuredReference(text: string): StructuredRefNode | null {
	const open = text.indexOf('[');
	if (open < 0 || text[text.length - 1] !== ']') {
		return null;
	}
	const inner = text.slice(open + 1, -1);
	const parts = hasUnescapedBracket(inner) ? parseComplex(inner) : parseSimple(inner);
	if (parts === null) {
		return null;
	}
	const node: StructuredRefNode = { type: 'structured', specifiers: parts.specifiers, text };
	if (open > 0) {
		node.table = text.slice(0, open);
	}
	if (parts.columnStart !== undefined) {
		node.columnStart = parts.columnStart;
	}
	if (parts.columnEnd !== undefined) {
		node.columnEnd = parts.columnEnd;
	}
	return node;
}

/** Column names that can be written without their own brackets (`Table1[Sales Amount]`). */
function isSimpleColumn(name: string): boolean {
	return name !== '' && !/[\t\n\r,:.[\]#'"{}$^&*+=\-<>/@]/.test(name) && name.trim() === name;
}

/**
 * Canonical text of a structured reference. Display form uses `@` for
 * "this row" (`Table1[@Qty]`); the file form spells it `[#This Row]`
 * (`Table1[[#This Row],[Qty]]`), the only syntax the file format defines.
 */
export function formatStructuredReference(
	node: Pick<StructuredRefNode, 'table' | 'specifiers' | 'columnStart' | 'columnEnd'>,
	options: { forFile?: boolean } = {},
): string {
	const table = node.table ?? '';
	const { specifiers, columnStart, columnEnd } = node;
	let columns = '';
	if (columnStart !== undefined) {
		columns = `[${escapeStructuredColumn(columnStart)}]`;
		if (columnEnd !== undefined) {
			columns += `:[${escapeStructuredColumn(columnEnd)}]`;
		}
	}
	if (!options.forFile && specifiers.length === 1 && specifiers[0] === '#This Row') {
		if (columnStart === undefined) {
			return `${table}[@]`;
		}
		if (columnEnd === undefined && isSimpleColumn(columnStart)) {
			return `${table}[@${columnStart}]`;
		}
		return `${table}[@${columns}]`;
	}
	if (specifiers.length === 0) {
		if (columnStart === undefined) {
			return `${table}[]`;
		}
		if (columnEnd === undefined && isSimpleColumn(columnStart)) {
			return `${table}[${columnStart}]`;
		}
		return `${table}[${columns}]`;
	}
	if (specifiers.length === 1 && columnStart === undefined) {
		return `${table}[${specifiers[0]}]`;
	}
	const items = specifiers.map((s) => `[${s}]`);
	if (columns) {
		items.push(columns);
	}
	return `${table}[${items.join(',')}]`;
}

/** The reference uses the `@` (this row) shorthand, which the file format does not define. */
export function usesThisRowShorthand(text: string): boolean {
	const open = text.indexOf('[');
	if (open < 0) {
		return false;
	}
	let i = open + 1;
	while (text[i] === ' ') {
		i++;
	}
	return text[i] === '@';
}
