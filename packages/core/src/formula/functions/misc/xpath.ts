import { parseXml } from '../../../xml/parser';
import type { XmlElement, XmlNode } from '../../../xml/types';

/**
 * A small XPath 1.0 subset for FILTERXML: absolute paths with `/` and `//`
 * steps, name tests and `*`, `@attr`, `text()`, `.` and `..`, and predicates
 * with positions, `last()`, `position()` comparisons, `.` / `@a` comparisons,
 * `contains()`, `starts-with()`, `not()`, `and` and `or`. That covers the
 * usual "split text" uses of FILTERXML.
 */

type Item = XmlElement | { attr: string; value: string; owner: XmlElement } | { text: string };

interface Ctx {
	parent: Map<XmlElement, XmlElement>;
}

const isElement = (x: Item | XmlNode): x is XmlElement =>
	typeof x === 'object' && x !== null && 'name' in x && 'children' in x;

function localName(name: string): string {
	const i = name.indexOf(':');
	return i < 0 ? name : name.slice(i + 1);
}

function textOf(item: Item): string {
	if ('text' in item) {
		return item.text;
	}
	if ('attr' in item) {
		return item.value;
	}
	let out = '';
	const walk = (el: XmlElement) => {
		for (const c of el.children) {
			if (typeof c === 'string') {
				out += c;
			} else if (isElement(c)) {
				walk(c);
			}
		}
	};
	walk(item);
	return out;
}

function descendants(el: XmlElement, out: XmlElement[]): void {
	for (const c of el.children) {
		if (isElement(c)) {
			out.push(c);
			descendants(c, out);
		}
	}
}

/** Split a path at top-level `/` (outside predicates and quotes). */
function splitSteps(path: string): Array<{ deep: boolean; step: string }> {
	const steps: Array<{ deep: boolean; step: string }> = [];
	let i = 0;
	while (i < path.length) {
		let deep = false;
		if (path.startsWith('//', i)) {
			deep = true;
			i += 2;
		} else if (path[i] === '/') {
			i += 1;
		}
		let depth = 0;
		let quote = '';
		let j = i;
		for (; j < path.length; j++) {
			const ch = path[j];
			if (quote) {
				if (ch === quote) {
					quote = '';
				}
			} else if (ch === '"' || ch === "'") {
				quote = ch;
			} else if (ch === '[') {
				depth++;
			} else if (ch === ']') {
				depth--;
			} else if (ch === '/' && depth === 0) {
				break;
			}
		}
		steps.push({ deep, step: path.slice(i, j).trim() });
		i = j;
	}
	return steps;
}

function splitPredicates(step: string): { test: string; predicates: string[] } {
	const at = step.indexOf('[');
	if (at < 0) {
		return { test: step, predicates: [] };
	}
	const predicates: string[] = [];
	let depth = 0;
	let start = -1;
	let quote = '';
	for (let i = at; i < step.length; i++) {
		const ch = step[i];
		if (quote) {
			if (ch === quote) {
				quote = '';
			}
		} else if (ch === '"' || ch === "'") {
			quote = ch;
		} else if (ch === '[') {
			if (depth++ === 0) {
				start = i + 1;
			}
		} else if (ch === ']' && --depth === 0) {
			predicates.push(step.slice(start, i).trim());
		}
	}
	return { test: step.slice(0, at).trim(), predicates };
}

function applyStep(items: Item[], deep: boolean, step: string, ctx: Ctx): Item[] {
	const { test, predicates } = splitPredicates(step);
	let out: Item[] = [];
	for (const item of items) {
		if (!isElement(item)) {
			continue;
		}
		const pool: XmlElement[] = [];
		if (deep) {
			descendants(item, pool);
		} else {
			for (const c of item.children) {
				if (isElement(c)) {
					pool.push(c);
				}
			}
		}
		let matched: Item[];
		if (test === '.') {
			matched = [item];
		} else if (test === '..') {
			const p = ctx.parent.get(item);
			matched = p ? [p] : [];
		} else if (test === 'text()') {
			const owners = deep ? [item, ...pool] : [item];
			matched = owners.flatMap((o) =>
				o.children.filter((c): c is string => typeof c === 'string').map((t) => ({ text: t })),
			);
		} else if (test.startsWith('@')) {
			const name = test.slice(1);
			const owners = deep ? [item, ...pool] : [item];
			matched = owners.flatMap((o) =>
				Object.entries(o.attrs)
					.filter(([k]) => name === '*' || localName(k) === name || k === name)
					.map(([k, v]) => ({ attr: k, value: v, owner: o })),
			);
		} else {
			matched = pool.filter(
				(el) => test === '*' || el.name === test || localName(el.name) === test,
			);
		}
		for (const pred of predicates) {
			matched = matched.filter((m, i, all) => predicate(pred, m, i + 1, all.length));
		}
		out = out.concat(matched);
	}
	return out;
}

function literal(
	token: string,
	item: Item,
	position: number,
	size: number,
): string | number | boolean {
	const t = token.trim();
	if ((t.startsWith('"') && t.endsWith('"')) || (t.startsWith("'") && t.endsWith("'"))) {
		return t.slice(1, -1);
	}
	if (/^-?\d+(\.\d+)?$/.test(t)) {
		return Number(t);
	}
	if (t === 'position()') {
		return position;
	}
	if (t === 'last()') {
		return size;
	}
	if (t === '.' || t === 'text()') {
		return textOf(item);
	}
	if (t.startsWith('@') && isElement(item)) {
		return item.attrs[t.slice(1)] ?? '';
	}
	const fn = /^([a-z-]+)\((.*)\)$/.exec(t);
	if (fn) {
		const args = splitArgs(fn[2]).map((a) => String(literal(a, item, position, size)));
		switch (fn[1]) {
			case 'contains':
				return args[0].includes(args[1] ?? '');
			case 'starts-with':
				return args[0].startsWith(args[1] ?? '');
			case 'string-length':
				return (args[0] ?? textOf(item)).length;
			case 'number':
				return Number(args[0] ?? textOf(item));
			case 'not':
				return !truthy(literal(fn[2], item, position, size));
		}
	}
	if (isElement(item)) {
		const child = item.children.find((c) => isElement(c) && localName(c.name) === t);
		return child && isElement(child) ? textOf(child) : '';
	}
	return '';
}

function splitArgs(text: string): string[] {
	const out: string[] = [];
	let depth = 0;
	let quote = '';
	let start = 0;
	for (let i = 0; i < text.length; i++) {
		const ch = text[i];
		if (quote) {
			if (ch === quote) {
				quote = '';
			}
		} else if (ch === '"' || ch === "'") {
			quote = ch;
		} else if (ch === '(') {
			depth++;
		} else if (ch === ')') {
			depth--;
		} else if (ch === ',' && depth === 0) {
			out.push(text.slice(start, i));
			start = i + 1;
		}
	}
	out.push(text.slice(start));
	return out.filter((a) => a.trim() !== '');
}

const truthy = (v: string | number | boolean) =>
	typeof v === 'boolean' ? v : typeof v === 'number' ? v !== 0 : v !== '';

function compare(a: string | number | boolean, op: string, b: string | number | boolean): boolean {
	const num = typeof a === 'number' || typeof b === 'number';
	const x = num ? Number(a) : String(a);
	const y = num ? Number(b) : String(b);
	switch (op) {
		case '=':
			return x === y;
		case '!=':
			return x !== y;
		case '<':
			return x < y;
		case '<=':
			return x <= y;
		case '>':
			return x > y;
		default:
			return x >= y;
	}
}

function predicate(expr: string, item: Item, position: number, size: number): boolean {
	const or = splitKeyword(expr, ' or ');
	if (or.length > 1) {
		return or.some((p) => predicate(p, item, position, size));
	}
	const and = splitKeyword(expr, ' and ');
	if (and.length > 1) {
		return and.every((p) => predicate(p, item, position, size));
	}
	const mod = /^(.+?)\s+mod\s+(\d+)\s*(=|!=)\s*(\d+)$/.exec(expr);
	if (mod) {
		const v = Number(literal(mod[1], item, position, size)) % Number(mod[2]);
		return compare(v, mod[3], Number(mod[4]));
	}
	const cmp = topLevelComparison(expr);
	if (cmp) {
		return compare(
			literal(cmp.left, item, position, size),
			cmp.op,
			literal(cmp.right, item, position, size),
		);
	}
	const v = literal(expr, item, position, size);
	return typeof v === 'number' ? v === position : truthy(v);
}

/** The first comparison operator outside parentheses and quotes. */
function topLevelComparison(expr: string): { left: string; op: string; right: string } | null {
	let depth = 0;
	let quote = '';
	for (let i = 0; i < expr.length; i++) {
		const ch = expr[i];
		if (quote) {
			if (ch === quote) {
				quote = '';
			}
		} else if (ch === '"' || ch === "'") {
			quote = ch;
		} else if (ch === '(' || ch === '[') {
			depth++;
		} else if (ch === ')' || ch === ']') {
			depth--;
		} else if (depth === 0 && '=!<>'.includes(ch)) {
			const two = expr.slice(i, i + 2);
			const op = two === '!=' || two === '<=' || two === '>=' ? two : ch;
			if (op === '!') {
				continue;
			}
			return { left: expr.slice(0, i), op, right: expr.slice(i + op.length) };
		}
	}
	return null;
}

function splitKeyword(expr: string, keyword: string): string[] {
	const out: string[] = [];
	let depth = 0;
	let quote = '';
	let start = 0;
	for (let i = 0; i < expr.length; i++) {
		const ch = expr[i];
		if (quote) {
			if (ch === quote) {
				quote = '';
			}
		} else if (ch === '"' || ch === "'") {
			quote = ch;
		} else if (ch === '(' || ch === '[') {
			depth++;
		} else if (ch === ')' || ch === ']') {
			depth--;
		} else if (depth === 0 && expr.startsWith(keyword, i)) {
			out.push(expr.slice(start, i));
			start = i + keyword.length;
			i += keyword.length - 1;
		}
	}
	out.push(expr.slice(start));
	return out.map((s) => s.trim());
}

/** Text of every node selected by `xpath` in `xml`, or null when the XML or the path is invalid. */
export function evaluateXPath(xml: string, xpath: string): string[] | null {
	let root: XmlElement;
	try {
		root = parseXml(xml).root;
	} catch {
		return null;
	}
	const path = xpath.trim();
	if (!path.startsWith('/')) {
		return null;
	}
	const parent = new Map<XmlElement, XmlElement>();
	const index = (el: XmlElement) => {
		for (const c of el.children) {
			if (isElement(c)) {
				parent.set(c, el);
				index(c);
			}
		}
	};
	index(root);
	// A virtual document node above the root element.
	const doc: XmlElement = { name: '#document', attrs: {}, children: [root] };
	parent.set(root, doc);
	let items: Item[] = [doc];
	try {
		for (const { deep, step } of splitSteps(path)) {
			if (step === '') {
				continue;
			}
			items = applyStep(items, deep, step, { parent });
		}
	} catch {
		return null;
	}
	return items.map(textOf);
}
