import type {
	AstNode,
	CellRefNode,
	FunctionNode,
	NameNode,
	RangeRefNode,
	StructuredRefNode,
} from '../formula/ast';

/**
 * Static analysis of a parsed formula: the references it reads (as
 * written), the functions it calls and a few flags. The result depends only
 * on the formula text, so it is cached alongside the AST and resolved per
 * cell by `resolve.ts`.
 */
export interface FormulaAnalysis {
	/** Single-cell references whose values are read (spill anchors of `A1#` included). */
	readonly cells: CellRefNode[];
	readonly ranges: RangeRefNode[];
	/** Names not bound by LET / LAMBDA (defined names, table names). */
	readonly names: NameNode[];
	readonly structured: StructuredRefNode[];
	/** Called functions (upper case) not bound by LET / LAMBDA. */
	readonly functions: ReadonlySet<string>;
	/** Calls a function that is always volatile. */
	readonly volatile: boolean;
	/** References another workbook. */
	readonly external: boolean;
	/** Contains SUBTOTAL or AGGREGATE (such cells are skipped by other subtotals). */
	readonly subtotal: boolean;
}

/** Functions that recalculate on every recalculation. */
export const VOLATILE_FUNCTIONS: ReadonlySet<string> = new Set([
	'NOW',
	'TODAY',
	'RAND',
	'RANDBETWEEN',
	'RANDARRAY',
	'OFFSET',
	'INDIRECT',
	'CELL',
	'INFO',
]);

/** Functions whose reference arguments are inspected, not read (ROW(A1) does not depend on A1's value). */
const REFERENCE_ONLY: ReadonlySet<string> = new Set([
	'ROW',
	'ROWS',
	'COLUMN',
	'COLUMNS',
	'AREAS',
	'ISREF',
	'SHEET',
	'SHEETS',
]);

interface Collector {
	cells: CellRefNode[];
	ranges: RangeRefNode[];
	names: NameNode[];
	structured: StructuredRefNode[];
	functions: Set<string>;
	volatile: boolean;
	external: boolean;
	subtotal: boolean;
}

type Scope = ReadonlySet<string> | null;

function extend(scope: Scope, names: readonly string[]): Set<string> {
	const next = new Set(scope ?? []);
	for (const n of names) {
		next.add(n);
	}
	return next;
}

function paramName(node: AstNode): string | null {
	if (node.type === 'name' && node.sheet === undefined) {
		return node.name.toUpperCase();
	}
	if (node.type === 'structured' && node.table === undefined && node.columnStart !== undefined) {
		return node.columnStart.toUpperCase();
	}
	return null;
}

function walkFunction(node: FunctionNode, scope: Scope, out: Collector): void {
	const name = node.name;
	const args = node.args;
	if (scope?.has(name)) {
		args.forEach((a) => walk(a, scope, false, out));
		return;
	}
	out.functions.add(name);
	if (VOLATILE_FUNCTIONS.has(name)) {
		out.volatile = true;
	}
	if (name === 'SUBTOTAL' || name === 'AGGREGATE') {
		out.subtotal = true;
	}
	if (name === 'LET') {
		let inner: Set<string> = extend(scope, []);
		for (let i = 0; i + 1 < args.length; i += 2) {
			walk(args[i + 1], inner, false, out);
			const bound = paramName(args[i]);
			if (bound !== null) {
				inner = extend(inner, [bound]);
			}
		}
		if (args.length > 0) {
			walk(args[args.length - 1], inner, false, out);
		}
		return;
	}
	if (name === 'LAMBDA') {
		const params = args
			.slice(0, -1)
			.map(paramName)
			.filter((p): p is string => p !== null);
		if (args.length > 0) {
			walk(args[args.length - 1], extend(scope, params), false, out);
		}
		return;
	}
	const refOnly = REFERENCE_ONLY.has(name);
	args.forEach((a, i) => {
		// OFFSET's base and CELL's reference (other than "contents") are positions, not values.
		const positional =
			refOnly ||
			(name === 'OFFSET' && i === 0) ||
			(name === 'CELL' &&
				i === 1 &&
				args[0].type === 'string' &&
				args[0].value.trim().toLowerCase() !== 'contents');
		walk(a, scope, positional, out);
	});
}

/**
 * Collect dependencies. `refOnly` marks positions where a bare reference is
 * inspected rather than read; references nested in calls there still count.
 */
function walk(node: AstNode, scope: Scope, refOnly: boolean, out: Collector): void {
	switch (node.type) {
		case 'cell':
			if (node.sheet?.workbook !== undefined) {
				out.external = true;
			} else if (!refOnly) {
				out.cells.push(node);
			}
			return;
		case 'range':
			if (node.sheet?.workbook !== undefined) {
				out.external = true;
			} else if (!refOnly) {
				out.ranges.push(node);
			}
			return;
		case 'spill':
			if (node.ref.sheet?.workbook !== undefined) {
				out.external = true;
			} else {
				out.cells.push(node.ref);
			}
			return;
		case 'name':
			if (node.sheet?.workbook !== undefined) {
				out.external = true;
			} else if (node.sheet !== undefined || !scope?.has(node.name.toUpperCase())) {
				out.names.push(node);
			}
			return;
		case 'structured':
			if (!refOnly) {
				out.structured.push(node);
			}
			return;
		case 'error':
			if (node.sheet?.workbook !== undefined) {
				out.external = true;
			}
			return;
		case 'function':
			walkFunction(node, scope, out);
			return;
		case 'call':
			walk(node.callee, scope, false, out);
			node.args.forEach((a) => walk(a, scope, false, out));
			return;
		case 'unary':
			walk(node.operand, scope, refOnly && node.op !== '-', out);
			return;
		case 'percent':
			walk(node.operand, scope, false, out);
			return;
		case 'binary': {
			const keep = refOnly && (node.op === ':' || node.op === ',' || node.op === ' ');
			walk(node.left, scope, keep, out);
			walk(node.right, scope, keep, out);
			return;
		}
		default:
			return;
	}
}

/** Analyse a parsed formula. */
export function analyzeFormula(ast: AstNode): FormulaAnalysis {
	const out: Collector = {
		cells: [],
		ranges: [],
		names: [],
		structured: [],
		functions: new Set(),
		volatile: false,
		external: false,
		subtotal: false,
	};
	walk(ast, null, false, out);
	return out;
}
