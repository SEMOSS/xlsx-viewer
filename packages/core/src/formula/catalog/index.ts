/**
 * Function catalog: names, categories, descriptions and argument help for
 * every Excel worksheet function (Insert Function dialog, formula
 * autocomplete and argument tooltips). The data lives in one file per
 * category (`financial.ts`, `math.ts`, ...), gathered by `entries.ts`;
 * `versions.ts` records the Excel release that introduced newer functions.
 */
import { normalizeFunctionName } from '../language/names';
import { ALL_FUNCTION_INFO } from './entries';
import type { FunctionArgInfo, FunctionCategory, FunctionInfo } from './types';

export type { FunctionArgInfo, FunctionCategory, FunctionInfo } from './types';

/** Insert Function dialog categories, in Excel's order. */
export const FUNCTION_CATEGORIES: readonly FunctionCategory[] = [
	'Financial',
	'Date & Time',
	'Math & Trig',
	'Statistical',
	'Lookup & Reference',
	'Database',
	'Text',
	'Logical',
	'Information',
	'Engineering',
	'Cube',
	'Compatibility',
	'Web',
];

/** Ordinal name order, independent of the runtime locale. */
function byName(a: FunctionInfo, b: FunctionInfo): number {
	return a.name < b.name ? -1 : a.name > b.name ? 1 : 0;
}

const CATALOG = new Map<string, FunctionInfo>(ALL_FUNCTION_INFO.map((f) => [f.name, f]));
const SORTED: readonly FunctionInfo[] = [...CATALOG.values()].sort(byName);

/** Help for a function (any case; file prefixes such as `_xlfn.` are ignored), or undefined. */
export function getFunctionInfo(name: string): FunctionInfo | undefined {
	return CATALOG.get(normalizeFunctionName(name.trim()));
}

/** All catalogued functions, optionally filtered by category, sorted by name. */
export function listFunctions(category?: FunctionCategory): FunctionInfo[] {
	return category ? SORTED.filter((f) => f.category === category) : [...SORTED];
}

/** Functions whose name starts with `prefix` (autocomplete), best matches first. */
export function searchFunctions(prefix: string, limit = 12): FunctionInfo[] {
	const p = prefix.trim().toUpperCase();
	if (!p) {
		return [];
	}
	const out: FunctionInfo[] = [];
	for (const f of SORTED) {
		if (out.length >= limit) {
			break;
		}
		if (f.name.startsWith(p)) {
			out.push(f);
		}
	}
	return out;
}

/** One argument as shown in a syntax line: `name` or `[name]`. */
function formatArg(a: FunctionArgInfo): string {
	return a.optional ? `[${a.name}]` : a.name;
}

/**
 * Excel-style syntax line: `VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])`.
 *
 * A run of repeating arguments is one group. At the end of the list it is
 * followed by an ellipsis (`SUM(number1, [number2], ...)`,
 * `SUMIFS(..., [criteria_range2, criteria2], ...)`); before other arguments
 * the ellipsis goes inside the group (`LAMBDA([parameter1, ...], calculation)`).
 */
export function formatFunctionSyntax(info: FunctionInfo): string {
	const parts: string[] = [];
	const args = info.args;
	let i = 0;
	while (i < args.length) {
		if (!args[i].repeating) {
			parts.push(formatArg(args[i]));
			i++;
			continue;
		}
		let end = i;
		while (end < args.length && args[end].repeating) {
			end++;
		}
		const group = args.slice(i, end);
		const allOptional = group.every((a) => a.optional);
		const names = allOptional ? group.map((a) => a.name) : group.map(formatArg);
		if (end < args.length) {
			const inner = [...names, '...'].join(', ');
			parts.push(allOptional ? `[${inner}]` : inner);
		} else {
			parts.push(allOptional ? `[${names.join(', ')}]` : names.join(', '), '...');
		}
		i = end;
	}
	return `${info.name}(${parts.join(', ')})`;
}
