/**
 * Builders used by the catalog data files. Entries and their arguments are
 * frozen: the catalog hands out shared objects, so callers must not mutate them.
 */
import type { FunctionArgInfo, FunctionCategory, FunctionInfo } from './types';
import { addedIn } from './versions';

/** Required argument. */
export function arg(name: string, description: string): FunctionArgInfo {
	return Object.freeze({ name, description });
}

/** Optional argument, shown as `[name]`. */
export function opt(name: string, description: string): FunctionArgInfo {
	return Object.freeze({ name, description, optional: true });
}

/**
 * Optional repeating argument. Consecutive repeating arguments form one group
 * that can be given any number of times: `[criteria_range2, criteria2], ...`.
 */
export function more(name: string, description: string): FunctionArgInfo {
	return Object.freeze({ name, description, optional: true, repeating: true });
}

/** `name1, [name2], ...`: a required first item followed by optional repeats. */
export function series(name: string, first: string, rest: string): readonly FunctionArgInfo[] {
	return Object.freeze([arg(`${name}1`, first), more(`${name}2`, rest)]);
}

/** Signature of the entry builder returned by `categoryBuilder`. */
export type EntryBuilder = (
	name: string,
	description: string,
	args?: readonly FunctionArgInfo[],
) => FunctionInfo;

/** Entry builder for one Insert Function category; `since` comes from `versions.ts`. */
export function categoryBuilder(category: FunctionCategory): EntryBuilder {
	return (name, description, args = []) => {
		const list = [...args];
		Object.freeze(list);
		const since = addedIn(name);
		const info: FunctionInfo = since
			? { name, category, description, args: list, since }
			: { name, category, description, args: list };
		return Object.freeze(info);
	};
}
