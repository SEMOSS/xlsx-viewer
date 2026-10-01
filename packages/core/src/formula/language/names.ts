/** `_xlfn.`, `_xlws.` and `_xlpm.` file prefixes (any order, any case) at the start of a name. */
const FILE_PREFIXES = /^(?:_xlfn\.|_xlws\.|_xlpm\.)+/i;

/** Remove the future-function / parameter prefixes Excel writes in files. */
export function stripFilePrefixes(text: string): string {
	return text.charCodeAt(0) === 95 /* _ */ ? text.replace(FILE_PREFIXES, '') : text;
}

/** Canonical function name: prefixes removed, upper case (`_xlfn._xlws.filter` -> `FILTER`). */
export function normalizeFunctionName(text: string): string {
	return stripFilePrefixes(text).toUpperCase();
}
