import { filePrefixFor } from './function-prefixes';

/**
 * Textual (regex) prefix handling, kept as the fallback for formulas the
 * tokenizer cannot classify cleanly.
 */

/**
 * Split a formula into alternating code / string-literal segments so that
 * text inside double quotes is never rewritten. Quoted sheet names are kept
 * inside code segments untouched as well.
 */
export function mapCodeSegments(formula: string, fn: (code: string) => string): string {
	let out = '';
	let i = 0;
	let codeStart = 0;
	while (i < formula.length) {
		const ch = formula[i];
		if (ch === '"' || ch === "'") {
			out += fn(formula.slice(codeStart, i));
			let j = i + 1;
			while (j < formula.length) {
				if (formula[j] === ch) {
					if (formula[j + 1] === ch) {
						j += 2;
						continue;
					}
					break;
				}
				j++;
			}
			out += formula.slice(i, j + 1);
			i = j + 1;
			codeStart = i;
			continue;
		}
		i++;
	}
	out += fn(formula.slice(codeStart));
	return out;
}

/** Regex version of `stripFormulaPrefixes`. */
export function stripPrefixesTextually(formula: string): string {
	return mapCodeSegments(formula, (code) => code.replace(/_xlfn\.|_xlws\.|_xlpm\./gi, ''));
}

/** Regex version of `addFormulaPrefixes` (function calls only). */
export function addPrefixesTextually(formula: string): string {
	return mapCodeSegments(formula, (code) =>
		code.replace(
			/(^|[^A-Za-z0-9_.])([A-Za-z][A-Za-z0-9.]*)\s*\(/g,
			(m, lead: string, name: string) => {
				const prefix = filePrefixFor(name);
				return prefix ? `${lead}${prefix}${name}(` : m;
			},
		),
	);
}
