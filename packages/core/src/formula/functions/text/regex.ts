import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionArg, FunctionContext, FunctionDefinition, Scalar } from '../../types';
import { makeArray } from '../../types';
import { defaultedNumber, textArg } from '../math/args';
import { XLFN } from '../math/define';
import { compilePattern, expandReplacement } from './regex-compile';
import { limitText } from './text-utils';

/** An integer option argument restricted to `allowed` values (defaults apply when empty). */
function optionArg(
	args: readonly FunctionArg[],
	index: number,
	ctx: FunctionContext,
	fallback: number,
	allowed: readonly number[],
): number | CellError {
	const v = defaultedNumber(args, index, ctx, fallback);
	if (typeof v !== 'number') {
		return v;
	}
	const n = Math.trunc(v);
	return allowed.includes(n) ? n : ERRORS.VALUE;
}

/** Text and compiled pattern of the first two arguments. */
function textAndPattern(
	args: readonly FunctionArg[],
	ignoreCase: boolean,
	global: boolean,
): { text: string; re: RegExp } | CellError {
	const text = textArg(args[0]);
	if (isCellError(text)) {
		return text;
	}
	const pattern = textArg(args[1]);
	if (isCellError(pattern)) {
		return pattern;
	}
	const re = compilePattern(pattern, ignoreCase, global);
	return re ? { text, re } : ERRORS.VALUE;
}

const REGEXTEST: FunctionDefinition = {
	name: 'REGEXTEST',
	minArgs: 2,
	maxArgs: 3,
	params: ['value', 'value', 'value'],
	filePrefix: XLFN,
	impl: (args, ctx) => {
		const mode = optionArg(args, 2, ctx, 0, [0, 1]);
		if (isCellError(mode)) {
			return mode;
		}
		const p = textAndPattern(args, mode === 1, false);
		return isCellError(p) ? p : p.re.test(p.text);
	},
};

/**
 * REGEXEXTRACT return modes: 0 the first match, 1 every match (spilled
 * down a column), 2 the capturing groups of the first match (across a row).
 */
const REGEXEXTRACT: FunctionDefinition = {
	name: 'REGEXEXTRACT',
	minArgs: 2,
	maxArgs: 4,
	params: ['value', 'value', 'value', 'value'],
	filePrefix: XLFN,
	returnsArray: true,
	impl: (args, ctx) => {
		const returnMode = optionArg(args, 2, ctx, 0, [0, 1, 2]);
		if (isCellError(returnMode)) {
			return returnMode;
		}
		const mode = optionArg(args, 3, ctx, 0, [0, 1]);
		if (isCellError(mode)) {
			return mode;
		}
		const p = textAndPattern(args, mode === 1, returnMode === 1);
		if (isCellError(p)) {
			return p;
		}
		if (returnMode === 1) {
			const all = [...p.text.matchAll(p.re)].map((m): Scalar[] => [m[0]]);
			return all.length === 0 ? ERRORS.NA : makeArray(all);
		}
		const m = p.re.exec(p.text);
		if (!m) {
			return ERRORS.NA;
		}
		if (returnMode === 0) {
			return m[0];
		}
		const groups = m.length > 1 ? m.slice(1).map((g) => g ?? '') : [m[0]];
		return makeArray([groups]);
	},
};

/** REGEXREPLACE occurrence: 0 replaces every match, n the n-th, -n the n-th from the end. */
const REGEXREPLACE: FunctionDefinition = {
	name: 'REGEXREPLACE',
	minArgs: 3,
	maxArgs: 5,
	params: ['value', 'value', 'value', 'value', 'value'],
	filePrefix: XLFN,
	impl: (args, ctx) => {
		const replacement = textArg(args[2]);
		if (isCellError(replacement)) {
			return replacement;
		}
		const occurrence = defaultedNumber(args, 3, ctx, 0);
		if (typeof occurrence !== 'number') {
			return occurrence;
		}
		const mode = optionArg(args, 4, ctx, 0, [0, 1]);
		if (isCellError(mode)) {
			return mode;
		}
		const p = textAndPattern(args, mode === 1, true);
		if (isCellError(p)) {
			return p;
		}
		const matches = [...p.text.matchAll(p.re)];
		const n = Math.trunc(occurrence);
		const target = n > 0 ? n - 1 : matches.length + n;
		let out = '';
		let from = 0;
		matches.forEach((m, i) => {
			if (n !== 0 && i !== target) {
				return;
			}
			const at = m.index ?? 0;
			out += p.text.slice(from, at) + expandReplacement(replacement, m);
			from = at + m[0].length;
		});
		return limitText(out + p.text.slice(from));
	},
};

export const REGEX_FUNCTIONS: FunctionDefinition[] = [REGEXTEST, REGEXEXTRACT, REGEXREPLACE];
