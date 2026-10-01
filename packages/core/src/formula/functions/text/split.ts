import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionArg, FunctionContext, FunctionDefinition, Scalar } from '../../types';
import { makeArray } from '../../types';
import { booleanArg, defaultedNumber, hasValue, scalarArg, textArg } from '../math/args';
import { XLFN } from '../math/define';
import { foldCase, textList } from './text-utils';

interface Match {
	start: number;
	end: number;
}

/** Longest delimiter starting at `i`, or 0. */
function matchAt(text: string, delimiters: readonly string[], i: number): number {
	let best = 0;
	for (const d of delimiters) {
		if (d.length > best && text.startsWith(d, i)) {
			best = d.length;
		}
	}
	return best;
}

/** Non-overlapping delimiter occurrences scanning from the start (at most `limit`). */
function forwardMatches(text: string, delimiters: readonly string[], limit: number): Match[] {
	const out: Match[] = [];
	let i = 0;
	while (i < text.length && out.length < limit) {
		const len = matchAt(text, delimiters, i);
		if (len > 0) {
			out.push({ start: i, end: i + len });
			i += len;
		} else {
			i++;
		}
	}
	return out;
}

/** Non-overlapping delimiter occurrences scanning from the end (at most `limit`). */
function backwardMatches(text: string, delimiters: readonly string[], limit: number): Match[] {
	const out: Match[] = [];
	let end = text.length;
	while (end > 0 && out.length < limit) {
		let len = 0;
		for (const d of delimiters) {
			if (d.length > len && d.length <= end && text.endsWith(d, end)) {
				len = d.length;
			}
		}
		if (len > 0) {
			out.push({ start: end - len, end });
			end -= len;
		} else {
			end--;
		}
	}
	return out;
}

/** A 0 / 1 switch argument (match_mode, match_end); anything else is #VALUE!. */
function switchArg(
	args: readonly FunctionArg[],
	index: number,
	ctx: FunctionContext,
): boolean | CellError {
	const v = defaultedNumber(args, index, ctx, 0);
	if (typeof v !== 'number') {
		return v;
	}
	return v === 0 || v === 1 ? v === 1 : ERRORS.VALUE;
}

/** TEXTBEFORE / TEXTAFTER. */
function textAround(name: string, before: boolean): FunctionDefinition {
	return {
		name,
		minArgs: 2,
		maxArgs: 6,
		params: ['value', 'array', 'value', 'value', 'value', 'value'],
		filePrefix: XLFN,
		impl: (args, ctx) => {
			const text = textArg(args[0]);
			if (isCellError(text)) {
				return text;
			}
			const delimiters = textList(args[1]);
			if (isCellError(delimiters)) {
				return delimiters;
			}
			const instance = defaultedNumber(args, 2, ctx, 1);
			if (typeof instance !== 'number') {
				return instance;
			}
			const ignoreCase = switchArg(args, 3, ctx);
			if (isCellError(ignoreCase)) {
				return ignoreCase;
			}
			const matchEnd = switchArg(args, 4, ctx);
			if (isCellError(matchEnd)) {
				return matchEnd;
			}
			const n = Math.trunc(instance);
			if (n === 0) {
				return ERRORS.VALUE;
			}
			if (text === '') {
				return '';
			}
			if (Math.abs(n) > text.length) {
				return ERRORS.VALUE;
			}
			const hay = ignoreCase ? foldCase(text) : text;
			const needles = ignoreCase ? delimiters.map(foldCase) : delimiters;
			let match: Match | undefined;
			if (needles.length === 0 || needles.includes('')) {
				// The empty delimiter matches immediately, at every position.
				const at = n > 0 ? n - 1 : text.length + n + 1;
				match = { start: at, end: at };
			} else if (n > 0) {
				const found = forwardMatches(hay, needles, n);
				if (found.length === n) {
					match = found[n - 1];
				} else if (matchEnd && found.length === n - 1) {
					match = { start: text.length, end: text.length };
				}
			} else {
				const found = backwardMatches(hay, needles, -n);
				if (found.length === -n) {
					match = found[-n - 1];
				} else if (matchEnd && found.length === -n - 1) {
					match = { start: 0, end: 0 };
				}
			}
			if (!match) {
				return hasValue(args, 5) ? (scalarArg(args, 5) as Scalar) : ERRORS.NA;
			}
			return before ? text.slice(0, match.start) : text.slice(match.end);
		},
	};
}

/** Split text at every delimiter occurrence. */
export function splitText(
	text: string,
	delimiters: readonly string[],
	ignoreCase: boolean,
): string[] {
	const hay = ignoreCase ? foldCase(text) : text;
	const needles = ignoreCase ? delimiters.map(foldCase) : delimiters;
	const parts: string[] = [];
	let from = 0;
	for (const m of forwardMatches(hay, needles, Infinity)) {
		parts.push(text.slice(from, m.start));
		from = m.end;
	}
	parts.push(text.slice(from));
	return parts;
}

function delimiterList(args: readonly FunctionArg[], index: number): string[] | CellError {
	if (index >= args.length) {
		return [];
	}
	const list = textList(args[index]);
	return isCellError(list) ? list : list.filter((d) => d !== '');
}

const TEXTSPLIT: FunctionDefinition = {
	name: 'TEXTSPLIT',
	minArgs: 2,
	maxArgs: 6,
	params: ['value', 'array', 'array', 'value', 'value', 'value'],
	filePrefix: XLFN,
	returnsArray: true,
	impl: (args, ctx) => {
		const text = textArg(args[0]);
		if (isCellError(text)) {
			return text;
		}
		const columnDelimiters = delimiterList(args, 1);
		if (isCellError(columnDelimiters)) {
			return columnDelimiters;
		}
		const rowDelimiters = delimiterList(args, 2);
		if (isCellError(rowDelimiters)) {
			return rowDelimiters;
		}
		if (columnDelimiters.length === 0 && rowDelimiters.length === 0) {
			return ERRORS.VALUE;
		}
		const ignoreEmpty = hasValue(args, 3) ? booleanArg(args[3]) : false;
		if (isCellError(ignoreEmpty)) {
			return ignoreEmpty;
		}
		const ignoreCase = switchArg(args, 4, ctx);
		if (isCellError(ignoreCase)) {
			return ignoreCase;
		}
		const pad: Scalar = hasValue(args, 5) ? (scalarArg(args, 5) as Scalar) : ERRORS.NA;
		const dropEmpty = (parts: string[]) => (ignoreEmpty ? parts.filter((p) => p !== '') : parts);
		const lines = dropEmpty(
			rowDelimiters.length > 0 ? splitText(text, rowDelimiters, ignoreCase) : [text],
		);
		const cells = lines.map((line) =>
			dropEmpty(
				columnDelimiters.length > 0 ? splitText(line, columnDelimiters, ignoreCase) : [line],
			),
		);
		const width = cells.reduce((w, row) => Math.max(w, row.length), 0);
		if (width === 0) {
			return ERRORS.CALC;
		}
		return makeArray(
			cells.map((row): Scalar[] => [...row, ...new Array<Scalar>(width - row.length).fill(pad)]),
		);
	},
};

export const SPLIT_FUNCTIONS: FunctionDefinition[] = [
	textAround('TEXTBEFORE', true),
	textAround('TEXTAFTER', false),
	TEXTSPLIT,
];
