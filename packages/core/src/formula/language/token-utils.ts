import { normalizeFunctionName } from './names';
import type { FormulaToken } from './tokens';

/** A replacement of `[start, end)` in the source text (`start === end` inserts). */
export interface TextEdit {
	start: number;
	end: number;
	text: string;
}

/**
 * Apply non-overlapping edits to `source`. Insertions sort before
 * replacements at the same offset; equal edits keep their creation order.
 */
export function applyEdits(source: string, edits: TextEdit[]): string {
	if (edits.length === 0) {
		return source;
	}
	if (edits.length > 1) {
		edits.sort((a, b) => a.start - b.start || a.end - a.start - (b.end - b.start));
	}
	let out = '';
	let pos = 0;
	for (const e of edits) {
		out += source.slice(pos, e.start) + e.text;
		pos = e.end;
	}
	return out + source.slice(pos);
}

/** The tokenizer met characters it could not classify or an unterminated literal. */
export function hasInvalidTokens(tokens: readonly FormulaToken[]): boolean {
	for (const t of tokens) {
		if (t.type === 'unknown' || t.incomplete) {
			return true;
		}
	}
	return false;
}

/** Index of the first non-whitespace token at or after `i` (`tokens.length` when none). */
export function skipWhitespace(tokens: readonly FormulaToken[], i: number): number {
	while (i < tokens.length && (tokens[i].type === 'whitespace' || tokens[i].type === 'equals')) {
		i++;
	}
	return i;
}

/** Index of the last non-whitespace token before `i` (-1 when none). */
export function previousSignificant(tokens: readonly FormulaToken[], i: number): number {
	let k = i - 1;
	while (k >= 0 && (tokens[k].type === 'whitespace' || tokens[k].type === 'equals')) {
		k--;
	}
	return k;
}

function opens(t: FormulaToken): boolean {
	return t.type === 'paren-open' || t.type === 'call-open' || t.type === 'array-open';
}

function closes(t: FormulaToken): boolean {
	return t.type === 'paren-close' || t.type === 'call-close' || t.type === 'array-close';
}

/** Index of the token closing the bracket opened at `open`, or `tokens.length` when unclosed. */
export function matchingClose(tokens: readonly FormulaToken[], open: number): number {
	let depth = 0;
	for (let i = open; i < tokens.length; i++) {
		const t = tokens[i];
		if (opens(t)) {
			depth++;
		} else if (closes(t) && --depth === 0) {
			return i;
		}
	}
	return tokens.length;
}

/** Index just past the primary operand starting at `i` (calls, groups, arrays, `A1#`). */
function primaryEnd(tokens: readonly FormulaToken[], i: number): number {
	const t = tokens[i];
	if (t === undefined) {
		return i;
	}
	let end: number;
	switch (t.type) {
		case 'function':
			end = tokens[i + 1]?.type === 'call-open' ? matchingClose(tokens, i + 1) + 1 : i + 1;
			break;
		case 'paren-open':
		case 'array-open':
			end = matchingClose(tokens, i) + 1;
			break;
		case 'ref':
		case 'name':
			return tokens[i + 1]?.type === 'spill' ? i + 2 : i + 1;
		case 'structured':
		case 'error':
		case 'number':
		case 'string':
		case 'boolean':
			return i + 1;
		default:
			return i;
	}
	while (end < tokens.length && tokens[end].type === 'call-open') {
		end = matchingClose(tokens, end) + 1;
	}
	return Math.min(end, tokens.length);
}

export function isReferenceOperator(t: FormulaToken | undefined): boolean {
	return (
		t !== undefined &&
		(t.type === 'intersect' || t.type === 'union' || (t.type === 'operator' && t.text === ':'))
	);
}

/**
 * Index just past the operand of a unary operator whose operand starts at
 * `i`: nested prefixes, a primary and any reference operators (`:`, ` `,
 * `,`) that bind tighter than the unary operator.
 */
export function unaryOperandEnd(tokens: readonly FormulaToken[], i: number): number {
	let k = skipWhitespace(tokens, i);
	while (k < tokens.length && tokens[k].type === 'prefix') {
		k = skipWhitespace(tokens, k + 1);
	}
	let end = primaryEnd(tokens, k);
	for (;;) {
		const next = skipWhitespace(tokens, end);
		if (end === k || !isReferenceOperator(tokens[next])) {
			return end;
		}
		const operand = skipWhitespace(tokens, next + 1);
		const after = primaryEnd(tokens, operand);
		if (after === operand) {
			return end;
		}
		end = after;
	}
}

export interface FormulaCallInfo {
	/** Index of the function token, or -1 for a call on an expression (`LAMBDA(x,x)(1)`). */
	fn: number;
	/** Canonical function name (`''` for calls on expressions). */
	name: string;
	/** Index of the `call-open` token. */
	open: number;
	/** Index of the `call-close` token, or -1 when the call is not closed yet. */
	close: number;
	/** Token index ranges `[start, end)` of each argument, whitespace included. */
	args: Array<[number, number]>;
}

type Frame = { call: FormulaCallInfo; argStart: number } | 'group' | 'array';

function finishCall(
	tokens: readonly FormulaToken[],
	frame: { call: FormulaCallInfo; argStart: number },
	end: number,
): void {
	const { call, argStart } = frame;
	if (call.args.length > 0 || skipWhitespace(tokens, argStart) < end) {
		call.args.push([argStart, end]);
	}
}

/** Every function call in the token stream, in order of their opening parenthesis. */
export function analyzeCalls(tokens: readonly FormulaToken[]): FormulaCallInfo[] {
	const calls: FormulaCallInfo[] = [];
	const stack: Frame[] = [];
	for (let i = 0; i < tokens.length; i++) {
		const t = tokens[i];
		const top = stack[stack.length - 1];
		switch (t.type) {
			case 'call-open': {
				const fn = tokens[i - 1]?.type === 'function' ? i - 1 : -1;
				const call: FormulaCallInfo = {
					fn,
					name: fn >= 0 ? normalizeFunctionName(tokens[fn].text) : '',
					open: i,
					close: -1,
					args: [],
				};
				calls.push(call);
				stack.push({ call, argStart: i + 1 });
				break;
			}
			case 'paren-open':
				stack.push('group');
				break;
			case 'array-open':
				stack.push('array');
				break;
			case 'arg-sep':
				if (typeof top === 'object') {
					top.call.args.push([top.argStart, i]);
					top.argStart = i + 1;
				}
				break;
			case 'call-close':
				if (typeof top === 'object') {
					finishCall(tokens, top, i);
					top.call.close = i;
					stack.pop();
				}
				break;
			case 'paren-close':
				if (top === 'group') {
					stack.pop();
				}
				break;
			case 'array-close':
				if (top === 'array') {
					stack.pop();
				}
				break;
		}
	}
	for (const frame of stack) {
		if (typeof frame === 'object') {
			finishCall(tokens, frame, tokens.length);
		}
	}
	return calls;
}

/** The only significant token in `[start, end)`, or -1. */
export function soleToken(tokens: readonly FormulaToken[], start: number, end: number): number {
	const first = skipWhitespace(tokens, start);
	if (first >= end) {
		return -1;
	}
	return skipWhitespace(tokens, first + 1) >= end ? first : -1;
}
