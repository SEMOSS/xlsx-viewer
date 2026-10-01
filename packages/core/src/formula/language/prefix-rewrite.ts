import { filePrefixFor } from './function-prefixes';
import { stripFilePrefixes } from './names';
import {
	formatStructuredReference,
	parseStructuredReference,
	usesThisRowShorthand,
} from './structured';
import {
	analyzeCalls,
	applyEdits,
	type FormulaCallInfo,
	hasInvalidTokens,
	isReferenceOperator,
	matchingClose,
	previousSignificant,
	skipWhitespace,
	soleToken,
	type TextEdit,
	unaryOperandEnd,
} from './token-utils';
import { tokenizeFormula } from './tokenizer';
import type { FormulaToken } from './tokens';

/**
 * Token-based conversion between the file form of a formula and the form
 * shown to users. Only the affected tokens are rewritten, so spacing and
 * line breaks survive. Both functions return `null` when the formula does
 * not tokenize cleanly; the caller then falls back to a textual approach.
 */

/** `_xlfn.SINGLE(x)` -> `@x`, wrapping where the `@` would otherwise bind differently. */
function stripSingle(
	tokens: readonly FormulaToken[],
	call: FormulaCallInfo,
	edits: TextEdit[],
): void {
	const [argStart, argEnd] = call.args[0];
	const first = skipWhitespace(tokens, argStart);
	const simple = skipWhitespace(tokens, unaryOperandEnd(tokens, first)) >= argEnd;
	const prev = tokens[previousSignificant(tokens, call.fn)];
	const next = tokens[skipWhitespace(tokens, call.close + 1)];
	const wrap = isReferenceOperator(prev) || isReferenceOperator(next) || next?.type === 'spill';
	const fn = tokens[call.fn];
	const close = tokens[call.close];
	edits.push({
		start: fn.start,
		end: tokens[call.open].end,
		text: (wrap ? '(' : '') + (simple ? '@' : '@('),
	});
	edits.push({ start: close.start, end: close.end, text: (simple ? '' : ')') + (wrap ? ')' : '') });
}

/** `_xlfn.ANCHORARRAY(A1)` -> `A1#` when the argument is a single cell or name. */
function stripAnchorArray(
	tokens: readonly FormulaToken[],
	call: FormulaCallInfo,
	edits: TextEdit[],
): boolean {
	const [argStart, argEnd] = call.args[0];
	const only = soleToken(tokens, argStart, argEnd);
	if (only !== argStart || argEnd !== only + 1) {
		return false;
	}
	const arg = tokens[only];
	if (arg.type !== 'name' && !(arg.type === 'ref' && arg.area?.kind === 'cell')) {
		return false;
	}
	edits.push({ start: tokens[call.fn].start, end: tokens[call.open].end, text: '' });
	edits.push({ start: tokens[call.close].start, end: tokens[call.close].end, text: '#' });
	return true;
}

/** File form -> display form; `null` when the formula does not tokenize cleanly. */
export function stripPrefixesWithTokens(formula: string): string | null {
	const tokens = tokenizeFormula(formula);
	if (hasInvalidTokens(tokens)) {
		return null;
	}
	const edits: TextEdit[] = [];
	let callsByFn: Map<number, FormulaCallInfo> | undefined;
	for (let i = 0; i < tokens.length; i++) {
		const t = tokens[i];
		if (t.type !== 'function' && (t.type !== 'name' || t.sheet !== undefined)) {
			continue;
		}
		const stripped = stripFilePrefixes(t.text);
		if (t.type === 'function') {
			const upper = stripped.toUpperCase();
			if (upper === 'SINGLE' || upper === 'ANCHORARRAY') {
				callsByFn ??= new Map(analyzeCalls(tokens).map((c) => [c.fn, c]));
				const call = callsByFn.get(i);
				if (call !== undefined && call.close >= 0 && call.args.length === 1) {
					if (upper === 'SINGLE' && skipWhitespace(tokens, call.args[0][0]) < call.args[0][1]) {
						stripSingle(tokens, call, edits);
						continue;
					}
					if (upper === 'ANCHORARRAY' && stripAnchorArray(tokens, call, edits)) {
						continue;
					}
				}
			}
		}
		if (stripped !== t.text) {
			edits.push({ start: t.start, end: t.end, text: stripped });
		}
	}
	return applyEdits(formula, edits);
}

interface Binding {
	name: string;
	/** Token range `[from, to)` where the parameter is visible. */
	from: number;
	to: number;
	/** Calls (`f(x)`) may refer to it: LAMBDA parameters and LET names bound to a LAMBDA. */
	callable: boolean;
}

function isLambdaValue(tokens: readonly FormulaToken[], range: [number, number]): boolean {
	const first = skipWhitespace(tokens, range[0]);
	return (
		first < range[1] &&
		tokens[first].type === 'function' &&
		stripFilePrefixes(tokens[first].text).toUpperCase() === 'LAMBDA'
	);
}

/** Token indices of LET / LAMBDA parameter declarations and their uses (they take `_xlpm.`). */
function parameterTokens(
	tokens: readonly FormulaToken[],
	calls: readonly FormulaCallInfo[],
): Set<number> {
	const bound = new Set<number>();
	const bindings: Binding[] = [];
	for (const call of calls) {
		const isLet = call.name === 'LET';
		if (!isLet && call.name !== 'LAMBDA') {
			continue;
		}
		const { args } = call;
		const last = args.length - 1;
		const to = call.close >= 0 ? call.close : tokens.length;
		for (let k = 0; k < last; k += isLet ? 2 : 1) {
			const idx = soleToken(tokens, args[k][0], args[k][1]);
			if (idx < 0 || tokens[idx].type !== 'name' || tokens[idx].sheet !== undefined) {
				continue;
			}
			bound.add(idx);
			const name = stripFilePrefixes(tokens[idx].text).toUpperCase();
			if (isLet) {
				// A LET name is visible from the argument after its value.
				const from = k + 2 <= last ? args[k + 2][0] : to;
				bindings.push({ name, from, to, callable: isLambdaValue(tokens, args[k + 1]) });
			} else {
				bindings.push({ name, from: args[last][0], to, callable: true });
			}
		}
	}
	if (bindings.length === 0) {
		return bound;
	}
	for (let i = 0; i < tokens.length; i++) {
		const t = tokens[i];
		if ((t.type !== 'name' && t.type !== 'function') || t.sheet !== undefined || bound.has(i)) {
			continue;
		}
		const upper = stripFilePrefixes(t.text).toUpperCase();
		for (const b of bindings) {
			if (b.name === upper && i >= b.from && i < b.to && (t.type === 'name' || b.callable)) {
				bound.add(i);
				break;
			}
		}
	}
	return bound;
}

/** `@x` -> `_xlfn.SINGLE(x)`; `@(x)` -> `_xlfn.SINGLE(x)`. */
function addSingle(tokens: readonly FormulaToken[], at: number, edits: TextEdit[]): void {
	const op = tokens[at];
	const first = skipWhitespace(tokens, at + 1);
	const end = unaryOperandEnd(tokens, at + 1);
	if (
		first === at + 1 &&
		tokens[first]?.type === 'paren-open' &&
		matchingClose(tokens, first) === end - 1
	) {
		edits.push({ start: op.start, end: tokens[first].end, text: '_xlfn.SINGLE(' });
		return;
	}
	// A union in the operand must stay one argument: `@A1,B1` -> `_xlfn.SINGLE((A1,B1))`.
	let depth = 0;
	let union = false;
	for (let k = first; k < end; k++) {
		const type = tokens[k].type;
		if (type === 'paren-open' || type === 'call-open' || type === 'array-open') {
			depth++;
		} else if (type === 'paren-close' || type === 'call-close' || type === 'array-close') {
			depth--;
		} else if (type === 'union' && depth === 0) {
			union = true;
		}
	}
	const insertAt = end > first ? tokens[end - 1].end : op.end;
	edits.push({ start: op.start, end: op.end, text: union ? '_xlfn.SINGLE((' : '_xlfn.SINGLE(' });
	edits.push({ start: insertAt, end: insertAt, text: union ? '))' : ')' });
}

function withParamPrefix(text: string): string {
	return /^_xlpm\./i.test(text) ? text : `_xlpm.${stripFilePrefixes(text)}`;
}

/** Display form -> file form; `null` when the formula does not tokenize cleanly. */
export function addPrefixesWithTokens(formula: string): string | null {
	const tokens = tokenizeFormula(formula);
	if (hasInvalidTokens(tokens)) {
		return null;
	}
	const bound = parameterTokens(tokens, analyzeCalls(tokens));
	const edits: TextEdit[] = [];
	for (let i = 0; i < tokens.length; i++) {
		const t = tokens[i];
		let text: string | undefined;
		switch (t.type) {
			case 'function': {
				const stripped = stripFilePrefixes(t.text);
				text = bound.has(i) ? withParamPrefix(t.text) : filePrefixFor(stripped) + stripped;
				break;
			}
			case 'name':
				if (bound.has(i)) {
					text = withParamPrefix(t.text);
				}
				break;
			case 'spill':
				edits.push({
					start: tokens[i - 1].start,
					end: tokens[i - 1].start,
					text: '_xlfn.ANCHORARRAY(',
				});
				text = ')';
				break;
			case 'prefix':
				if (t.text === '@') {
					addSingle(tokens, i, edits);
				}
				break;
			case 'structured':
				if (usesThisRowShorthand(t.text)) {
					const node = parseStructuredReference(t.text);
					text = node === null ? undefined : formatStructuredReference(node, { forFile: true });
				}
				break;
		}
		if (text !== undefined && text !== t.text) {
			edits.push({ start: t.start, end: t.end, text });
		}
	}
	return applyEdits(formula, edits);
}
