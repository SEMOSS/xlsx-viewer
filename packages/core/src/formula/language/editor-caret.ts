import type { RefCorner } from '../ast';
import { isNameChar, isNameStart } from './chars';
import { formatRefArea } from './references';
import { analyzeCalls, previousSignificant } from './token-utils';
import { tokenizeFormula } from './tokenizer';
import type { FormulaToken, FormulaTokenType, RefArea } from './tokens';

/** The function call whose argument list contains the caret (for argument tooltips). */
export interface ActiveFunctionCall {
	/** Canonical (upper-case, unprefixed) function name. */
	name: string;
	/** 0-based index of the argument the caret is in. */
	argIndex: number;
	/** Offset of the function name in the formula. */
	nameStart: number;
}

/** Identifier being typed before the caret, for function / name completion. */
export interface CompletionPrefix {
	prefix: string;
	/** Offset where the identifier starts (replace `[start, caret)` with the completion). */
	start: number;
}

export interface ReferenceToggle {
	formula: string;
	caret: number;
}

/**
 * Innermost named function call around `caret` (between its `(` and `)`,
 * or after an unclosed `(`), with the index of the current argument.
 * Returns `null` outside any call or inside a call on an expression.
 */
export function getActiveFunctionCall(formula: string, caret: number): ActiveFunctionCall | null {
	const tokens = tokenizeFormula(formula);
	const calls = analyzeCalls(tokens);
	for (let c = calls.length - 1; c >= 0; c--) {
		const call = calls[c];
		const open = tokens[call.open];
		const inside = open.end <= caret && (call.close < 0 || caret <= tokens[call.close].start);
		if (!inside) {
			continue;
		}
		if (call.fn < 0) {
			return null;
		}
		let argIndex = 0;
		for (const [, end] of call.args) {
			const sep = tokens[end];
			if (sep !== undefined && sep.type === 'arg-sep' && sep.end <= caret) {
				argIndex++;
			}
		}
		return { name: call.name, argIndex, nameStart: tokens[call.fn].start };
	}
	return null;
}

/** Tokens after which an operand (a reference, name or function) may start. */
function expectsOperand(type: FormulaTokenType | undefined): boolean {
	switch (type) {
		case undefined:
		case 'equals':
		case 'operator':
		case 'prefix':
		case 'paren-open':
		case 'call-open':
		case 'arg-sep':
		case 'union':
		case 'intersect':
			return true;
		default:
			return false;
	}
}

/**
 * The identifier ending at `caret` when it sits where a function or name
 * may be typed (not inside a string, a sheet-qualified reference or a table
 * reference), else `null`. A caret in the middle of an identifier offers
 * no completion.
 */
export function getCompletionPrefix(formula: string, caret: number): CompletionPrefix | null {
	if (caret <= 0 || caret > formula.length || isNameChar(formula.charCodeAt(caret))) {
		return null;
	}
	let start = caret;
	while (start > 0 && isNameChar(formula.charCodeAt(start - 1))) {
		start--;
	}
	if (start === caret || !isNameStart(formula.charCodeAt(start))) {
		return null;
	}
	const tokens = tokenizeFormula(formula.slice(0, caret));
	const index = tokens.findIndex((t) => t.start <= start && start < t.end);
	const token = tokens[index];
	if (token === undefined || token.start !== start || token.sheet !== undefined) {
		return null;
	}
	if (token.type !== 'name' && token.type !== 'ref' && token.type !== 'boolean') {
		return null;
	}
	const prev = tokens[previousSignificant(tokens, index)];
	if (!expectsOperand(prev?.type)) {
		return null;
	}
	return { prefix: formula.slice(start, caret), start };
}

/**
 * The caret is where clicking a cell should insert a reference ("point
 * mode"): right after `=`, an operator, `(`, or an argument separator,
 * ignoring whitespace, and not inside a token.
 */
export function isPointModeAllowed(formula: string, caret: number): boolean {
	if (caret <= 0 || caret > formula.length) {
		return false;
	}
	const tokens = tokenizeFormula(formula);
	let prev: FormulaToken | undefined;
	for (const t of tokens) {
		if (t.end > caret) {
			if (t.start < caret && t.type !== 'whitespace' && t.type !== 'intersect') {
				return false;
			}
			break;
		}
		if (t.type !== 'whitespace' && t.type !== 'intersect') {
			prev = t;
		}
	}
	if (prev === undefined) {
		return false;
	}
	switch (prev.type) {
		case 'equals':
		case 'operator':
		case 'prefix':
		case 'paren-open':
		case 'call-open':
		case 'arg-sep':
		case 'union':
			return true;
		default:
			return false;
	}
}

function cycleCorner(c: RefCorner, colAbs: boolean, rowAbs: boolean): RefCorner {
	return { row: c.row, col: c.col, rowAbs, colAbs };
}

/** Next F4 state: A1 -> $A$1 -> A$1 -> $A1 -> A1 (whole rows / columns just toggle). */
function nextAbsolute(area: RefArea): RefArea {
	const { start, end } = area;
	if (area.kind === 'columns') {
		const abs = !start.colAbs;
		return {
			kind: 'columns',
			start: cycleCorner(start, abs, false),
			end: cycleCorner(end, abs, false),
		};
	}
	if (area.kind === 'rows') {
		const abs = !start.rowAbs;
		return {
			kind: 'rows',
			start: cycleCorner(start, false, abs),
			end: cycleCorner(end, false, abs),
		};
	}
	let colAbs: boolean;
	let rowAbs: boolean;
	if (!start.colAbs && !start.rowAbs) {
		[colAbs, rowAbs] = [true, true];
	} else if (start.colAbs && start.rowAbs) {
		[colAbs, rowAbs] = [false, true];
	} else if (start.rowAbs) {
		[colAbs, rowAbs] = [true, false];
	} else {
		[colAbs, rowAbs] = [false, false];
	}
	const first = cycleCorner(start, colAbs, rowAbs);
	if (area.kind === 'cell') {
		return { kind: 'cell', start: first, end: first };
	}
	return { kind: 'cells', start: first, end: cycleCorner(end, colAbs, rowAbs) };
}

/**
 * F4: cycle the absolute / relative state of the reference at (or just
 * before) the caret. A range changes as a whole, following its first
 * corner. The caret moves to the end of the reference. Without a reference
 * at the caret the input is returned unchanged.
 */
export function toggleReferenceAbsolute(formula: string, caret: number): ReferenceToggle {
	for (const t of tokenizeFormula(formula)) {
		if (t.start > caret) {
			break;
		}
		if (t.type !== 'ref' || t.area === undefined || caret > t.end) {
			continue;
		}
		const text = t.text.slice(0, t.qualifierLength ?? 0) + formatRefArea(nextAbsolute(t.area));
		return {
			formula: formula.slice(0, t.start) + text + formula.slice(t.end),
			caret: t.start + text.length,
		};
	}
	return { formula, caret };
}
