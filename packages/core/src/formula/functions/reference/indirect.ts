import { ERRORS, isCellError } from '../../../values/errors';
import type { AstNode } from '../../ast';
import { toBoolean, toText } from '../../coerce';
import { parseFormula } from '../../language/parser';
import type { FormulaValue, FunctionDefinition, Scalar } from '../../types';
import { isReference } from '../../types';
import { isOmittedArg } from './shared';

/**
 * INDIRECT(ref_text, [a1]): the reference named by a text, in A1 or (a1 =
 * FALSE) R1C1 notation relative to the formula cell. Sheet-qualified
 * references, defined names and table references are accepted; anything
 * that does not denote a reference is `#REF!`. Volatile.
 */

const MAX_CACHE = 2000;
const cache = new Map<string, AstNode | null>();

/** The parsed text denotes a reference (not a computation). */
function isReferenceSyntax(node: AstNode): boolean {
	switch (node.type) {
		case 'cell':
		case 'range':
		case 'name':
		case 'structured':
		case 'spill':
			return true;
		case 'binary':
			return (
				(node.op === ':' || node.op === ',' || node.op === ' ') &&
				isReferenceSyntax(node.left) &&
				isReferenceSyntax(node.right)
			);
		default:
			return false;
	}
}

function parseReference(text: string, a1: boolean, row: number, col: number): AstNode | null {
	const key = a1 ? `A|${text}` : `R${row},${col}|${text}`;
	let ast = cache.get(key);
	if (ast === undefined) {
		try {
			const parsed = parseFormula(text, a1 ? undefined : { r1c1: { row, col } });
			ast = isReferenceSyntax(parsed) ? parsed : null;
		} catch {
			ast = null;
		}
		if (cache.size >= MAX_CACHE) {
			cache.clear();
		}
		cache.set(key, ast);
	}
	return ast;
}

const INDIRECT: FunctionDefinition = {
	name: 'INDIRECT',
	minArgs: 1,
	maxArgs: 2,
	params: ['value', 'value'],
	volatile: true,
	impl: (args, ctx) => {
		const text = toText(isOmittedArg(args, 0) ? '' : (args[0] as Scalar));
		if (isCellError(text)) {
			return text;
		}
		const a1 = isOmittedArg(args, 1) ? true : toBoolean(args[1] as Scalar);
		if (isCellError(a1)) {
			return a1;
		}
		const trimmed = text.trim();
		if (trimmed === '' || trimmed.startsWith('=')) {
			return ERRORS.REF;
		}
		const ast = parseReference(trimmed, a1, ctx.row, ctx.col);
		if (ast === null) {
			return ERRORS.REF;
		}
		// Parsed text holds absolute positions: evaluate it outside any defined-name offset.
		const detached = (ctx as { evaluateDetached?: (node: AstNode) => FormulaValue })
			.evaluateDetached;
		const v = detached ? detached.call(ctx, ast) : ctx.evaluate(ast);
		if (isReference(v)) {
			return v;
		}
		return isCellError(v) && v.code !== '#NAME?' ? v : ERRORS.REF;
	},
};

export const INDIRECT_FUNCTIONS: FunctionDefinition[] = [INDIRECT];
