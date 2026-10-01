import { ERRORS } from '../../../values/errors';
import type { AstNode } from '../../ast';
import { isOmittedValue, paramKey } from '../../evaluator/lambda';
import type { FormulaValue, FunctionArg, FunctionDefinition } from '../../types';
import { isMissingNode } from '../logical/lift';
import { withFilePrefix } from '../logical/prefix';

/**
 * LET, LAMBDA and ISOMITTED. Names are bound upper case. LAMBDA captures
 * the bindings in effect where it is created; optional parameters are
 * written `[name]` and stored as `[NAME]` in the LAMBDA value.
 */

function asNode(arg: FunctionArg): AstNode | null {
	return typeof arg === 'object' && arg !== null && 'type' in arg ? (arg as AstNode) : null;
}

/** Name bound by a LET / LAMBDA argument, or null when the argument is not a plain name. */
function bindingName(arg: FunctionArg): string | null {
	const node = asNode(arg);
	return node !== null && node.type === 'name' && node.sheet === undefined
		? node.name.toUpperCase()
		: null;
}

/** Parameter spelling: `X` for a required parameter, `[X]` for an optional one. */
function parameterSpelling(arg: FunctionArg): string | null {
	const name = bindingName(arg);
	if (name !== null) {
		return name;
	}
	const node = asNode(arg);
	if (
		node !== null &&
		node.type === 'structured' &&
		node.table === undefined &&
		node.specifiers.length === 0 &&
		node.columnStart !== undefined &&
		node.columnEnd === undefined &&
		/^[A-Za-z_\\][A-Za-z0-9_.]*$/.test(node.columnStart)
	) {
		return `[${node.columnStart.toUpperCase()}]`;
	}
	return null;
}

const LET: FunctionDefinition = withFilePrefix({
	name: 'LET',
	minArgs: 3,
	maxArgs: 253,
	params: ['lazy'],
	impl: (args, ctx) => {
		if (args.length % 2 === 0) {
			return ERRORS.VALUE;
		}
		const bindings = new Map<string, FormulaValue>();
		for (let i = 0; i + 1 < args.length; i += 2) {
			const name = bindingName(args[i]);
			if (name === null) {
				return ERRORS.VALUE;
			}
			bindings.set(name, ctx.evaluateWith(args[i + 1] as AstNode, bindings));
		}
		return ctx.evaluateWith(args[args.length - 1] as AstNode, bindings);
	},
});

const LAMBDA: FunctionDefinition = withFilePrefix({
	name: 'LAMBDA',
	minArgs: 1,
	maxArgs: 254,
	params: ['lazy'],
	impl: (args, ctx) => {
		const body = asNode(args[args.length - 1]);
		if (body === null || body.type === 'missing') {
			return ERRORS.VALUE;
		}
		const params: string[] = [];
		const seen = new Set<string>();
		let optionalSeen = false;
		for (let i = 0; i < args.length - 1; i++) {
			const p = parameterSpelling(args[i]);
			if (p === null) {
				return ERRORS.VALUE;
			}
			const optional = p.startsWith('[');
			const key = paramKey(p);
			if (seen.has(key) || (optionalSeen && !optional)) {
				return ERRORS.VALUE;
			}
			optionalSeen ||= optional;
			seen.add(key);
			params.push(p);
		}
		return { kind: 'lambda', params, body, scope: ctx.scope };
	},
});

/** ISOMITTED(argument): the LAMBDA parameter was not supplied by the caller. */
const ISOMITTED: FunctionDefinition = withFilePrefix({
	name: 'ISOMITTED',
	minArgs: 1,
	maxArgs: 1,
	params: ['lazy'],
	impl: (args, ctx) => {
		const name = bindingName(args[0]);
		if (name !== null) {
			const bound = ctx.scope.get(name);
			if (bound !== undefined) {
				return isOmittedValue(bound);
			}
		}
		return isMissingNode(args[0]);
	},
});

export const LET_LAMBDA_FUNCTIONS: FunctionDefinition[] = [LET, LAMBDA, ISOMITTED];
