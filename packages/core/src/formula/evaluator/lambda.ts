import type { AstNode, FunctionNode, NameNode } from '../ast';
import type { FormulaValue, FunctionDefinition, LambdaValue } from '../types';

/**
 * LAMBDA plumbing shared by the evaluator and the LAMBDA function module:
 * the marker bound to omitted parameters, optional-parameter spelling and
 * eta-reduced LAMBDAs (a bare function name such as `SUM` passed to MAP,
 * BYROW, REDUCE, ...).
 */

/**
 * Value bound to a parameter the caller omitted (`ISOMITTED` tests for it).
 * Reading the parameter yields a blank. It has the shape of a `missing`
 * AST node so function argument helpers treat it as an empty argument.
 */
export const OMITTED: FormulaValue = Object.freeze({ type: 'missing' }) as unknown as FormulaValue;

/** The binding is the omitted-parameter marker. */
export function isOmittedValue(v: unknown): boolean {
	return v === OMITTED;
}

/** Optional parameters (`LAMBDA(a, [b], ...)`) are stored as `[B]` in `LambdaValue.params`. */
export function optionalParamName(param: string): string | null {
	return param.length > 2 && param.charCodeAt(0) === 91 /* [ */ && param.endsWith(']')
		? param.slice(1, -1)
		: null;
}

/** Scope key of a parameter (upper case, brackets of optional parameters removed). */
export function paramKey(param: string): string {
	return (optionalParamName(param) ?? param).toUpperCase();
}

const ETA_DEFS = new WeakMap<LambdaValue, FunctionDefinition>();
const ETA_BY_NAME = new Map<string, LambdaValue>();
const ETA_BODIES = new Map<string, FunctionNode>();

function etaParam(i: number): string {
	return `\u0001ETA${i}`;
}

/** Call body `NAME(p0, ..., pn-1)` for an eta-reduced function called with `n` arguments. */
export function etaBody(def: FunctionDefinition, n: number): FunctionNode {
	const key = `${def.name}/${n}`;
	let body = ETA_BODIES.get(key);
	if (body === undefined) {
		const args: AstNode[] = [];
		for (let i = 0; i < n; i++) {
			args.push({ type: 'name', name: etaParam(i) } satisfies NameNode);
		}
		body = { type: 'function', name: def.name, args };
		ETA_BODIES.set(key, body);
	}
	return body;
}

/** Scope binding the arguments of an eta-reduced call. */
export function etaScope(args: readonly FormulaValue[]): Map<string, FormulaValue> {
	const scope = new Map<string, FormulaValue>();
	args.forEach((a, i) => scope.set(etaParam(i), a));
	return scope;
}

/** The LAMBDA a bare function name stands for (`MAP(A1:A3, ABS)`). */
export function etaLambda(def: FunctionDefinition): LambdaValue {
	let fn = ETA_BY_NAME.get(def.name);
	if (fn === undefined || ETA_DEFS.get(fn) !== def) {
		const n = Math.max(1, Math.min(def.minArgs, 8));
		const params: string[] = [];
		for (let i = 0; i < n; i++) {
			params.push(etaParam(i));
		}
		fn = { kind: 'lambda', params, body: etaBody(def, n), scope: new Map() };
		ETA_DEFS.set(fn, def);
		ETA_BY_NAME.set(def.name, fn);
	}
	return fn;
}

/** The function behind an eta-reduced LAMBDA, if `fn` is one. */
export function etaDefinition(fn: LambdaValue): FunctionDefinition | undefined {
	return ETA_DEFS.get(fn);
}
