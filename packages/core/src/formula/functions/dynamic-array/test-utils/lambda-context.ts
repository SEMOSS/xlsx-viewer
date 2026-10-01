import { ERRORS } from '../../../../values/errors';
import type { AstNode } from '../../../ast';
import { MockContext } from '../../../test-utils/mock-context';
import type { FormulaValue, LambdaValue } from '../../../types';

/**
 * MockContext with LAMBDA support for tests: lambdas are backed by JS
 * functions and names evaluate through a binding table (`#NAME?` otherwise).
 */
export class LambdaContext extends MockContext {
	private readonly impls = new Map<LambdaValue, (args: FormulaValue[]) => FormulaValue>();
	readonly bindings = new Map<string, FormulaValue>();

	lambda(params: string[], impl: (args: FormulaValue[]) => FormulaValue): LambdaValue {
		const fn: LambdaValue = { kind: 'lambda', params, body: { type: 'missing' }, scope: new Map() };
		this.impls.set(fn, impl);
		return fn;
	}

	override callLambda(fn: LambdaValue, args: FormulaValue[]): FormulaValue {
		const impl = this.impls.get(fn);
		return impl ? impl(args) : ERRORS.CALC;
	}

	override evaluate(node: AstNode): FormulaValue {
		if (node.type === 'name') {
			return this.bindings.get(node.name.toUpperCase()) ?? ERRORS.NAME;
		}
		return super.evaluate(node);
	}
}

/** AST node for an eta-reduced function or a bound name, e.g. `name('SUM')`. */
export function name(text: string): AstNode {
	return { type: 'name', name: text };
}
