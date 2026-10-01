import type {
	AstNode,
	BinaryOperator,
	FunctionNode,
	NameNode,
	SheetQualifier,
	StructuredRefNode,
} from '../ast';
import { filePrefixFor } from './function-prefixes';
import { formatRefArea, formatRefCorner, formatSheetQualifier } from './references';
import { formatStructuredReference, usesThisRowShorthand } from './structured';

export interface SerializeOptions {
	/**
	 * Emit the file (OOXML) form: `_xlfn.` / `_xlfn._xlws.` prefixes on newer
	 * functions, `_xlpm.` on LET / LAMBDA parameters, `A1#` as
	 * `_xlfn.ANCHORARRAY(A1)`, `@x` as `_xlfn.SINGLE(x)` and `[#This Row]`
	 * instead of the `@` table shorthand.
	 */
	forFile?: boolean;
}

const ATOM = 11;
const PREC_UNARY = 7;
const PREC_PERCENT = 6;

const BINARY_PRECEDENCE: Record<BinaryOperator, number> = {
	':': 10,
	' ': 9,
	',': 8,
	'^': 5,
	'*': 4,
	'/': 4,
	'+': 3,
	'-': 3,
	'&': 2,
	'=': 1,
	'<>': 1,
	'<': 1,
	'>': 1,
	'<=': 1,
	'>=': 1,
};

/** Number literal text: shortest round-trip digits, `E` exponents (`1E+21`, `1.5E-7`). */
export function formatFormulaNumber(n: number): string {
	if (!Number.isFinite(n)) {
		return '#NUM!';
	}
	if (n === 0) {
		return '0';
	}
	const s = String(n);
	const e = s.indexOf('e');
	if (e < 0) {
		return s;
	}
	const exp = s.slice(e + 1);
	return `${s.slice(0, e)}E${exp[0] === '-' || exp[0] === '+' ? exp : `+${exp}`}`;
}

/** String literal with embedded quotes doubled. */
export function formatFormulaString(value: string): string {
	return `"${value.replace(/"/g, '""')}"`;
}

/**
 * Print an AST as formula text (no leading `=`), adding only the
 * parentheses the grammar needs plus those recorded in `parens` by the
 * parser, so `serializeFormula(parseFormula(f))` is equivalent to `f`.
 */
export function serializeFormula(node: AstNode, options: SerializeOptions = {}): string {
	return new Printer(options.forFile === true).expr(node, 0, false);
}

class Printer {
	/** LET / LAMBDA parameters in scope, innermost last. */
	private readonly scope: Array<{ upper: string; spelling: string; callable: boolean }> = [];

	constructor(private readonly forFile: boolean) {}

	private precedence(node: AstNode): number {
		switch (node.type) {
			case 'binary':
				return BINARY_PRECEDENCE[node.op];
			case 'unary':
				return node.op === '@' && this.forFile ? ATOM : PREC_UNARY;
			case 'percent':
				return PREC_PERCENT;
			case 'number':
				return node.value < 0 ? PREC_UNARY : ATOM;
			default:
				return ATOM;
		}
	}

	/**
	 * Print `node` where the context needs precedence `min`. In an argument
	 * list (`inArgs`) a bare union would read as an argument separator.
	 */
	expr(node: AstNode, min: number, inArgs: boolean): string {
		const own = node.parens ?? 0;
		if (own > 0) {
			return '('.repeat(own) + this.bare(node, false) + ')'.repeat(own);
		}
		if (this.precedence(node) < min || (inArgs && node.type === 'binary' && node.op === ',')) {
			return `(${this.bare(node, false)})`;
		}
		return this.bare(node, inArgs);
	}

	private bare(node: AstNode, inArgs: boolean): string {
		switch (node.type) {
			case 'number':
				return formatFormulaNumber(node.value);
			case 'string':
				return formatFormulaString(node.value);
			case 'boolean':
				return node.value ? 'TRUE' : 'FALSE';
			case 'error':
				return this.qualifier(node.sheet) + node.code;
			case 'cell':
				return this.qualifier(node.sheet) + formatRefCorner(node.ref);
			case 'range':
				return this.qualifier(node.sheet) + formatRefArea(node);
			case 'name':
				return this.name(node);
			case 'structured':
				return this.structured(node);
			case 'function':
				return this.call(node);
			case 'call': {
				const callable =
					node.callee.type === 'function' ||
					node.callee.type === 'call' ||
					node.callee.type === 'name';
				const callee = this.expr(node.callee, callable ? 0 : ATOM + 1, false);
				return `${callee}(${this.args(node.args)})`;
			}
			case 'unary':
				if (node.op === '@' && this.forFile) {
					return `_xlfn.SINGLE(${this.singleArgument(node.operand)})`;
				}
				return node.op + this.expr(node.operand, PREC_UNARY, inArgs);
			case 'percent':
				return `${this.expr(node.operand, PREC_PERCENT, inArgs)}%`;
			case 'binary': {
				const p = BINARY_PRECEDENCE[node.op];
				return this.expr(node.left, p, inArgs) + node.op + this.expr(node.right, p + 1, inArgs);
			}
			case 'array':
				return `{${node.rows.map((r) => r.map((v) => this.bare(v, false)).join(',')).join(';')}}`;
			case 'missing':
				return '';
			case 'spill': {
				const ref = this.bare(node.ref, false);
				return this.forFile ? `_xlfn.ANCHORARRAY(${ref})` : `${ref}#`;
			}
		}
	}

	/** Operand of `_xlfn.SINGLE(...)`: the call's parentheses replace one written pair (`@(x)`). */
	private singleArgument(operand: AstNode): string {
		const own = operand.parens ?? 0;
		const union = operand.type === 'binary' && operand.op === ',';
		if (own === 0 || (own === 1 && union)) {
			return this.expr(operand, 0, true);
		}
		return '('.repeat(own - 1) + this.bare(operand, false) + ')'.repeat(own - 1);
	}

	private qualifier(sheet: SheetQualifier | undefined): string {
		return sheet === undefined ? '' : formatSheetQualifier(sheet);
	}

	/** Spelling of the innermost parameter called `name` (for calls: one that may hold a LAMBDA). */
	private lookup(name: string, asCall: boolean): string | undefined {
		const upper = name.toUpperCase();
		for (let k = this.scope.length - 1; k >= 0; k--) {
			const entry = this.scope[k];
			if (entry.upper === upper) {
				return !asCall || entry.callable ? entry.spelling : undefined;
			}
		}
		return undefined;
	}

	private bind(name: string, callable: boolean): void {
		this.scope.push({ upper: name.toUpperCase(), spelling: name, callable });
	}

	private param(name: string): string {
		return this.forFile ? `_xlpm.${name}` : name;
	}

	private name(node: NameNode): string {
		if (node.sheet !== undefined) {
			return formatSheetQualifier(node.sheet) + node.name;
		}
		return this.lookup(node.name, false) !== undefined ? this.param(node.name) : node.name;
	}

	private structured(node: StructuredRefNode): string {
		if (this.forFile) {
			return node.text && !usesThisRowShorthand(node.text)
				? node.text
				: formatStructuredReference(node, { forFile: true });
		}
		return node.text || formatStructuredReference(node);
	}

	private args(args: readonly AstNode[]): string {
		return args.map((a) => this.expr(a, 0, true)).join(',');
	}

	private functionName(name: string): string {
		const bound = this.lookup(name, true);
		if (bound !== undefined) {
			return this.param(bound);
		}
		return this.forFile ? filePrefixFor(name) + name : name;
	}

	private call(node: FunctionNode): string {
		const { name, args } = node;
		if (name === 'LET' || name === 'LAMBDA') {
			return `${this.functionName(name)}(${this.binding(node)})`;
		}
		if (
			!this.forFile &&
			name === 'ANCHORARRAY' &&
			args.length === 1 &&
			args[0].type === 'name' &&
			!args[0].parens
		) {
			return `${this.name(args[0])}#`;
		}
		return `${this.functionName(name)}(${this.args(args)})`;
	}

	/** LET / LAMBDA arguments with their parameters in scope where Excel binds them. */
	private binding(node: FunctionNode): string {
		const { args } = node;
		const last = args.length - 1;
		const isLet = node.name === 'LET';
		const depth = this.scope.length;
		const parts: string[] = [];
		for (let k = 0; k < args.length; k++) {
			const arg = args[k];
			const isParam =
				k < last && (!isLet || k % 2 === 0) && arg.type === 'name' && arg.sheet === undefined;
			if (isParam) {
				parts.push(this.param(arg.name));
				if (!isLet) {
					this.bind(arg.name, true);
				}
				continue;
			}
			parts.push(this.expr(arg, 0, true));
			const declared = isLet && k % 2 === 1 ? args[k - 1] : undefined;
			if (declared !== undefined && declared.type === 'name' && declared.sheet === undefined) {
				this.bind(declared.name, arg.type === 'function' && arg.name === 'LAMBDA');
			}
		}
		this.scope.length = depth;
		return parts.join(',');
	}
}
