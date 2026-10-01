import type { RangeAddress } from '../../types/address';
import { refAreaToRange } from './references';
import { isLocalReference } from './ref-edits';
import { tokenizeFormula } from './tokenizer';
import type { FormulaToken, FormulaTokenType } from './tokens';

/** Token classes for syntax highlighting in the formula bar and cell editor. */
export type FormulaDisplayTokenKind =
	| 'operator'
	| 'function'
	| 'reference'
	| 'name'
	| 'structured'
	| 'number'
	| 'string'
	| 'boolean'
	| 'error'
	| 'paren'
	| 'separator'
	| 'array'
	| 'whitespace'
	| 'invalid';

export interface FormulaDisplayToken {
	kind: FormulaDisplayTokenKind;
	text: string;
	start: number;
	end: number;
	/** Color slot of a `reference` token, matching {@link getFormulaReferences}. */
	colorIndex?: number;
	/** Nesting depth of a parenthesis or array brace (0 = outermost), for bracket matching. */
	depth?: number;
}

/** A cell / range reference found in a formula, for highlighting it on the grid. */
export interface FormulaReference {
	start: number;
	end: number;
	/** Reference text as written, including any sheet prefix. */
	text: string;
	/** Sheet the reference points to (the formula's sheet when unqualified). */
	sheet?: string;
	range: RangeAddress;
	/**
	 * Color slot: references are numbered in order of first appearance and a
	 * repeated reference reuses its slot. Use `colorIndex % palette.length`.
	 */
	colorIndex: number;
	/** Followed by the spill operator (`A1#`): the whole spill range is meant. */
	spill?: boolean;
}

const DISPLAY_KIND: Record<FormulaTokenType, FormulaDisplayTokenKind> = {
	equals: 'operator',
	whitespace: 'whitespace',
	intersect: 'operator',
	number: 'number',
	string: 'string',
	boolean: 'boolean',
	error: 'error',
	ref: 'reference',
	name: 'name',
	function: 'function',
	structured: 'structured',
	operator: 'operator',
	prefix: 'operator',
	percent: 'operator',
	spill: 'operator',
	'paren-open': 'paren',
	'paren-close': 'paren',
	'call-open': 'paren',
	'call-close': 'paren',
	'arg-sep': 'separator',
	union: 'separator',
	'array-open': 'array',
	'array-close': 'array',
	'array-col-sep': 'array',
	'array-row-sep': 'array',
	unknown: 'invalid',
};

/** Assigns stable color slots to references (same sheet + range = same slot). */
class ReferenceColors {
	private readonly slots = new Map<string, number>();

	slot(sheet: string | undefined, range: RangeAddress): number {
		const { top, left, bottom, right } = range;
		const key = `${(sheet ?? '').toLowerCase()}!${top},${left},${bottom},${right}`;
		let slot = this.slots.get(key);
		if (slot === undefined) {
			slot = this.slots.size;
			this.slots.set(key, slot);
		}
		return slot;
	}
}

function resolvedSheet(t: FormulaToken, formulaSheet: string | undefined): string | undefined {
	return t.sheet !== undefined ? t.sheet.sheet : formulaSheet;
}

/**
 * Tokens for syntax highlighting. Every character belongs to exactly one
 * token; references carry the same color slot as {@link getFormulaReferences}
 * and brackets their nesting depth. Never throws, so it suits partial input.
 */
export function tokenizeForDisplay(formula: string, formulaSheet?: string): FormulaDisplayToken[] {
	const colors = new ReferenceColors();
	let depth = 0;
	return tokenizeFormula(formula).map((t) => {
		const token: FormulaDisplayToken = {
			kind: DISPLAY_KIND[t.type],
			text: t.text,
			start: t.start,
			end: t.end,
		};
		switch (t.type) {
			case 'ref':
				if (isLocalReference(t)) {
					token.colorIndex = colors.slot(resolvedSheet(t, formulaSheet), refAreaToRange(t.area));
				}
				break;
			case 'paren-open':
			case 'call-open':
			case 'array-open':
				token.depth = depth++;
				break;
			case 'paren-close':
			case 'call-close':
			case 'array-close':
				depth = Math.max(0, depth - 1);
				token.depth = depth;
				break;
		}
		return token;
	});
}

/**
 * Cell and range references of a formula with their target ranges and
 * color slots, in order of appearance (Excel's reference highlighting).
 * External and 3-D references and `#REF!` are skipped; names and table
 * references need workbook context and are not included.
 */
export function getFormulaReferences(formula: string, formulaSheet?: string): FormulaReference[] {
	const colors = new ReferenceColors();
	const out: FormulaReference[] = [];
	const tokens = tokenizeFormula(formula);
	for (let i = 0; i < tokens.length; i++) {
		const t = tokens[i];
		if (!isLocalReference(t)) {
			continue;
		}
		const range = refAreaToRange(t.area);
		const sheet = resolvedSheet(t, formulaSheet);
		const ref: FormulaReference = {
			start: t.start,
			end: t.end,
			text: t.text,
			range,
			colorIndex: colors.slot(sheet, range),
		};
		if (sheet !== undefined) {
			ref.sheet = sheet;
		}
		if (tokens[i + 1]?.type === 'spill') {
			ref.spill = true;
		}
		out.push(ref);
	}
	return out;
}
