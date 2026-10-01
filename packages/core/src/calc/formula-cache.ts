import type { AstNode } from '../formula/ast';
import { parseFormula } from '../formula/language/parser';
import { analyzeFormula, type FormulaAnalysis } from './analysis';

/** A formula text parsed once: the AST (null when the text does not parse) and its analysis. */
export interface ParsedFormula {
	readonly text: string;
	readonly ast: AstNode | null;
	readonly analysis: FormulaAnalysis | null;
	/** Parser message when the text does not parse. */
	readonly error?: string;
}

/** Entries kept before the cache is flushed (formulas differ per cell, so it can grow large). */
const MAX_ENTRIES = 400_000;

/**
 * Parse cache keyed by formula text (display form, no leading `=`). ASTs
 * and analyses are immutable and shared by every cell with the same text.
 */
export class FormulaCache {
	private readonly entries = new Map<string, ParsedFormula>();

	get(text: string): ParsedFormula {
		let entry = this.entries.get(text);
		if (entry === undefined) {
			entry = parse(text);
			if (this.entries.size >= MAX_ENTRIES) {
				this.entries.clear();
			}
			this.entries.set(text, entry);
		}
		return entry;
	}

	clear(): void {
		this.entries.clear();
	}
}

function parse(text: string): ParsedFormula {
	try {
		const ast = parseFormula(text);
		return { text, ast, analysis: analyzeFormula(ast) };
	} catch (e) {
		return { text, ast: null, analysis: null, error: e instanceof Error ? e.message : String(e) };
	}
}
