import { cellKey } from '../address/cell-address';
import type { CellValue } from '../types/cell';
import type { CfRule } from '../types/conditional-format';
import type { DifferentialFormat } from '../types/style';
import type { Workbook } from '../types/workbook';
import { localDateToSerial } from '../values/dates';
import { CfBlock, type CompiledRule, type RuleEnv } from './compiled';
import { mergeDifferentialFormats } from './dxf';
import {
	isNumericLiteral,
	parseLiteralFormula,
	stripLeadingEquals,
	workbookCellReader,
	type CellReader,
	type FormulaEvaluator,
} from './formula-eval';
import { matchesHighlightRule } from './highlight-rules';
import { sanitizeRange } from './range-cells';
import type {
	CfCellResult,
	ConditionalFormatEngineOptions,
	DataBarResult,
	IconResult,
} from './types';
import { colorScaleColor, dataBarFor, iconFor } from './visual-rules';

/** Rows per bucket of the lazily built row-band index. */
const BAND_SHIFT = 7;
/** Per-sheet cap of memoised cell results (the cache is reset when exceeded). */
const MAX_CACHED_RESULTS = 100_000;

interface SheetState {
	readonly env: RuleEnv;
	/** Compiled rules sorted by priority (then document order). */
	readonly rules: CompiledRule[];
	/** Rules whose bounding box touches each row band, built on demand. */
	readonly bands: Map<number, CompiledRule[]>;
	/** Memoised results per cell key (`null` = no conditional formatting). */
	readonly results: Map<number, CfCellResult | null>;
	/** Some rule evaluates formulas (results may depend on other sheets). */
	readonly formulaDependent: boolean;
}

function isVisualRule(rule: CfRule): boolean {
	return rule.type === 'colorScale' || rule.type === 'dataBar' || rule.type === 'iconSet';
}

/** Whether evaluating the rule calls the formula evaluator (results may depend on other sheets). */
function usesFormulas(rule: CfRule): boolean {
	const hasFormula = !!rule.formulas?.length;
	switch (rule.type) {
		case 'expression':
			return true;
		case 'cellIs':
			return (rule.formulas ?? []).some(
				(f) => parseLiteralFormula(stripLeadingEquals(f)) === undefined,
			);
		case 'containsText':
		case 'notContainsText':
		case 'beginsWith':
		case 'endsWith':
			return rule.text === undefined && hasFormula;
		case 'timePeriod':
			return rule.timePeriod === undefined && hasFormula;
		case 'colorScale':
		case 'dataBar':
		case 'iconSet': {
			const cfvos = rule.colorScale?.cfvos ?? rule.dataBar?.cfvos ?? rule.iconSet?.cfvos ?? [];
			return cfvos.some(
				(c) =>
					c.type === 'formula' ||
					((c.type === 'num' || c.type === 'percent' || c.type === 'percentile') &&
						c.value !== undefined &&
						!isNumericLiteral(c.value)),
			);
		}
		default:
			return false;
	}
}

/**
 * Evaluates conditional formatting for rendering.
 *
 * Results are computed lazily per cell and memoised, and the statistics a
 * rule needs (sorted numbers, percentiles, averages, duplicate counts) are
 * computed once per rule block, so asking for every visible cell on each
 * frame is cheap. Nothing is recomputed automatically: call
 * {@link invalidateSheet} after cell values (or the sheet's rules) change and
 * {@link invalidate} after structural changes (sheets added, removed or
 * reordered, styles replaced).
 */
export class ConditionalFormatEngine {
	private readonly states = new Map<number, SheetState | null>();
	private readonly today: () => number;
	private readonly read: CellReader;

	constructor(
		readonly workbook: Workbook,
		private readonly evaluate: FormulaEvaluator,
		options: ConditionalFormatEngineOptions = {},
	) {
		this.read = workbookCellReader(workbook);
		this.today =
			options.today ?? (() => Math.floor(localDateToSerial(new Date(), workbook.date1904)));
	}

	/** Whether the sheet has any conditional formatting rule. */
	hasRules(sheetIndex: number): boolean {
		return (this.state(sheetIndex)?.rules.length ?? 0) > 0;
	}

	/** Conditional formatting of one cell, or `undefined` when no rule applies. */
	getCellResult(sheetIndex: number, row: number, col: number): CfCellResult | undefined {
		const state = this.state(sheetIndex);
		if (!state || state.rules.length === 0) {
			return undefined;
		}
		const key = cellKey(row, col);
		const cached = state.results.get(key);
		if (cached !== undefined) {
			return cached ?? undefined;
		}
		const result = this.compute(state, row, col);
		if (state.results.size >= MAX_CACHED_RESULTS) {
			state.results.clear();
		}
		state.results.set(key, result ?? null);
		return result;
	}

	/** Drop every cached result and statistic. */
	invalidate(): void {
		this.states.clear();
	}

	/**
	 * Drop the caches of one sheet (its values or rules changed). Memoised
	 * results of other sheets whose rules evaluate formulas are dropped too,
	 * since those formulas may reference the changed sheet.
	 */
	invalidateSheet(sheetIndex: number): void {
		this.states.delete(sheetIndex);
		for (const state of this.states.values()) {
			if (state?.formulaDependent) {
				state.results.clear();
				for (const rule of state.rules) {
					rule.state = undefined;
				}
			}
		}
	}

	private state(sheetIndex: number): SheetState | undefined {
		let state = this.states.get(sheetIndex);
		if (state === undefined) {
			state = this.compile(sheetIndex);
			this.states.set(sheetIndex, state);
		}
		return state ?? undefined;
	}

	private compile(sheetIndex: number): SheetState | null {
		const workbook = this.workbook;
		const sheet = workbook.sheets[sheetIndex];
		if (!sheet) {
			return null;
		}
		const rules: CompiledRule[] = [];
		let order = 0;
		for (const format of sheet.conditionalFormats ?? []) {
			const ranges = (format.ranges ?? []).map(sanitizeRange).filter((r) => r !== undefined);
			if (ranges.length === 0 || !format.rules?.length) {
				continue;
			}
			const block = new CfBlock(sheet, ranges);
			for (const rule of format.rules) {
				const dxf =
					rule.dxf ?? (rule.dxfId !== undefined ? workbook.styles.dxfs[rule.dxfId] : undefined);
				rules.push({ rule, block, dxf, order: order++ });
			}
		}
		const priority = (r: CompiledRule) =>
			Number.isFinite(r.rule.priority) ? r.rule.priority : Number.MAX_SAFE_INTEGER;
		rules.sort((a, b) => priority(a) - priority(b) || a.order - b.order);
		const env: RuleEnv = {
			workbook,
			sheet,
			sheetIndex,
			evaluate: this.evaluate,
			read: this.read,
			today: this.today,
			colorCtx: { theme: workbook.theme, indexedColors: workbook.styles.indexedColors },
		};
		return {
			env,
			rules,
			bands: new Map(),
			results: new Map(),
			formulaDependent: rules.some((r) => usesFormulas(r.rule)),
		};
	}

	private bandRules(state: SheetState, row: number): CompiledRule[] {
		const band = row >> BAND_SHIFT;
		let list = state.bands.get(band);
		if (!list) {
			const top = band << BAND_SHIFT;
			const bottom = top + (1 << BAND_SHIFT) - 1;
			list = state.rules.filter((r) => r.block.bbox.top <= bottom && r.block.bbox.bottom >= top);
			state.bands.set(band, list);
		}
		return list;
	}

	private compute(state: SheetState, row: number, col: number): CfCellResult | undefined {
		const env = state.env;
		let value: CellValue = null;
		let loaded = false;
		let style: DifferentialFormat | undefined;
		let colorScale: string | undefined;
		let dataBar: DataBarResult | undefined;
		let icon: IconResult | undefined;
		let iconDecided = false;
		let hideValue = false;
		let fillTaken = false;

		for (const cr of this.bandRules(state, row)) {
			if (!cr.block.contains(row, col)) {
				continue;
			}
			if (!loaded) {
				value = this.read(env.sheetIndex, row, col);
				loaded = true;
			}
			const rule = cr.rule;
			if (isVisualRule(rule)) {
				if (rule.type === 'colorScale') {
					if (!fillTaken) {
						const color = colorScaleColor(cr, value, env);
						if (color !== undefined) {
							colorScale = color;
							fillTaken = true;
						}
					}
				} else if (rule.type === 'dataBar') {
					if (!dataBar) {
						dataBar = dataBarFor(cr, value, env);
						if (dataBar && rule.dataBar?.showValue === false) {
							hideValue = true;
						}
					}
				} else if (!iconDecided) {
					const outcome = iconFor(cr, value, env);
					if (outcome) {
						iconDecided = true;
						icon = outcome.icon;
						if (!outcome.showValue) {
							hideValue = true;
						}
					}
				}
				continue;
			}
			if (!matchesHighlightRule(cr, value, row, col, env)) {
				continue;
			}
			if (cr.dxf) {
				style = mergeDifferentialFormats(style, cr.dxf, fillTaken);
				if (style.fill) {
					fillTaken = true;
				}
			}
			if (rule.stopIfTrue) {
				break;
			}
		}

		if (!style && colorScale === undefined && !dataBar && !icon && !hideValue) {
			return undefined;
		}
		const result: CfCellResult = {};
		if (style) {
			result.style = style;
		}
		if (colorScale !== undefined) {
			result.colorScale = colorScale;
		}
		if (dataBar) {
			result.dataBar = dataBar;
		}
		if (icon) {
			result.icon = icon;
		}
		if (hideValue) {
			result.hideValue = true;
		}
		return result;
	}
}
