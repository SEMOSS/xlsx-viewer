import { compareScalars, toText } from '../formula/coerce';
import type { Scalar } from '../formula/types';
import type { CellValue } from '../types/cell';
import type { CfOperator } from '../types/conditional-format';
import { isCellError } from '../values/errors';
import { INVALID_RULE, type CompiledRule, type RuleEnv, type RuleState } from './compiled';
import { isConditionTrue, RelativeFormula } from './formula-eval';
import { duplicateKey } from './statistics';
import { createTextMatcher, isBlankForRule, type TextRuleKind } from './text-match';
import { matchesTimePeriod } from './time-period';

const TEXT_OPERATORS: Partial<Record<CfOperator, TextRuleKind>> = {
	containsText: 'containsText',
	notContains: 'notContainsText',
	beginsWith: 'beginsWith',
	endsWith: 'endsWith',
};

function formulaState(cr: CompiledRule): RuleState {
	if (!cr.state) {
		const { anchorRow, anchorCol } = cr.block;
		cr.state = {
			kind: 'formula',
			formulas: (cr.rule.formulas ?? []).map((f) => new RelativeFormula(f, anchorRow, anchorCol)),
		};
	}
	return cr.state;
}

function formulaAt(
	cr: CompiledRule,
	index: number,
	row: number,
	col: number,
	env: RuleEnv,
): Scalar | undefined {
	const state = formulaState(cr);
	const formula = state.kind === 'formula' ? state.formulas[index] : undefined;
	return formula?.evaluate(env.evaluate, env.sheetIndex, row, col, env.read);
}

/** "Use a formula to determine which cells to format". */
function matchExpression(cr: CompiledRule, row: number, col: number, env: RuleEnv): boolean {
	return isConditionTrue(formulaAt(cr, 0, row, col, env) ?? null);
}

/** "Format only cells with: Cell Value <operator> ...", compared with formula semantics. */
function matchCellIs(
	cr: CompiledRule,
	value: CellValue,
	row: number,
	col: number,
	env: RuleEnv,
): boolean {
	if (isCellError(value)) {
		return false;
	}
	const a = formulaAt(cr, 0, row, col, env);
	if (a === undefined || isCellError(a)) {
		return false;
	}
	const op = cr.rule.operator ?? 'equal';
	if (op === 'between' || op === 'notBetween') {
		const b = formulaAt(cr, 1, row, col, env);
		if (b === undefined || isCellError(b)) {
			return false;
		}
		const [lo, hi] = compareScalars(a, b) <= 0 ? [a, b] : [b, a];
		const inside = compareScalars(value, lo) >= 0 && compareScalars(value, hi) <= 0;
		return op === 'between' ? inside : !inside;
	}
	const textKind = TEXT_OPERATORS[op];
	if (textKind) {
		const text = toText(a);
		return typeof text === 'string' && createTextMatcher(textKind, text)(value);
	}
	const c = compareScalars(value, a);
	switch (op) {
		case 'equal':
			return c === 0;
		case 'notEqual':
			return c !== 0;
		case 'greaterThan':
			return c > 0;
		case 'greaterThanOrEqual':
			return c >= 0;
		case 'lessThan':
			return c < 0;
		case 'lessThanOrEqual':
			return c <= 0;
		default:
			return false;
	}
}

function matchText(
	cr: CompiledRule,
	value: CellValue,
	row: number,
	col: number,
	env: RuleEnv,
): boolean {
	const text = cr.rule.text;
	if (text === undefined) {
		return matchExpression(cr, row, col, env);
	}
	cr.state ??= { kind: 'match', test: createTextMatcher(cr.rule.type as TextRuleKind, text) };
	return cr.state.kind === 'match' && cr.state.test(value);
}

function numericTest(test: (v: number) => boolean): RuleState {
	return { kind: 'match', test: (v) => typeof v === 'number' && test(v) };
}

/** Top/Bottom N (or N%) with ties included, and Above/Below Average (optionally by std devs). */
function statisticState(cr: CompiledRule): RuleState {
	const rule = cr.rule;
	const stats = cr.block.statistics;
	const n = stats.count;
	if (n === 0) {
		return INVALID_RULE;
	}
	if (rule.type === 'top10') {
		let k = Math.floor(rule.rank ?? 10);
		if (!(k > 0)) {
			return INVALID_RULE;
		}
		if (rule.percent) {
			k = Math.max(1, Math.floor((n * k) / 100));
		}
		k = Math.min(k, n);
		if (rule.bottom) {
			const t = stats.small(k)!;
			return numericTest((v) => v <= t);
		}
		const t = stats.large(k)!;
		return numericTest((v) => v >= t);
	}
	const avg = stats.average!;
	const deviations = rule.stdDev ?? 0;
	const above = rule.aboveAverage !== false;
	const t = deviations > 0 ? avg + (above ? 1 : -1) * deviations * (stats.stdDev ?? 0) : avg;
	if (rule.equalAverage) {
		return numericTest(above ? (v) => v >= t : (v) => v <= t);
	}
	return numericTest(above ? (v) => v > t : (v) => v < t);
}

function matchStatistic(cr: CompiledRule, value: CellValue): boolean {
	if (typeof value !== 'number') {
		return false;
	}
	cr.state ??= statisticState(cr);
	return cr.state.kind === 'match' && cr.state.test(value);
}

function matchDuplicates(cr: CompiledRule, value: CellValue): boolean {
	const key = duplicateKey(value);
	if (key === undefined) {
		return false;
	}
	const count = cr.block.statistics.occurrences(key);
	return cr.rule.type === 'duplicateValues' ? count > 1 : count === 1;
}

/**
 * Whether a highlight-type rule (everything except color scales, data bars
 * and icon sets) applies to a cell holding `value`.
 */
export function matchesHighlightRule(
	cr: CompiledRule,
	value: CellValue,
	row: number,
	col: number,
	env: RuleEnv,
): boolean {
	const rule = cr.rule;
	switch (rule.type) {
		case 'cellIs':
			return matchCellIs(cr, value, row, col, env);
		case 'expression':
			return matchExpression(cr, row, col, env);
		case 'containsText':
		case 'notContainsText':
		case 'beginsWith':
		case 'endsWith':
			return matchText(cr, value, row, col, env);
		case 'containsBlanks':
			return isBlankForRule(value) === true;
		case 'notContainsBlanks':
			return isBlankForRule(value) === false;
		case 'containsErrors':
			return isCellError(value);
		case 'notContainsErrors':
			return !isCellError(value);
		case 'timePeriod':
			if (!rule.timePeriod) {
				return matchExpression(cr, row, col, env);
			}
			return (
				typeof value === 'number' &&
				matchesTimePeriod(value, rule.timePeriod, env.today(), env.workbook.date1904)
			);
		case 'top10':
		case 'aboveAverage':
			return matchStatistic(cr, value);
		case 'duplicateValues':
		case 'uniqueValues':
			return matchDuplicates(cr, value);
		default:
			return false;
	}
}
