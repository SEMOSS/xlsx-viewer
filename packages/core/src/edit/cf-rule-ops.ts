/**
 * Conditional Formatting Rules Manager: read a sheet's rules in priority
 * order, describe them the way Excel lists them, and write an edited list
 * back (deleted, reordered, re-targeted or Stop If True changes).
 */
import { rangesEqual } from '../address/range';
import { buildConditionalRuleFormula } from '../conditional/rule-formulas';
import type { RangeAddress } from '../types/address';
import type { CfRule, ConditionalFormat } from '../types/conditional-format';
import type { EditTransaction } from './transaction';
import type { Workbook } from '../types/workbook';

/** One rule as the manager edits it: the rule and the ranges it applies to. */
export interface ManagedRule {
	rule: CfRule;
	ranges: RangeAddress[];
	/** The rule belongs to a PivotTable block (kept in its own block). */
	pivot?: boolean;
}

/** Every rule of a sheet, highest priority first, each with its ranges. */
export function listManagedRules(formats: readonly ConditionalFormat[]): ManagedRule[] {
	const out: Array<ManagedRule & { order: number }> = [];
	for (const format of formats) {
		for (const rule of format.rules) {
			out.push({ rule, ranges: [...format.ranges], pivot: format.pivot, order: out.length });
		}
	}
	out.sort((a, b) => a.rule.priority - b.rule.priority || a.order - b.order);
	return out.map(({ order: _order, ...rest }) => rest);
}

function sameRanges(a: readonly RangeAddress[], b: readonly RangeAddress[]): boolean {
	return a.length === b.length && a.every((r, i) => rangesEqual(r, b[i]));
}

/** Rule types whose stored formula is derived from the rule (and must follow its anchor cell). */
const DERIVED_FORMULA = new Set<CfRule['type']>([
	'containsText',
	'notContainsText',
	'beginsWith',
	'endsWith',
	'containsBlanks',
	'notContainsBlanks',
	'containsErrors',
	'notContainsErrors',
	'timePeriod',
]);

/**
 * Replace a sheet's rules with `rules` (highest priority first). Rules
 * with identical ranges share one block, priorities are renumbered 1..n
 * and formulas Excel derives from the rule (text, blanks, errors, dates)
 * are rebuilt for the rule's anchor cell.
 */
export function replaceConditionalRules(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	rules: readonly ManagedRule[],
): void {
	tx.touchSheet(sheetIndex);
	const sheet = workbook.sheets[sheetIndex];
	const blocks: ConditionalFormat[] = [];
	rules.forEach((managed, i) => {
		if (managed.ranges.length === 0) {
			return;
		}
		let rule: CfRule = { ...managed.rule, priority: i + 1 };
		if (DERIVED_FORMULA.has(rule.type)) {
			const anchor = managed.ranges[0];
			const formula = buildConditionalRuleFormula(rule, anchor.top, anchor.left);
			if (formula) {
				rule = { ...rule, formulas: [formula] };
			}
		}
		let block = blocks.find(
			(b) => Boolean(b.pivot) === Boolean(managed.pivot) && sameRanges(b.ranges, managed.ranges),
		);
		if (!block) {
			block = { ranges: managed.ranges.map((r) => ({ ...r })), rules: [] };
			if (managed.pivot) {
				block.pivot = true;
			}
			blocks.push(block);
		}
		block.rules.push(rule);
	});
	sheet.conditionalFormats = blocks;
	sheet.dirty = true;
}

const OPERATOR_TEXT: Partial<Record<NonNullable<CfRule['operator']>, string>> = {
	greaterThan: '>',
	lessThan: '<',
	greaterThanOrEqual: '>=',
	lessThanOrEqual: '<=',
	equal: '=',
	notEqual: '<>',
};

const PERIOD_TEXT: Record<NonNullable<CfRule['timePeriod']>, string> = {
	today: 'Today',
	yesterday: 'Yesterday',
	tomorrow: 'Tomorrow',
	last7Days: 'Last 7 days',
	thisWeek: 'This week',
	lastWeek: 'Last week',
	nextWeek: 'Next week',
	thisMonth: 'This month',
	lastMonth: 'Last month',
	nextMonth: 'Next month',
};

/** Constants are listed as typed; references and formulas with their `=` (as Excel does). */
function value(formula: string | undefined): string {
	if (formula === undefined) {
		return '';
	}
	const constant =
		/^-?(\d+\.?\d*|\.\d+)(E[+-]?\d+)?$/i.test(formula) ||
		/^"(?:[^"]|"")*"$/.test(formula) ||
		/^(TRUE|FALSE)$/i.test(formula);
	return constant ? formula : `=${formula}`;
}

/** The "Rule (applied in order shown)" text of Excel's Rules Manager. */
export function describeConditionalRule(rule: CfRule): string {
	const f = rule.formulas ?? [];
	switch (rule.type) {
		case 'cellIs': {
			if (rule.operator === 'between' || rule.operator === 'notBetween') {
				const op = rule.operator === 'between' ? 'between' : 'not between';
				return `Cell Value ${op} ${value(f[0])} and ${value(f[1])}`;
			}
			return `Cell Value ${OPERATOR_TEXT[rule.operator ?? 'equal'] ?? '='} ${value(f[0])}`;
		}
		case 'expression':
			return `Formula: ${value(f[0])}`;
		case 'colorScale':
			return 'Graded Color Scale';
		case 'dataBar':
			return 'Data Bar';
		case 'iconSet':
			return 'Icon Set';
		case 'top10':
			return `${rule.bottom ? 'Bottom' : 'Top'} ${rule.rank ?? 10}${rule.percent ? '%' : ''}`;
		case 'duplicateValues':
			return 'Duplicate Values';
		case 'uniqueValues':
			return 'Unique Values';
		case 'containsText':
			return `Cell value contains '${rule.text ?? ''}'`;
		case 'notContainsText':
			return `Cell value does not contain '${rule.text ?? ''}'`;
		case 'beginsWith':
			return `Cell value begins with '${rule.text ?? ''}'`;
		case 'endsWith':
			return `Cell value ends with '${rule.text ?? ''}'`;
		case 'containsBlanks':
			return 'Blanks';
		case 'notContainsBlanks':
			return 'No Blanks';
		case 'containsErrors':
			return 'Errors';
		case 'notContainsErrors':
			return 'No Errors';
		case 'timePeriod':
			return rule.timePeriod ? PERIOD_TEXT[rule.timePeriod] : 'A Date Occurring';
		case 'aboveAverage': {
			const above = rule.aboveAverage !== false;
			if (rule.stdDev) {
				return `${rule.stdDev} std dev ${above ? 'above' : 'below'} average`;
			}
			return `${rule.equalAverage ? 'Equal to or ' : ''}${above ? 'Above' : 'Below'} Average`;
		}
		default:
			return rule.type;
	}
}
