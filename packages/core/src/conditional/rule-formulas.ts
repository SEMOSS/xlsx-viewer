import { a1 } from '../address/cell-address';
import type { CfRule } from '../types/conditional-format';

function quote(text: string): string {
	return `"${text.replace(/"/g, '""')}"`;
}

/**
 * The formula Excel stores alongside rules that are evaluated through a
 * fixed formula ("Text that Contains", blanks, errors, "A Date Occurring"),
 * written relative to the anchor cell (top-left of the first range). Excel
 * expects it in the file; the engine itself evaluates these rules natively.
 * Returns `undefined` for rule types that carry no such formula.
 */
export function buildConditionalRuleFormula(
	rule: CfRule,
	anchorRow: number,
	anchorCol: number,
): string | undefined {
	const c = a1(anchorRow, anchorCol);
	const text = rule.text ?? '';
	switch (rule.type) {
		case 'containsText':
			return `NOT(ISERROR(SEARCH(${quote(text)},${c})))`;
		case 'notContainsText':
			return `ISERROR(SEARCH(${quote(text)},${c}))`;
		case 'beginsWith':
			return `LEFT(${c},LEN(${quote(text)}))=${quote(text)}`;
		case 'endsWith':
			return `RIGHT(${c},LEN(${quote(text)}))=${quote(text)}`;
		case 'containsBlanks':
			return `LEN(TRIM(${c}))=0`;
		case 'notContainsBlanks':
			return `LEN(TRIM(${c}))>0`;
		case 'containsErrors':
			return `ISERROR(${c})`;
		case 'notContainsErrors':
			return `NOT(ISERROR(${c}))`;
		case 'timePeriod':
			return rule.timePeriod ? timePeriodFormula(rule.timePeriod, c) : undefined;
		default:
			return undefined;
	}
}

function timePeriodFormula(period: NonNullable<CfRule['timePeriod']>, c: string): string {
	const day = `FLOOR(${c},1)`;
	const trunc = `ROUNDDOWN(${c},0)`;
	switch (period) {
		case 'today':
			return `${day}=TODAY()`;
		case 'yesterday':
			return `${day}=TODAY()-1`;
		case 'tomorrow':
			return `${day}=TODAY()+1`;
		case 'last7Days':
			return `AND(TODAY()-${day}<=6,${day}<=TODAY())`;
		case 'thisMonth':
			return `AND(MONTH(${c})=MONTH(TODAY()),YEAR(${c})=YEAR(TODAY()))`;
		case 'lastMonth':
			return `AND(MONTH(${c})=MONTH(EDATE(TODAY(),0-1)),YEAR(${c})=YEAR(EDATE(TODAY(),0-1)))`;
		case 'nextMonth':
			return `AND(MONTH(${c})=MONTH(EDATE(TODAY(),0+1)),YEAR(${c})=YEAR(EDATE(TODAY(),0+1)))`;
		case 'thisWeek':
			return `AND(TODAY()-${trunc}<=WEEKDAY(TODAY())-1,${trunc}-TODAY()<=7-WEEKDAY(TODAY()))`;
		case 'lastWeek':
			return `AND(TODAY()-${trunc}>=(WEEKDAY(TODAY())),TODAY()-${trunc}<(WEEKDAY(TODAY())+7))`;
		case 'nextWeek':
			return `AND(${trunc}-TODAY()>(7-WEEKDAY(TODAY())),${trunc}-TODAY()<(15-WEEKDAY(TODAY())))`;
	}
}
