/** Formula text helpers for the writer. */
import { addFormulaPrefixes } from '../formula/prefixes';
import type { CfRule } from '../types/conditional-format';

/**
 * Formula as stored in the file: future-function prefixes restored and
 * display-only syntax rewritten (`A1#` -> `_xlfn.ANCHORARRAY(A1)`, `@A1` ->
 * `_xlfn.SINGLE(A1)`, `[@Col]` -> `[[#This Row],[Col]]`). Every formula the
 * writer emits goes through here; `addFormulaPrefixes` has its own fast path.
 */
export function fileFormula(formula: string): string {
	return addFormulaPrefixes(formula);
}

function quote(text: string): string {
	return `"${text.replace(/"/g, '""')}"`;
}

function timePeriodFormula(period: CfRule['timePeriod'], c: string): string | undefined {
	switch (period) {
		case 'today':
			return `FLOOR(${c},1)=TODAY()`;
		case 'yesterday':
			return `FLOOR(${c},1)=TODAY()-1`;
		case 'tomorrow':
			return `FLOOR(${c},1)=TODAY()+1`;
		case 'last7Days':
			return `AND(TODAY()-FLOOR(${c},1)<=6,FLOOR(${c},1)<=TODAY())`;
		case 'thisMonth':
			return `AND(MONTH(${c})=MONTH(TODAY()),YEAR(${c})=YEAR(TODAY()))`;
		case 'lastMonth':
			return `AND(MONTH(${c})=MONTH(EDATE(TODAY(),0-1)),YEAR(${c})=YEAR(EDATE(TODAY(),0-1)))`;
		case 'nextMonth':
			return `AND(MONTH(${c})=MONTH(EDATE(TODAY(),0+1)),YEAR(${c})=YEAR(EDATE(TODAY(),0+1)))`;
		case 'thisWeek':
			return `AND(TODAY()-ROUNDDOWN(${c},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(${c},0)-TODAY()<=7-WEEKDAY(TODAY()))`;
		case 'lastWeek':
			return `AND(TODAY()-ROUNDDOWN(${c},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(${c},0)<(WEEKDAY(TODAY())+7))`;
		case 'nextWeek':
			return `AND(ROUNDDOWN(${c},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(${c},0)-TODAY()<(15-WEEKDAY(TODAY())))`;
		default:
			return undefined;
	}
}

/**
 * The formula Excel stores for rule types that are evaluated through one
 * (text, blanks, errors, time periods), relative to `cell` (the top-left
 * cell of the first range). Returns `undefined` for other types.
 */
export function defaultCfFormula(rule: CfRule, cell: string): string | undefined {
	const t = quote(rule.text ?? '');
	switch (rule.type) {
		case 'containsText':
			return `NOT(ISERROR(SEARCH(${t},${cell})))`;
		case 'notContainsText':
			return `ISERROR(SEARCH(${t},${cell}))`;
		case 'beginsWith':
			return `LEFT(${cell},LEN(${t}))=${t}`;
		case 'endsWith':
			return `RIGHT(${cell},LEN(${t}))=${t}`;
		case 'containsBlanks':
			return `LEN(TRIM(${cell}))=0`;
		case 'notContainsBlanks':
			return `LEN(TRIM(${cell}))>0`;
		case 'containsErrors':
			return `ISERROR(${cell})`;
		case 'notContainsErrors':
			return `NOT(ISERROR(${cell}))`;
		case 'timePeriod':
			return timePeriodFormula(rule.timePeriod, cell);
		default:
			return undefined;
	}
}
