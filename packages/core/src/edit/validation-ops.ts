import { rangeIntersection, rangesIntersect } from '../address/range';
import { subtractRange } from '../conditional/edit';
import { sanitizeRange } from '../conditional/range-cells';
import { translateFormula } from '../formula/translate';
import type { CellAddress, RangeAddress } from '../types/address';
import type { DataValidation } from '../types/data-validation';
import type { Worksheet } from '../types/sheet';
import type { Workbook } from '../types/workbook';
import type { EditTransaction } from './transaction';
import { validationUsesOperator } from './validation-criteria';

export * from './validation-criteria';

/** A data validation rule without the cells it applies to. */
export type DataValidationRule = Omit<DataValidation, 'ranges'>;

export interface SetDataValidationOptions {
	/**
	 * Cell the rule's relative references are written for (Excel's dialog
	 * uses the active cell). Defaults to the top-left cell of the first range.
	 */
	origin?: CellAddress;
}

function topLeft(r: RangeAddress): CellAddress {
	return { row: r.top, col: r.left };
}

/** Valid ranges without overlaps: each range minus the ones before it. */
function disjointRanges(ranges: readonly RangeAddress[]): RangeAddress[] {
	const out: RangeAddress[] = [];
	for (const raw of ranges) {
		const r = sanitizeRange(raw);
		if (!r) {
			continue;
		}
		let pieces = [r];
		for (const prev of out) {
			pieces = pieces.flatMap((p) => subtractRange(p, prev));
		}
		out.push(...pieces);
	}
	return out;
}

function subtractAll(
	ranges: readonly RangeAddress[],
	cuts: readonly RangeAddress[],
): RangeAddress[] {
	let remaining = [...ranges];
	for (const cut of cuts) {
		remaining = remaining.flatMap((r) => subtractRange(r, cut));
	}
	return remaining;
}

function settingsOf(dv: DataValidationRule | DataValidation): DataValidationRule {
	const copy: Partial<DataValidation> = { ...dv };
	delete copy.ranges;
	return copy as DataValidationRule;
}

/**
 * Shift relative references so the rule means the same cells when written
 * for `to` instead of `from`.
 */
function reanchor(
	rule: DataValidationRule,
	from: CellAddress,
	to: CellAddress,
): DataValidationRule {
	const dr = to.row - from.row;
	const dc = to.col - from.col;
	if (dr === 0 && dc === 0) {
		return rule;
	}
	const out = { ...rule };
	if (out.formula1 !== undefined) {
		out.formula1 = translateFormula(out.formula1, dr, dc);
	}
	if (out.formula2 !== undefined) {
		out.formula2 = translateFormula(out.formula2, dr, dc);
	}
	return out;
}

/**
 * Data > Data Validation on a set of ranges: the ranges are removed from
 * every validation of the sheet (validations left without cells are
 * dropped; trimmed ones keep meaning the same cells), then, unless `rule` is
 * `null` (Clear All), a validation with `rule` is added for the ranges.
 * Relative references of `rule` are taken as written for `options.origin`.
 */
export function setDataValidation(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	ranges: readonly RangeAddress[],
	rule: DataValidationRule | null,
	options: SetDataValidationOptions = {},
): void {
	const sheet = workbook.sheets[sheetIndex];
	tx.touchSheet(sheetIndex);
	const cuts = disjointRanges(ranges);
	const kept: DataValidation[] = [];
	for (const dv of sheet.dataValidations) {
		if (!dv.ranges.some((r) => cuts.some((cut) => rangesIntersect(r, cut)))) {
			kept.push(dv);
			continue;
		}
		const remaining = subtractAll(dv.ranges, cuts);
		if (remaining.length === 0) {
			continue;
		}
		const moved = reanchor(settingsOf(dv), topLeft(dv.ranges[0]), topLeft(remaining[0]));
		kept.push({ ...moved, ranges: remaining });
	}
	if (rule && cuts.length > 0) {
		const anchor = topLeft(cuts[0]);
		const settings = reanchor(settingsOf(rule), options.origin ?? anchor, anchor);
		kept.push({ ...settings, ranges: cuts });
	}
	sheet.dataValidations = kept;
	sheet.dirty = true;
}

/**
 * Comparable form of a rule's settings (cells excluded, defaults filled in),
 * so validations Excel would treat as the same compare equal.
 */
function signature(dv: DataValidationRule): string {
	const operator = validationUsesOperator(dv.type) ? (dv.operator ?? 'between') : '';
	return JSON.stringify([
		dv.type,
		operator,
		dv.formula1 ?? '',
		dv.formula2 ?? '',
		dv.allowBlank === true,
		dv.type === 'list' ? dv.inCellDropdown !== false : true,
		dv.showInputMessage === true,
		dv.showErrorMessage === true,
		dv.errorStyle ?? 'stop',
		dv.errorTitle ?? '',
		dv.error ?? '',
		dv.promptTitle ?? '',
		dv.prompt ?? '',
	]);
}

/** Whether two rules have the same settings (their cells aside). */
export function sameValidationSettings(a: DataValidationRule, b: DataValidationRule): boolean {
	return signature(a) === signature(b);
}

/**
 * Every range of the sheet whose validation has the same settings as `rule`
 * ("Apply these changes to all other cells with the same settings").
 */
export function sameValidationRanges(sheet: Worksheet, rule: DataValidationRule): RangeAddress[] {
	const sig = signature(rule);
	return sheet.dataValidations.filter((dv) => signature(dv) === sig).flatMap((dv) => dv.ranges);
}

/** What the Data Validation dialog finds in a selection. */
export interface SelectionValidationState {
	/** The validation of the selection, when exactly one (set of settings) applies. */
	rule?: DataValidation;
	/**
	 * Selected cells have different validations (Excel: "The selection
	 * contains more than one type of validation. Erase current settings and
	 * continue?").
	 */
	mixed: boolean;
	/**
	 * Some selected cells have no validation while others have one (Excel:
	 * "The selection contains some cells without Data Validation settings. Do
	 * you want to extend Data Validation to these cells?").
	 */
	partial: boolean;
	/** The parts of the selection covered by `rule` (where to apply when not extending). */
	covered: RangeAddress[];
}

/**
 * Inspect the validations of a selection by range intersections (cost does
 * not depend on the number of cells, so whole columns are fine).
 */
export function selectionValidationState(
	sheet: Worksheet,
	ranges: readonly RangeAddress[],
): SelectionValidationState {
	const selection = disjointRanges(ranges);
	const hits = sheet.dataValidations.filter((dv) =>
		dv.ranges.some((r) => selection.some((s) => rangesIntersect(r, s))),
	);
	if (hits.length === 0) {
		return { mixed: false, partial: false, covered: [] };
	}
	const mixed = new Set(hits.map(signature)).size > 1;
	const partial =
		subtractAll(
			selection,
			hits.flatMap((dv) => dv.ranges),
		).length > 0;
	if (mixed) {
		return { mixed, partial, covered: [] };
	}
	const rule: DataValidation =
		hits.length === 1 ? hits[0] : { ...hits[0], ranges: hits.flatMap((dv) => dv.ranges) };
	const covered: RangeAddress[] = [];
	for (const s of selection) {
		for (const r of rule.ranges) {
			const both = rangeIntersection(s, r);
			if (both) {
				covered.push(both);
			}
		}
	}
	return { rule, mixed, partial, covered };
}
