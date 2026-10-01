import { rangesEqual, rangesIntersect } from '../address/range';
import { translateFormula } from '../formula/translate';
import { StylePool } from '../styles/style-pool';
import type { RangeAddress } from '../types/address';
import type { CfRule, ConditionalFormat } from '../types/conditional-format';
import type { Worksheet } from '../types/sheet';
import type { StyleSheet } from '../types/style';
import { sanitizeRange } from './range-cells';
import { buildConditionalRuleFormula } from './rule-formulas';

function toRangeList(ranges: RangeAddress | readonly RangeAddress[]): RangeAddress[] {
	const list = 'top' in ranges ? [ranges] : [...ranges];
	return list.map(sanitizeRange).filter((r) => r !== undefined);
}

function sameRanges(a: readonly RangeAddress[], b: readonly RangeAddress[]): boolean {
	return a.length === b.length && a.every((r, i) => rangesEqual(r, b[i]));
}

/**
 * The parts of `r` not covered by `cut`: up to four rectangles (above,
 * left, right, below). Returns `[r]` when they do not overlap.
 */
export function subtractRange(r: RangeAddress, cut: RangeAddress): RangeAddress[] {
	if (!rangesIntersect(r, cut)) {
		return [r];
	}
	const out: RangeAddress[] = [];
	const top = Math.max(r.top, cut.top);
	const bottom = Math.min(r.bottom, cut.bottom);
	if (cut.top > r.top) {
		out.push({ top: r.top, left: r.left, bottom: cut.top - 1, right: r.right });
	}
	if (cut.left > r.left) {
		out.push({ top, left: r.left, bottom, right: cut.left - 1 });
	}
	if (cut.right < r.right) {
		out.push({ top, left: cut.right + 1, bottom, right: r.right });
	}
	if (cut.bottom < r.bottom) {
		out.push({ top: cut.bottom + 1, left: r.left, bottom: r.bottom, right: r.right });
	}
	return out;
}

/** Renumber priorities 1..n across the sheet, keeping their relative order. */
function renumberPriorities(formats: ConditionalFormat[]): ConditionalFormat[] {
	const all: Array<{ rule: CfRule; order: number }> = [];
	for (const format of formats) {
		for (const rule of format.rules) {
			all.push({ rule, order: all.length });
		}
	}
	all.sort((a, b) => a.rule.priority - b.rule.priority || a.order - b.order);
	const priorities = new Map<CfRule, number>();
	all.forEach((entry, i) => priorities.set(entry.rule, i + 1));
	return formats.map((format) => ({
		...format,
		rules: format.rules.map((rule) => {
			const priority = priorities.get(rule)!;
			return priority === rule.priority ? rule : { ...rule, priority };
		}),
	}));
}

/**
 * Add a rule to a sheet like Excel's Conditional Formatting menu: the new
 * rule gets priority 1 and every existing rule moves down by one. The rule
 * joins an existing block with the same ranges, or starts a new block.
 *
 * - Rules of the "Text that Contains" / blanks / errors / date families get
 *   the formula Excel stores with them when they have none.
 * - With `styles`, an inline `dxf` is interned into `styles.dxfs` and
 *   replaced by `dxfId` (what the file format expects for ordinary rules).
 *
 * The model is updated immutably (new arrays and rule objects) and the sheet
 * is marked dirty. Returns the stored rule. Remember to invalidate a
 * `ConditionalFormatEngine` afterwards.
 */
export function addConditionalFormat(
	sheet: Worksheet,
	ranges: RangeAddress | readonly RangeAddress[],
	rule: CfRule,
	styles?: StyleSheet,
): CfRule {
	const list = toRangeList(ranges);
	if (list.length === 0) {
		throw new RangeError('addConditionalFormat: no valid range');
	}
	let next: CfRule = { ...rule, priority: 1 };
	if (styles && next.dxf && next.dxfId === undefined) {
		const { dxf, ...rest } = next;
		next = { ...rest, dxfId: new StylePool(styles).internDxf(dxf) };
	}
	if (!next.formulas?.length) {
		const formula = buildConditionalRuleFormula(next, list[0].top, list[0].left);
		if (formula) {
			next = { ...next, formulas: [formula] };
		}
	}
	const formats = (sheet.conditionalFormats ?? []).map((format) => ({
		...format,
		rules: format.rules.map((r) => ({ ...r, priority: r.priority + 1 })),
	}));
	const target = formats.find((format) => !format.pivot && sameRanges(format.ranges, list));
	if (target) {
		target.rules = [next, ...target.rules];
	} else {
		formats.push({ ranges: list, rules: [next] });
	}
	sheet.conditionalFormats = formats;
	sheet.dirty = true;
	return next;
}

/**
 * "Clear Rules": without `range`, remove every rule of the sheet ("Clear
 * Rules from Entire Sheet"). With `range` ("from Selected Cells"), blocks
 * whose ranges lie fully inside are removed and overlapping blocks are
 * trimmed; relative formulas of trimmed blocks are re-anchored to their new
 * top-left cell so they keep applying to the same cells. Priorities are
 * renumbered and the sheet is marked dirty. Returns whether anything changed.
 */
export function clearConditionalFormats(
	sheet: Worksheet,
	range?: RangeAddress | readonly RangeAddress[],
): boolean {
	const formats = sheet.conditionalFormats ?? [];
	if (formats.length === 0) {
		return false;
	}
	if (range === undefined) {
		sheet.conditionalFormats = [];
		sheet.dirty = true;
		return true;
	}
	const cuts = toRangeList(range);
	let changed = false;
	const kept: ConditionalFormat[] = [];
	for (const format of formats) {
		const touched = format.ranges.some((r) => cuts.some((cut) => rangesIntersect(r, cut)));
		if (!touched) {
			kept.push(format);
			continue;
		}
		changed = true;
		let remaining: RangeAddress[] = [...format.ranges];
		for (const cut of cuts) {
			remaining = remaining.flatMap((r) => subtractRange(r, cut));
		}
		if (remaining.length === 0) {
			continue;
		}
		const dr = remaining[0].top - format.ranges[0].top;
		const dc = remaining[0].left - format.ranges[0].left;
		const rules =
			dr === 0 && dc === 0
				? format.rules
				: format.rules.map((rule) =>
						rule.formulas?.length
							? { ...rule, formulas: rule.formulas.map((f) => translateFormula(f, dr, dc)) }
							: rule,
					);
		kept.push({ ...format, ranges: remaining, rules });
	}
	if (!changed) {
		return false;
	}
	sheet.conditionalFormats = renumberPriorities(kept);
	sheet.dirty = true;
	return true;
}

/** Every rule of a sheet with its block, in evaluation (priority) order. */
export function getConditionalFormatRules(
	sheet: Worksheet,
): Array<{ rule: CfRule; format: ConditionalFormat }> {
	const out: Array<{ rule: CfRule; format: ConditionalFormat; order: number }> = [];
	for (const format of sheet.conditionalFormats ?? []) {
		for (const rule of format.rules) {
			out.push({ rule, format, order: out.length });
		}
	}
	out.sort((a, b) => a.rule.priority - b.rule.priority || a.order - b.order);
	return out.map(({ rule, format }) => ({ rule, format }));
}
