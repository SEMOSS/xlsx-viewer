import type { Cfvo } from '../types/conditional-format';
import { isNumericLiteral, stripLeadingEquals } from './formula-eval';
import type { RangeStatistics } from './statistics';

/** Evaluates a threshold formula (at the rule's anchor cell) to a number. */
export type ThresholdEvaluator = (formula: string) => number | undefined;

/** Number from a cfvo `val`: a numeric literal, or a formula evaluated through `evaluate`. */
export function cfvoNumber(
	value: string | undefined,
	evaluate: ThresholdEvaluator,
): number | undefined {
	if (value === undefined) {
		return undefined;
	}
	const text = value.trim();
	if (text === '') {
		return undefined;
	}
	if (isNumericLiteral(text)) {
		const n = Number(text);
		return Number.isFinite(n) ? n : undefined;
	}
	return evaluate(stripLeadingEquals(text));
}

/**
 * Resolve a conditional formatting value object (color scale, data bar or
 * icon set threshold) to a number:
 *
 * - `min` / `max`: lowest / highest number in the ranges;
 * - `autoMin` / `autoMax` (data bars, Excel 2010+): `min(0, lowest)` / `max(0, highest)`;
 * - `num` / `formula`: the literal or the formula's value;
 * - `percent`: `min + (max - min) * p / 100`;
 * - `percentile`: PERCENTILE.INC of the numbers at `p / 100`.
 *
 * Returns `undefined` when the threshold cannot be computed (no numbers in
 * the ranges, invalid formula result), which disables the rule.
 */
export function resolveCfvo(
	cfvo: Cfvo | undefined,
	stats: RangeStatistics,
	evaluate: ThresholdEvaluator,
): number | undefined {
	if (!cfvo) {
		return undefined;
	}
	switch (cfvo.type) {
		case 'min':
			return stats.min;
		case 'max':
			return stats.max;
		case 'autoMin': {
			const min = stats.min;
			return min === undefined ? undefined : Math.min(0, min);
		}
		case 'autoMax': {
			const max = stats.max;
			return max === undefined ? undefined : Math.max(0, max);
		}
		case 'num':
		case 'formula':
			return cfvoNumber(cfvo.value, evaluate);
		case 'percent': {
			const p = cfvoNumber(cfvo.value, evaluate);
			const min = stats.min;
			const max = stats.max;
			if (p === undefined || min === undefined || max === undefined) {
				return undefined;
			}
			return min + ((max - min) * p) / 100;
		}
		case 'percentile': {
			const p = cfvoNumber(cfvo.value, evaluate);
			return p === undefined ? undefined : stats.percentile(p / 100);
		}
		default:
			return undefined;
	}
}
