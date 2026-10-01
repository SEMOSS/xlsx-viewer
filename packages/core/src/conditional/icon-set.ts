import type { CfIconSet, Cfvo } from '../types/conditional-format';
import { ICON_SETS } from './icon-sets';
import type { IconResult } from './types';

/** Number of icons of an icon-set rule (from the set name, else the cfvo count). */
export function iconSetSize(set: CfIconSet): number {
	return ICON_SETS[set.iconSet]?.count || set.cfvos?.length || 0;
}

/** Excel's default thresholds: percent 0/33/67, 0/25/50/75 or 0/20/40/60/80. */
export function defaultIconCfvos(count: number): Cfvo[] {
	return Array.from({ length: count }, (_, i) => ({
		type: 'percent' as const,
		value: String(Math.round((i * 100) / count)),
	}));
}

/**
 * Threshold index for a value: the highest `i >= 1` whose threshold the value
 * reaches (`>=` when `gte[i]`, else `>`), checked from the top like Excel's
 * dialog; 0 when none (the first cfvo is never tested).
 */
export function selectIconThreshold(
	value: number,
	thresholds: readonly number[],
	gte: readonly boolean[],
): number {
	for (let i = thresholds.length - 1; i >= 1; i--) {
		const t = thresholds[i];
		if (gte[i] === false ? value > t : value >= t) {
			return i;
		}
	}
	return 0;
}

/**
 * Icon for a threshold index. Custom icons (Excel 2010+) map threshold
 * positions to explicit icons; otherwise the set's icon is used, reversed
 * when `reverse` is set. `null` means "no cell icon".
 */
export function iconForThreshold(
	set: CfIconSet,
	thresholdIndex: number,
	count: number,
): IconResult | null {
	const custom = set.customIcons?.[thresholdIndex];
	if (custom) {
		const descriptor = ICON_SETS[custom.iconSet];
		if (!descriptor || custom.iconId < 0 || custom.iconId >= descriptor.count) {
			return null;
		}
		return { iconSet: custom.iconSet, index: custom.iconId };
	}
	const descriptor = ICON_SETS[set.iconSet];
	if (!descriptor || descriptor.count === 0) {
		return null;
	}
	const index = set.reverse ? count - 1 - thresholdIndex : thresholdIndex;
	return index >= 0 && index < descriptor.count ? { iconSet: set.iconSet, index } : null;
}
