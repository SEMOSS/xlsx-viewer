import { resolveColorCss, type ColorContext } from '../color/resolve-color';
import type { CfDataBar } from '../types/conditional-format';
import type { DataBarResult } from './types';

/** Where the axis between negative and positive bars sits. */
export type DataBarAxisMode = 'none' | 'automatic' | 'middle';

export interface DataBarLayoutInput {
	value: number;
	/** Resolved minimum / maximum thresholds (cfvos). */
	lo: number;
	hi: number;
	/** Shortest / longest bar as fractions of the available width (0..1). */
	minLength: number;
	maxLength: number;
	axis: DataBarAxisMode;
}

/** Bar edges as fractions of the cell width, laid out left to right. */
export interface DataBarLayout {
	start: number;
	end: number;
	axis?: number;
}

function clamp01(v: number): number {
	return v <= 0 ? 0 : v >= 1 ? 1 : v;
}

/**
 * Data bar geometry for a value (left-to-right; see {@link mirrorDataBarLayout}).
 *
 * Without an axis (Excel 2007 bars, `axisPosition="none"`, or no negative
 * range) the bar starts at the left edge: `lo` maps to `minLength`, `hi` to
 * `maxLength`, values in between linearly, values outside clamp.
 *
 * With an automatic axis (Excel 2010+, `lo < 0`) the axis sits where zero
 * falls between `lo` and `hi` (at the right edge when everything is
 * negative): positive bars grow right from it scaled to `hi`, negative bars
 * grow left scaled to `lo`. A middle axis sits at 50% and both sides share
 * the scale of the larger magnitude. Zero has no bar on an axis.
 */
export function layoutDataBar(input: DataBarLayoutInput): DataBarLayout {
	const { value: v, lo, hi } = input;
	const minL = clamp01(input.minLength);
	const maxL = Math.max(minL, clamp01(input.maxLength));
	const sideLength = (f: number) => (f <= 0 ? 0 : minL + Math.min(1, f) * (maxL - minL));
	const useAxis = input.axis === 'middle' || (input.axis === 'automatic' && lo < 0);
	if (!useAxis) {
		let len: number;
		if (v <= lo) {
			len = minL;
		} else if (v >= hi) {
			len = maxL;
		} else {
			len = minL + ((v - lo) / (hi - lo)) * (maxL - minL);
		}
		return { start: 0, end: len };
	}
	let axis: number;
	let negFraction: number;
	let posFraction: number;
	if (input.axis === 'automatic') {
		axis = hi <= 0 ? 1 : -lo / (hi - lo);
		negFraction = v < 0 ? v / lo : 0;
		posFraction = v > 0 && hi > 0 ? v / hi : 0;
	} else {
		axis = 0.5;
		const m = Math.max(Math.abs(lo), Math.abs(hi));
		const negLimit = Math.min(lo, 0);
		const posLimit = Math.max(hi, 0);
		negFraction = v < 0 && negLimit < 0 && m > 0 ? Math.max(v, negLimit) / -m : 0;
		posFraction = v > 0 && posLimit > 0 && m > 0 ? Math.min(v, posLimit) / m : 0;
	}
	if (v < 0) {
		const len = sideLength(negFraction) * axis;
		return { start: axis - len, end: axis, axis };
	}
	const len = sideLength(posFraction) * (1 - axis);
	return { start: axis, end: axis + len, axis };
}

/** Mirror a left-to-right layout for right-to-left bars. */
export function mirrorDataBarLayout(layout: DataBarLayout): DataBarLayout {
	const out: DataBarLayout = { start: 1 - layout.end, end: 1 - layout.start };
	if (layout.axis !== undefined) {
		out.axis = 1 - layout.axis;
	}
	return out;
}

/**
 * Whether a data bar carries Excel 2010+ (`x14:dataBar`) properties. Plain
 * Excel 2007 bars have no axis or negative colors and always use a gradient.
 */
export function isExtendedDataBar(db: CfDataBar): boolean {
	return (
		db.gradient !== undefined ||
		db.border !== undefined ||
		db.borderColor !== undefined ||
		db.negativeFillColor !== undefined ||
		db.negativeBorderColor !== undefined ||
		db.negativeBarColorSameAsPositive !== undefined ||
		db.axisPosition !== undefined ||
		db.axisColor !== undefined ||
		db.direction !== undefined ||
		(db.cfvos ?? []).some((c) => c.type === 'autoMin' || c.type === 'autoMax')
	);
}

/** Excel's default data bar color (the "Blue Data Bar" preset). */
export const DEFAULT_DATA_BAR_COLOR = '#638EC6';
const DEFAULT_NEGATIVE_COLOR = '#FF0000';
const DEFAULT_AXIS_COLOR = '#000000';

/** Rule-level data bar settings with colors resolved, shared by every cell of the rule. */
export interface DataBarStyle {
	extended: boolean;
	minLength: number;
	maxLength: number;
	axis: DataBarAxisMode;
	rtl: boolean;
	gradient: boolean;
	color: string;
	negativeColor: string;
	borderColor?: string;
	negativeBorderColor?: string;
	axisColor: string;
}

/** Resolve the rule-level settings of a data bar (defaults follow the OOXML schemas). */
export function prepareDataBarStyle(
	db: CfDataBar,
	ctx: ColorContext,
	sheetRtl: boolean,
): DataBarStyle {
	const extended = isExtendedDataBar(db);
	const color = resolveColorCss(db.color, ctx, DEFAULT_DATA_BAR_COLOR) ?? DEFAULT_DATA_BAR_COLOR;
	const negativeColor = db.negativeBarColorSameAsPositive
		? color
		: (resolveColorCss(db.negativeFillColor, ctx, DEFAULT_NEGATIVE_COLOR) ??
			DEFAULT_NEGATIVE_COLOR);
	const style: DataBarStyle = {
		extended,
		minLength: (db.minLength ?? 10) / 100,
		maxLength: (db.maxLength ?? 90) / 100,
		axis: extended ? (db.axisPosition ?? 'automatic') : 'none',
		rtl: db.direction === 'rightToLeft' || (db.direction !== 'leftToRight' && sheetRtl),
		gradient: extended ? db.gradient !== false : true,
		color,
		negativeColor: extended ? negativeColor : color,
		axisColor: resolveColorCss(db.axisColor, ctx, DEFAULT_AXIS_COLOR) ?? DEFAULT_AXIS_COLOR,
	};
	if (extended && db.border) {
		const border = resolveColorCss(db.borderColor, ctx, color) ?? color;
		style.borderColor = border;
		style.negativeBorderColor = resolveColorCss(db.negativeBorderColor, ctx, border) ?? border;
	}
	return style;
}

/** Data bar for one numeric cell value, given the resolved thresholds. */
export function buildDataBar(
	style: DataBarStyle,
	value: number,
	lo: number,
	hi: number,
): DataBarResult {
	let layout = layoutDataBar({
		value,
		lo,
		hi,
		minLength: style.minLength,
		maxLength: style.maxLength,
		axis: style.axis,
	});
	if (style.rtl) {
		layout = mirrorDataBarLayout(layout);
	}
	const negative = style.extended && value < 0;
	const result: DataBarResult = {
		start: layout.start,
		end: layout.end,
		negative,
		color: negative ? style.negativeColor : style.color,
		gradient: style.gradient,
		direction: style.rtl ? 'rtl' : 'ltr',
	};
	const border = negative ? style.negativeBorderColor : style.borderColor;
	if (border) {
		result.borderColor = border;
	}
	if (layout.axis !== undefined && style.axis !== 'none') {
		result.axis = layout.axis;
		result.axisColor = style.axisColor;
	}
	return result;
}
