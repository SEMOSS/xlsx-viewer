import type { ColorSpec, DifferentialFormat, FillSpec } from '../types/style';

function fillMissing<T extends object>(
	target: T | undefined,
	source: T | undefined,
): T | undefined {
	if (!source) {
		return target;
	}
	if (!target) {
		return source;
	}
	const out = { ...target } as Record<string, unknown>;
	for (const [key, value] of Object.entries(source)) {
		if (value !== undefined && out[key] === undefined) {
			out[key] = value;
		}
	}
	return out as T;
}

function withoutFill(dxf: DifferentialFormat): DifferentialFormat {
	const { fill: _fill, ...rest } = dxf;
	return rest;
}

/**
 * Merge a lower-priority differential format under a higher-priority one:
 * properties already set by `base` win, missing ones are taken from `next`
 * (per font attribute, per border edge, whole fill, number format, ...).
 * `skipFill` drops `next.fill` (the cell fill is already taken, e.g. by a
 * higher-priority color scale). Inputs are never mutated; when `base` is
 * undefined `next` may be returned as is.
 */
export function mergeDifferentialFormats(
	base: DifferentialFormat | undefined,
	next: DifferentialFormat,
	skipFill = false,
): DifferentialFormat {
	if (!base) {
		return skipFill && next.fill ? withoutFill(next) : next;
	}
	const out: DifferentialFormat = { ...base };
	const font = fillMissing(base.font, next.font);
	if (font) {
		out.font = font;
	}
	if (!out.numFmt && next.numFmt) {
		out.numFmt = next.numFmt;
	}
	if (!out.fill && next.fill && !skipFill) {
		out.fill = next.fill;
	}
	const border = fillMissing(base.border, next.border);
	if (border) {
		out.border = border;
	}
	const alignment = fillMissing(base.alignment, next.alignment);
	if (alignment) {
		out.alignment = alignment;
	}
	const protection = fillMissing(base.protection, next.protection);
	if (protection) {
		out.protection = protection;
	}
	return out;
}

/**
 * The background color a differential format paints. Excel stores the color
 * of a solid DXF fill in `bgColor` (often without a `patternType`), the
 * opposite of cell fills; this resolves both conventions. `undefined` for
 * no fill, `patternType="none"` and gradient fills.
 */
export function dxfBackgroundColor(fill: FillSpec | undefined): ColorSpec | undefined {
	if (!fill || fill.type !== 'pattern' || fill.patternType === 'none') {
		return undefined;
	}
	return fill.bgColor ?? fill.fgColor;
}
