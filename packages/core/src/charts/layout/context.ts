/**
 * Shared layout state plus the mapping from DrawingML formatting (fills,
 * outlines, text properties) to render-tree paint, with Excel's defaults for
 * everything the chart leaves automatic.
 */
import { resolveDrawingColor, resolveDrawingColorCss } from '../../color/resolve-color';
import { toCss } from '../../color/color-math';
import { formatValue } from '../../numfmt';
import type { ChartModel } from '../../types/chart';
import type { DrawingColor, DrawingTextStyle, ShapeFill, ShapeLine } from '../../types/drawing';
import type { Theme } from '../../types/theme';
import type { ChartData } from '../data/types';
import { autoChartColor, chartStyleIndex } from './palette';
import { type ChartFont, type ChartTextMeasurer, defaultMeasureText, lineHeightFor } from './text';
import type { ChartGradient, ChartPaint, ChartTextNode } from './types';

/** Pixels per point at 96 DPI. */
export const PX_PER_PT = 96 / 72;
/** EMU per pixel at 96 DPI. */
export const EMU_PER_PX = 9525;

/** Excel's automatic text color (Text 1, 65% luminance): #595959. */
export const DEFAULT_TEXT_COLOR = '#595959';
export const DEFAULT_GRIDLINE_COLOR = '#D9D9D9';
export const DEFAULT_MINOR_GRIDLINE_COLOR = '#F2F2F2';
export const DEFAULT_AXIS_COLOR = '#BFBFBF';

const FONT_FALLBACKS = "Calibri, Carlito, 'Segoe UI', Arial, sans-serif";

export interface ResolvedText {
	/** First family, used for measuring. */
	family: string;
	/** CSS font-family stack. */
	cssFamily: string;
	sizePx: number;
	bold: boolean;
	italic: boolean;
	underline: boolean;
	strike: boolean;
	color: string;
	font: ChartFont;
	lineHeight: number;
}

export interface LayoutContext {
	chart: ChartModel;
	data: ChartData;
	theme: Theme;
	measure: ChartTextMeasurer;
	date1904: boolean;
	/** Gallery style index (1-48). */
	style: number;
}

export function createLayoutContext(
	chart: ChartModel,
	data: ChartData,
	theme: Theme,
	measure?: ChartTextMeasurer,
): LayoutContext {
	return {
		chart,
		data,
		theme,
		measure: measure ?? defaultMeasureText,
		date1904: Boolean(data.date1904),
		style: chartStyleIndex(chart.style),
	};
}

export function colorCss(
	ctx: LayoutContext,
	color: DrawingColor | undefined,
	fallback: string,
): string {
	return resolveDrawingColorCss(color, ctx.theme) ?? fallback;
}

function themeFont(ctx: LayoutContext, name: string | undefined): string {
	if (!name || name === '+mn-lt' || name === '+mn-ea' || name === '+mn-cs') {
		return ctx.theme.fonts.minorLatin || 'Calibri';
	}
	if (name === '+mj-lt' || name === '+mj-ea' || name === '+mj-cs') {
		return ctx.theme.fonts.majorLatin || 'Calibri Light';
	}
	return name;
}

/**
 * Resolve text formatting: `defaults`, then the chart-wide text style, then
 * each style in `styles` (most specific last).
 */
export function resolveText(
	ctx: LayoutContext,
	defaults: { sizePt: number; color?: string; bold?: boolean; relative?: number },
	...styles: Array<DrawingTextStyle | undefined>
): ResolvedText {
	const merged: DrawingTextStyle = { ...(ctx.chart.textStyle ?? {}) };
	for (const s of styles) {
		if (s) {
			Object.assign(
				merged,
				Object.fromEntries(Object.entries(s).filter(([, v]) => v !== undefined)),
			);
		}
	}
	const family = themeFont(ctx, merged.font);
	// An element without its own size scales the chart-wide size (titles are larger).
	let sizePt = defaults.sizePt;
	const own = styles.reduce<number | undefined>((acc, s) => s?.size ?? acc, undefined);
	if (own !== undefined) {
		sizePt = own;
	} else if (ctx.chart.textStyle?.size !== undefined) {
		sizePt = ctx.chart.textStyle.size * (defaults.relative ?? 1);
	}
	const sizePx = Math.max(1, sizePt * PX_PER_PT);
	const bold = merged.bold ?? defaults.bold ?? false;
	const italic = merged.italic ?? false;
	return {
		family,
		cssFamily: `'${family.replace(/'/g, '')}', ${FONT_FALLBACKS}`,
		sizePx,
		bold,
		italic,
		underline: merged.underline ?? false,
		strike: merged.strike ?? false,
		color: merged.color
			? colorCss(ctx, merged.color, defaults.color ?? DEFAULT_TEXT_COLOR)
			: (defaults.color ?? DEFAULT_TEXT_COLOR),
		font: { family, size: sizePx, bold, italic },
		lineHeight: lineHeightFor(sizePx),
	};
}

/** Text node attributes for a resolved style. */
export function textAttrs(
	t: ResolvedText,
): Pick<
	ChartTextNode,
	'fontSize' | 'fontFamily' | 'fontWeight' | 'fontStyle' | 'decoration' | 'fill' | 'lineHeight'
> {
	const attrs: Pick<
		ChartTextNode,
		'fontSize' | 'fontFamily' | 'fontWeight' | 'fontStyle' | 'decoration' | 'fill' | 'lineHeight'
	> = { fontSize: t.sizePx, fontFamily: t.cssFamily, fill: t.color, lineHeight: t.lineHeight };
	if (t.bold) {
		attrs.fontWeight = 'bold';
	}
	if (t.italic) {
		attrs.fontStyle = 'italic';
	}
	if (t.underline) {
		attrs.decoration = 'underline';
	} else if (t.strike) {
		attrs.decoration = 'line-through';
	}
	return attrs;
}

/** Automatic color (CSS) for series / point `index` of `count`. */
export function autoColor(ctx: LayoutContext, index: number, count: number): string {
	return colorCss(ctx, autoChartColor(index, count, ctx.style), '#4472C4');
}

/** Pattern presets approximated by blending foreground over background. */
function patternDensity(preset: string): number {
	const pct = /^pct(\d+)$/.exec(preset);
	if (pct) {
		return Number(pct[1]) / 100;
	}
	if (/^(dk|wd)/.test(preset)) {
		return 0.6;
	}
	if (/^(lt|sm|narrow|dash)/.test(preset)) {
		return 0.3;
	}
	return 0.5;
}

function gradientOf(
	ctx: LayoutContext,
	fill: Extract<ShapeFill, { type: 'gradient' }>,
): ChartGradient {
	return {
		kind: fill.path ? 'radial' : 'linear',
		angle: fill.angle ?? 0,
		stops: fill.stops.map((s) => ({
			offset: s.pos > 1 ? s.pos / 100000 : s.pos,
			color: colorCss(ctx, s.color, '#FFFFFF'),
		})),
	};
}

/**
 * Fill paint. `auto` is the CSS color used when the fill is automatic
 * (undefined) or cannot be drawn (pictures); `null` means no automatic fill.
 */
export function fillPaint(
	ctx: LayoutContext,
	fill: ShapeFill | undefined,
	auto: string | null,
): Pick<ChartPaint, 'fill' | 'gradient'> {
	if (!fill || fill.type === 'picture') {
		return auto ? { fill: auto } : {};
	}
	switch (fill.type) {
		case 'none':
			return {};
		case 'solid':
			return { fill: colorCss(ctx, fill.color, auto ?? '#FFFFFF') };
		case 'gradient': {
			const g = gradientOf(ctx, fill);
			return g.stops.length ? { gradient: g, fill: g.stops[0].color } : auto ? { fill: auto } : {};
		}
		case 'pattern': {
			const fg = resolveDrawingColor(fill.fg, ctx.theme) ?? { r: 0, g: 0, b: 0, a: 1 };
			const bg = resolveDrawingColor(fill.bg, ctx.theme) ?? { r: 1, g: 1, b: 1, a: 1 };
			const d = patternDensity(fill.preset);
			const mix = (a: number, b: number) => a * d + b * (1 - d);
			return { fill: toCss({ r: mix(fg.r, bg.r), g: mix(fg.g, bg.g), b: mix(fg.b, bg.b), a: 1 }) };
		}
	}
}

/** Dash patterns in multiples of the line width (DrawingML preset dashes). */
const DASHES: Record<string, number[]> = {
	dot: [1, 3],
	dash: [4, 3],
	lgDash: [8, 3],
	dashDot: [4, 3, 1, 3],
	lgDashDot: [8, 3, 1, 3],
	lgDashDotDot: [8, 3, 1, 3, 1, 3],
	sysDash: [3, 1],
	sysDot: [1, 1],
	sysDashDot: [3, 1, 1, 1],
	sysDashDotDot: [3, 1, 1, 1, 1, 1],
};

export interface StrokeDefaults {
	color: string;
	/** Width in pixels. */
	width: number;
	cap?: ChartPaint['lineCap'];
}

/**
 * Stroke paint for an outline. `auto` supplies the automatic color and width;
 * `null` means the element has no outline unless one is specified.
 */
export function strokePaint(
	ctx: LayoutContext,
	line: ShapeLine | undefined,
	auto: StrokeDefaults | null,
): Pick<ChartPaint, 'stroke' | 'strokeWidth' | 'dash' | 'lineCap' | 'lineJoin'> {
	if (line?.fill?.type === 'none') {
		return {};
	}
	if (!line && !auto) {
		return {};
	}
	const width =
		line?.width !== undefined ? Math.max(0.5, line.width / EMU_PER_PX) : (auto?.width ?? 1);
	let stroke: string | undefined;
	if (line?.fill) {
		const p = fillPaint(ctx, line.fill, auto?.color ?? '#000000');
		stroke = p.fill;
	} else {
		stroke = auto?.color ?? '#000000';
	}
	const out: Pick<ChartPaint, 'stroke' | 'strokeWidth' | 'dash' | 'lineCap' | 'lineJoin'> = {
		stroke,
		strokeWidth: width,
	};
	const dash = line?.dash ? DASHES[line.dash] : undefined;
	if (dash) {
		out.dash = dash.map((d) => d * width);
	}
	const cap = line?.cap ?? (auto?.cap === 'round' ? 'rnd' : undefined);
	if (cap === 'rnd') {
		out.lineCap = 'round';
		out.lineJoin = 'round';
	} else if (cap === 'sq') {
		out.lineCap = 'square';
	}
	return out;
}

/** Format a number with an Excel format code (axis ticks, data labels). */
export function formatNumber(ctx: LayoutContext, value: number, code: string | undefined): string {
	return formatValue(value, code || 'General', { date1904: ctx.date1904 }).text;
}

/** Format a number and report the format's color section (`[Red]`), as Excel applies it to labels. */
export function formatNumberColored(
	ctx: LayoutContext,
	value: number,
	code: string | undefined,
): { text: string; color?: string } {
	const out = formatValue(value, code || 'General', { date1904: ctx.date1904 });
	return out.color ? { text: out.text, color: out.color } : { text: out.text };
}
