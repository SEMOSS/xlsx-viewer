import type { Cell, ColorContext, ResolvedCellStyle } from 'xlsx-viewer-core';
import { formatValue, isCellError, resolveColorCss } from 'xlsx-viewer-core';

import { fontToCss, type FontCss } from './fonts';

export type HAlign =
	| 'left'
	| 'center'
	| 'right'
	| 'justify'
	| 'fill'
	| 'distributed'
	| 'centerContinuous';
export type VAlign = 'top' | 'center' | 'bottom' | 'justify' | 'distributed';

export interface ContentRun {
	text: string;
	font: FontCss;
	color: string;
}

/** Everything the grid needs to paint a cell's text, independent of position. */
export interface CellContent {
	kind: 'number' | 'text' | 'boolean' | 'error';
	text: string;
	/** Rich text runs (when the value is formatted text). */
	runs?: ContentRun[];
	font: FontCss;
	color: string;
	hAlign: HAlign;
	vAlign: VAlign;
	wrap: boolean;
	shrink: boolean;
	/** Left / right indent in px (already zoomed). */
	indent: number;
	/** CSS rotation in degrees (positive = clockwise), 0 for none. */
	rotation: number;
	/** Stacked vertical text (textRotation 255). */
	vertical: boolean;
	/** Accounting-style repeat fill (`*x` in the format). */
	repeatFill?: { before: string; char: string; after: string };
	/** The value cannot be displayed in its format (negative date): show hashes. */
	invalid: boolean;
	/** Underline uses accounting style (full cell width). */
	accountingUnderline: boolean;
}

export interface DescribeOptions {
	zoom: number;
	colorCtx: ColorContext;
	date1904: boolean;
	/** Text color for automatic font colors. */
	defaultTextColor: string;
	/** Show formulas instead of values (View > Show Formulas). */
	showFormulas?: boolean;
	/** Hide zero values (sheet option). */
	hideZeros?: boolean;
}

/** Excel indent: about 3 characters (10 px at 100%) per level. */
export const INDENT_PX = 9;

function horizontal(
	style: ResolvedCellStyle,
	kind: CellContent['kind'],
	isFormulaText: boolean,
): HAlign {
	const h = style.alignment.horizontal;
	if (!h || h === 'general') {
		if (isFormulaText) {
			return 'left';
		}
		return kind === 'number' ? 'right' : kind === 'text' ? 'left' : 'center';
	}
	return h;
}

/**
 * Describe a cell's text for rendering, or `null` when nothing is drawn
 * (blank cells, hidden zeros).
 */
export function describeCellContent(
	cell: Cell | undefined,
	style: ResolvedCellStyle,
	opts: DescribeOptions,
): CellContent | null {
	if (!cell) {
		return null;
	}
	const showFormula = opts.showFormulas && cell.formula !== undefined;
	const value = cell.value;
	if (!showFormula && (value === null || value === '')) {
		return null;
	}
	if (!showFormula && opts.hideZeros && value === 0) {
		return null;
	}

	const font = fontToCss(style.font, opts.zoom);
	const fontColor =
		resolveColorCss(style.font.color, opts.colorCtx, opts.defaultTextColor) ??
		opts.defaultTextColor;

	let kind: CellContent['kind'];
	let text: string;
	let color = fontColor;
	let repeatFill: CellContent['repeatFill'];
	let invalid = false;
	if (showFormula) {
		kind = 'text';
		text = `=${cell.formula}`;
	} else if (typeof value === 'number') {
		kind = 'number';
		// Formula view shows constants without their number format, like Excel (dates as serials).
		const formatted = formatValue(value, opts.showFormulas ? 'General' : style.numFmt, {
			date1904: opts.date1904,
		});
		text = formatted.text;
		if (formatted.color) {
			color = formatted.color;
		}
		repeatFill = formatted.fill;
		invalid = formatted.invalid === true;
	} else if (typeof value === 'boolean') {
		kind = 'boolean';
		text = value ? 'TRUE' : 'FALSE';
	} else if (isCellError(value)) {
		kind = 'error';
		text = value.code;
	} else {
		kind = 'text';
		const formatted = formatValue(value as string, opts.showFormulas ? 'General' : style.numFmt, {
			date1904: opts.date1904,
		});
		text = formatted.text;
		if (formatted.color) {
			color = formatted.color;
		}
		repeatFill = formatted.fill;
	}

	let runs: ContentRun[] | undefined;
	if (!showFormula && kind === 'text' && cell.richText && cell.richText.length > 0) {
		runs = cell.richText.map((run) => {
			const merged = { ...style.font, ...(run.font ?? {}) };
			return {
				text: run.text,
				font: fontToCss(merged, opts.zoom),
				color:
					resolveColorCss(
						run.font?.color ?? style.font.color,
						opts.colorCtx,
						opts.defaultTextColor,
					) ?? opts.defaultTextColor,
			};
		});
	}

	const a = style.alignment;
	const rot = a.textRotation ?? 0;
	return {
		kind,
		text,
		runs,
		font,
		color,
		hAlign: horizontal(style, kind, Boolean(showFormula)),
		vAlign: a.vertical ?? 'bottom',
		wrap: Boolean(a.wrapText) || a.horizontal === 'justify' || a.horizontal === 'distributed',
		shrink: Boolean(a.shrinkToFit) && !a.wrapText,
		indent: (a.indent ?? 0) * INDENT_PX * opts.zoom,
		rotation: rot === 255 ? 0 : rot > 90 ? rot - 90 : -rot,
		vertical: rot === 255,
		repeatFill,
		invalid,
		accountingUnderline:
			style.font.underline === 'singleAccounting' || style.font.underline === 'doubleAccounting',
	};
}

/** The content can spill into neighbouring empty cells. */
export function canOverflow(content: CellContent): boolean {
	return (
		content.kind === 'text' &&
		!content.wrap &&
		!content.shrink &&
		!content.vertical &&
		content.rotation === 0 &&
		!content.repeatFill &&
		(content.hAlign === 'left' || content.hAlign === 'right' || content.hAlign === 'center')
	);
}
