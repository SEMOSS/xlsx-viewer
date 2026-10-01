import type { IconSetName } from '../types/conditional-format';
import type { DifferentialFormat } from '../types/style';

/**
 * Geometry and colors of a data bar inside one cell.
 *
 * Positions are fractions (0..1) of the cell width measured from the cell's
 * **left** edge, so a renderer can place the bar without knowing the sheet
 * direction: right-to-left bars are already mirrored. The gradient fades from
 * the bar's origin (the axis, or the edge positive bars grow from) towards
 * its value end; the origin is on the right when exactly one of `negative`
 * and `direction === 'rtl'` holds.
 */
export interface DataBarResult {
	/** Left edge of the bar. */
	start: number;
	/** Right edge of the bar. `start === end` is a zero-length bar (only the axis shows). */
	end: number;
	/** The bar represents a negative value (negative colors, grows away from the axis). */
	negative: boolean;
	/** Bar fill as a CSS color. */
	color: string;
	/** Bar outline as a CSS color, when the rule draws a border. */
	borderColor?: string;
	/** Gradient fill (color fading to white) rather than a solid fill. */
	gradient: boolean;
	/** Position of the axis line (fraction of the width from the left edge), when one is drawn. */
	axis?: number;
	/** Axis line color (CSS). */
	axisColor?: string;
	/** Growth direction of positive bars. */
	direction: 'ltr' | 'rtl';
}

/** An icon picked by an icon-set rule. */
export interface IconResult {
	iconSet: IconSetName;
	/** Index within the set; 0 is the icon for the lowest values (see `ICON_SETS`). */
	index: number;
}

/** Everything conditional formatting contributes to one cell. */
export interface CfCellResult {
	/**
	 * Differential format merged from all matching highlight rules in priority
	 * order: the higher-priority rule wins per property (fill, font color,
	 * bold, each border edge, number format, ...).
	 */
	style?: DifferentialFormat;
	/** Color-scale fill as a CSS color. Takes precedence over lower-priority fills. */
	colorScale?: string;
	dataBar?: DataBarResult;
	icon?: IconResult;
	/** "Show bar only" / "Show icon only": the cell text should not be drawn. */
	hideValue?: boolean;
}

export interface ConditionalFormatEngineOptions {
	/**
	 * Today's date as a serial number, used by "A Date Occurring" rules.
	 * Defaults to the local calendar date (in the workbook's date system).
	 */
	today?: () => number;
}
