import type { BorderEdge, DifferentialFormat, FillSpec } from '../types/style';

/** The formats offered by Excel's Highlight Cells Rules / Top Bottom Rules dialogs. */
export type HighlightPresetName =
	| 'lightRedFillDarkRedText'
	| 'yellowFillDarkYellowText'
	| 'greenFillDarkGreenText'
	| 'lightRedFill'
	| 'redText'
	| 'redBorder';

function deepFreeze<T>(value: T): T {
	if (value && typeof value === 'object') {
		for (const child of Object.values(value as Record<string, unknown>)) {
			deepFreeze(child);
		}
		Object.freeze(value);
	}
	return value;
}

/**
 * Solid DXF fill. Excel keeps the color of solid differential fills in
 * `bgColor` while cell fills use `fgColor`; setting both renders correctly
 * under either convention and round-trips the way Excel writes it.
 */
function solidFill(argb: string): FillSpec {
	return { type: 'pattern', patternType: 'solid', fgColor: { rgb: argb }, bgColor: { rgb: argb } };
}

const RED_EDGE: BorderEdge = { style: 'thin', color: { rgb: 'FF9C0006' } };

/** Differential formats of Excel's highlight presets. Frozen: clone before editing. */
export const HIGHLIGHT_PRESETS: Readonly<Record<HighlightPresetName, DifferentialFormat>> =
	deepFreeze({
		lightRedFillDarkRedText: { font: { color: { rgb: 'FF9C0006' } }, fill: solidFill('FFFFC7CE') },
		yellowFillDarkYellowText: { font: { color: { rgb: 'FF9C5700' } }, fill: solidFill('FFFFEB9C') },
		greenFillDarkGreenText: { font: { color: { rgb: 'FF006100' } }, fill: solidFill('FFC6EFCE') },
		lightRedFill: { fill: solidFill('FFFFC7CE') },
		redText: { font: { color: { rgb: 'FF9C0006' } } },
		redBorder: {
			border: { left: RED_EDGE, right: RED_EDGE, top: RED_EDGE, bottom: RED_EDGE },
		},
	});

/** Menu labels of the highlight presets, as shown by Excel. */
export const HIGHLIGHT_PRESET_LABELS: Readonly<Record<HighlightPresetName, string>> = Object.freeze(
	{
		lightRedFillDarkRedText: 'Light Red Fill with Dark Red Text',
		yellowFillDarkYellowText: 'Yellow Fill with Dark Yellow Text',
		greenFillDarkGreenText: 'Green Fill with Dark Green Text',
		lightRedFill: 'Light Red Fill',
		redText: 'Red Text',
		redBorder: 'Red Border',
	},
);

export interface ColorScalePreset {
	/** Gallery name; the first color named is the one for the highest values. */
	name: string;
	/** ARGB colors from the lowest threshold to the highest (2 or 3). */
	colors: readonly string[];
}

const RED = 'FFF8696B';
const YELLOW = 'FFFFEB84';
const GREEN = 'FF63BE7B';
const WHITE = 'FFFCFCFF';
const BLUE = 'FF5A8AC6';
const LIGHT_YELLOW = 'FFFFEF9C';

/** Excel's twelve color scale presets, in gallery order. */
export const COLOR_SCALE_PRESETS: readonly ColorScalePreset[] = deepFreeze([
	{ name: 'Green - Yellow - Red Color Scale', colors: [RED, YELLOW, GREEN] },
	{ name: 'Red - Yellow - Green Color Scale', colors: [GREEN, YELLOW, RED] },
	{ name: 'Green - White - Red Color Scale', colors: [RED, WHITE, GREEN] },
	{ name: 'Red - White - Green Color Scale', colors: [GREEN, WHITE, RED] },
	{ name: 'Blue - White - Red Color Scale', colors: [RED, WHITE, BLUE] },
	{ name: 'Red - White - Blue Color Scale', colors: [BLUE, WHITE, RED] },
	{ name: 'White - Red Color Scale', colors: [RED, WHITE] },
	{ name: 'Red - White Color Scale', colors: [WHITE, RED] },
	{ name: 'Green - White Color Scale', colors: [WHITE, GREEN] },
	{ name: 'White - Green Color Scale', colors: [GREEN, WHITE] },
	{ name: 'Green - Yellow Color Scale', colors: [LIGHT_YELLOW, GREEN] },
	{ name: 'Yellow - Green Color Scale', colors: [GREEN, LIGHT_YELLOW] },
]);

export interface DataBarPreset {
	name: string;
	/** ARGB bar color. */
	color: string;
}

/** Excel's six data bar colors (each offered with gradient and solid fill). */
export const DATA_BAR_PRESETS: readonly DataBarPreset[] = deepFreeze([
	{ name: 'Blue Data Bar', color: 'FF638EC6' },
	{ name: 'Green Data Bar', color: 'FF63C384' },
	{ name: 'Red Data Bar', color: 'FFFF555A' },
	{ name: 'Orange Data Bar', color: 'FFFFB628' },
	{ name: 'Light Blue Data Bar', color: 'FF008AEF' },
	{ name: 'Purple Data Bar', color: 'FFD6007B' },
]);
